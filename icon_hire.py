# -*- coding: utf-8 -*-
"""
ICON 6103 Coastal Schedule Management 로 항차별 Off-hire / Guaranteed 판정.

사용법:
    python icon_hire.py            # 판정이 필요한 항차만 조회
    python icon_hire.py --all      # 대상 항차를 모두 다시 조회
    python icon_hire.py --show     # 브라우저 창을 띄워서 실행

판정 방법
  - VVD 앞 4글자(선박) + 다음 4글자(항차 번호)로 6103 을 조회한다 (마지막 글자 무시).
  - 결과에서 그 VVD 가 처음 나오는 행의 Delay ETD(시간)를 본다. 그 행의 Status 가
    Skip(CSTS='S')이면 기항을 안 해서 Delay 가 비어 있으므로 바로 아래 행을 쓴다.
  - 그 행의 항구가 동남아(뷰어의 '동남아'와 같은 범위)면 72시간, 아니면 48시간 이상이
    Off-hire, 미만이 Guaranteed.

대상 항차: 스케줄 기준 출항일이 전달 ~ 이번 달인 모선. 매월 20일부터는 다음 달도.
한 번 판정한 항차는 출항일(스케줄 ETD)이 바뀌었을 때만 다시 조회한다.

결과: data/hire.json  {VVD: {status, delay, port, etd, checked}}
"""
import argparse
import json
import re
import sys
import traceback
from pathlib import Path

from playwright.sync_api import sync_playwright

from icon_actuals import login
from scraper import load_config, now_kst, log, dump_debug, describe_page

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT = DATA / "hire.json"
API = "https://icon.pcsline.co.kr/vms/sch/management/selectVMSSCH6103R11_NEW"

# 뷰어(index.html)의 동남아 구분과 같다
SEA_COUNTRY = {"VN", "TH", "HK", "SG", "MY", "ID", "PH", "KH", "MM"}
SEA_PORT = {"CNSHK", "CNXMN"}
LIMIT_SEA, LIMIT = 72, 48

FETCH_JS = """async ([url, payload]) => {
  const r = await fetch(url, {
    method: 'POST',
    headers: {'Content-Type': 'application/json; charset=UTF-8'},
    body: JSON.stringify(payload),
    credentials: 'include'
  });
  if (!r.ok) throw new Error('HTTP ' + r.status);
  return await r.text();
}"""


def payload(vsl, voy, today):
    """6103 Inquiry 가 보내는 조회 조건 그대로. 날짜 칸은 항차를 넣으면 결과에 영향이 없다."""
    return {"dma_dsIn": {
        "INPBUKRS": "2000", "INPROT": "", "INPVSL": vsl, "INPVOY": voy, "INPVOY2": "",
        "INPBND": "", "INPBND2": "", "INPVSN": "", "INPTAG": "N",
        "INPFR": today.strftime("%Y%m%d"), "INPTO": today.strftime("%Y%m%d"),
        "INPOWN": "Y", "INPYTAG": "N", "LASTSAVE": "", "INPVER": "", "INPPFM": "",
    }}


num = lambda v: isinstance(v, (int, float))


def pick_row(rows, vvd):
    """판정에 쓸 (행, Delay, 비고). 같은 VVD 의 행만 본다.

    - 첫 행이 Skip(CSTS='S')이면 Skip 이 아닌 행이 나올 때까지 바로 아래 행.
    - 그 행이 Special Calling(CSTS='C')이고 Delay 가 0(또는 빈칸)이면 다음 진행 행까지 보고,
      그 행 Delay 가 0 이 아니면 그 행으로, 아니면 0 으로 판정한다.
    - 항구가 전부 Skip 이면 그 항차는 기항하지 않으므로 판정하지 않는다 ("all-skip").
    """
    mine = [r for r in rows if (r.get("VVDX") or "").strip() == vvd]
    if not mine:
        return None, None, "missing"
    live = [r for r in mine if (r.get("CSTS") or "").strip() != "S"]
    if not live:
        return None, None, "all-skip"
    r, d = live[0], live[0].get("CEDDY")
    note = "skip" if r is not mine[0] else None
    if (r.get("CSTS") or "").strip() == "C" and not (num(d) and d != 0):
        nxt = live[1] if len(live) > 1 else None
        if nxt is not None and num(nxt.get("CEDDY")) and nxt["CEDDY"] != 0:
            return nxt, nxt["CEDDY"], "special"
        return r, 0, "special"
    return r, d, note


def judge(port, delay):
    sea = port[:2] in SEA_COUNTRY or port in SEA_PORT
    return "Off-hire" if delay >= (LIMIT_SEA if sea else LIMIT) else "Guaranteed"


# --- 대상 항차 -------------------------------------------------------------
def shift_ym(ym, k):
    y, m = int(ym[:4]), int(ym[5:7]) - 1 + k
    return "%04d-%02d" % (y + m // 12, m % 12 + 1)


def load_snapshots():
    """[(스냅샷, 쓸 달)] 최신 스냅샷은 전부, 최신에 전달이 없으면 전달이 들어 있던 마지막
    스냅샷에서 전달 모선만. 뷰어와 같은 방식 (그 스냅샷의 이번 달 모선은 이미 지난 정보다)."""
    name_re = re.compile(r"^(\d{4}-\d{2}-\d{2})_(AM|PM)\.json$")
    files = sorted((f for f in DATA.glob("*.json") if name_re.match(f.name)), reverse=True)
    if not files:
        return []
    latest = json.loads(files[0].read_text(encoding="utf-8"))
    pm = shift_ym(latest["date"][:7], -1)
    has_pm = lambda s: any((v.get("etd_date") or "").startswith(pm)
                           for r in s["routes"] for v in r.get("vessels", []))
    snaps = [(latest, None)]
    if not has_pm(latest):
        for f in files[1:]:
            if f.name[:10] < pm + "-01":
                break
            s = json.loads(f.read_text(encoding="utf-8"))
            if has_pm(s):
                snaps.append((s, pm))
                break
    return snaps


def targets(today):
    """{VVD: 가장 이른 ETD}  - 판정 대상 기간에 출항하는 모든 항차 (수입·수출)."""
    cur = today.strftime("%Y-%m")
    lo, hi = shift_ym(cur, -1), (shift_ym(cur, 1) if today.day >= 20 else cur)
    out = {}
    for s, only in load_snapshots():
        for r in s["routes"]:
            for v in r.get("vessels", []):
                vvd, etd = (v.get("vslvoy") or "").strip(), v.get("etd") or ""
                if len(vvd) < 8 or not (lo <= etd[:7] <= hi):
                    continue
                if only and not (v.get("etd_date") or "").startswith(only):
                    continue
                if vvd not in out or etd < out[vvd]:
                    out[vvd] = etd
    return out


# --- 메인 ----------------------------------------------------------------
def run(show=False, redo_all=False):
    cfg = load_config()
    today = now_kst()
    want = targets(today.date())
    old = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}

    # 지금 스케줄에 없는 항차의 기록은 버린다
    old = {v: rec for v, rec in old.items() if v in want}

    # 판정이 없거나(못 했거나) 스케줄 출항일이 바뀐 항차만 다시 본다.
    # 전부 Skip 인 항차는 판정 없이 기록만 해 두고, 출항일이 바뀌기 전에는 다시 보지 않는다.
    def need(v):
        rec = old.get(v)
        if redo_all or not rec or rec.get("etd") != want[v]:
            return True
        return not rec.get("status") and not rec.get("all_skip")
    todo = sorted(v for v in want if need(v))
    log("=" * 60)
    log("Off-hire 판정 시작: 대상 %d항차 중 조회 필요 %d항차" % (len(want), len(todo)))
    if not todo:
        OUT.write_text(json.dumps(old, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
        return 0

    groups = {}
    for v in todo:
        groups.setdefault((v[:4], v[4:8]), []).append(v)

    done = 0
    headless = cfg.get("headless", True) and not show
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        ctx = browser.new_context(viewport={"width": 2300, "height": 1000}, locale="ko-KR")
        page = ctx.new_page()
        page.on("dialog", lambda d: d.accept())
        try:
            login(page, cfg)
            for (vsl, voy), vvds in sorted(groups.items()):
                # 서버가 가끔 503 을 준다. 잠시 쉬었다가 세 번까지 다시 해 본다.
                rows = None
                for attempt in (1, 2, 3):
                    try:
                        txt = page.evaluate(FETCH_JS, [API, payload(vsl, voy, today)])
                        rows = json.loads(txt).get("dlt_dsOut") or []
                        break
                    except Exception as exc:
                        err = exc
                        page.wait_for_timeout(3000 * attempt)
                if rows is None:
                    log("  %s%s 조회 실패: %s (다음에 다시 조회)" % (vsl, voy, err))
                    continue
                for vvd in vvds:
                    row, delay, note = pick_row(rows, vvd)
                    rec = {"status": None, "delay": None, "port": None, "etd": want[vvd],
                           "checked": today.strftime("%Y-%m-%d %H:%M")}
                    if note == "all-skip":
                        rec["all_skip"] = True   # 기항하지 않는 항차 - 표시하지 않음
                        log("  %s: 모든 항구가 Skip - 판정·표시하지 않음" % vvd)
                    elif note == "missing":
                        log("  %s: 6103 결과에 이 VVD 가 없음 (다음에 다시 조회)" % vvd)
                    elif not num(delay):
                        rec["port"] = row.get("PORT")
                        log("  %s: 판정 행(%s)에 Delay ETD 가 비어 있음 (다음에 다시 조회)" % (vvd, rec["port"]))
                    else:
                        rec.update(port=row.get("PORT"), delay=delay,
                                   status=judge(row.get("PORT") or "", delay))
                        if note:
                            rec["note"] = note   # skip: 첫 항구 Skip / special: Special Calling 처리
                        done += 1
                        log("  %s: %s Delay %s -> %s%s" % (vvd, rec["port"], delay, rec["status"],
                            {"skip": " (첫 항구 Skip, 아래 행)", "special": " (Special Calling)"}.get(note, "")))
                    old[vvd] = rec
                page.wait_for_timeout(200)
        except Exception:
            log("  현재 화면: %s" % describe_page(page))
            dump_debug(page, "icon_hire_fatal")
            raise
        finally:
            browser.close()

    OUT.write_text(json.dumps(old, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
    log("판정 저장: data/%s (이번에 %d항차 판정, 전체 %d항차)" % (OUT.name, done, len(old)))
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="ICON 6103 Off-hire / Guaranteed 판정")
    ap.add_argument("--show", action="store_true", help="브라우저 창을 띄워서 실행")
    ap.add_argument("--all", action="store_true", help="대상 항차를 모두 다시 조회")
    args = ap.parse_args()
    try:
        sys.exit(run(args.show, args.all))
    except Exception as e:
        log("Off-hire 판정 치명적 오류: %s" % e)
        traceback.print_exc()
        sys.exit(2)
