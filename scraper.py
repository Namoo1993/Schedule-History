# -*- coding: utf-8 -*-
"""
PCS LINE (ebiz.pcsline.co.kr) 구간별 스케줄 수집기.

사용법:
    python scraper.py                 # 현재 시각 기준으로 오전/오후 자동 판정
    python scraper.py --session AM    # 오전 스냅샷으로 강제 저장
    python scraper.py --session PM    # 오후 스냅샷으로 강제 저장
    python scraper.py --show          # 브라우저 창을 띄워서 동작 확인

결과: data/YYYY-MM-DD_AM.json  /  data/YYYY-MM-DD_PM.json
"""
import argparse
import datetime as dt
import json
import os
import re
import sys
import traceback
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
LOGS = BASE / "logs"
SITE = "https://ebiz.pcsline.co.kr/"

# GitHub Actions 러너는 UTC 로 돌기 때문에 날짜/오전오후 판정은 항상 한국시간 기준으로 한다.
KST = dt.timezone(dt.timedelta(hours=9))

# --- 화면 element ---------------------------------------------------------
LOGIN_BTN = "#mf_wfm_header_loginBtn"
LOGIN_ID = "#mf_wfm_header_UIE0015P1_wframe_ibx_empCd"
LOGIN_PW = "#mf_wfm_header_UIE0015P1_wframe_ibx_pwd"
LOGIN_GO = "#mf_wfm_header_UIE0015P1_wframe_btn_login"
LOGIN_ERR = "#mf_wfm_header_UIE0015P1_wframe_tbx_loginChk"
SCHEDULE_MENU = "#mf_wfm_header_gen_firstGenerator_0_btn_menu1_Label"
SECTION_MENU = "a.menu_key[onclick*='UIE0210']"

P = "#mf_tac_layout_contents_00010003_body_"
POL = P + "ibx_pol"
POD = P + "ibx_pod"
POPUP_INPUT = P + "multiInputPop_wframe_ibx_location_input"
INQUIRY = P + "btn_inquiry"

OPEN_SECTION_JS = (
    "com.openMenu('%EA%B5%AC%EA%B0%84%EB%B3%84','/WS/sch/UIE0210.xml','00010003')"
)

# 같은 세션 쿠키로 다른 달을 추가 조회한다 (Inquiry 버튼이 실제로 쏘는 요청과 동일).
FETCH_MONTH_JS = """async ([payload]) => {
  const r = await fetch('/sch/selectScheList', {
    method: 'POST',
    headers: {'Content-Type': 'application/json; charset=UTF-8'},
    body: JSON.stringify(payload),
    credentials: 'include'
  });
  if (!r.ok) throw new Error('HTTP ' + r.status);
  return await r.text();
}"""


def now_kst():
    return dt.datetime.now(KST)


def log(msg):
    now = now_kst()
    line = "[%s] %s" % (now.strftime("%H:%M:%S"), msg)
    print(line, flush=True)
    LOGS.mkdir(exist_ok=True)
    with open(LOGS / ("scrape_%s.log" % now.strftime("%Y-%m")), "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_config():
    """config.json 우선, 없으면 config.example.json (GitHub Actions 는 후자를 쓴다).

    아이디/비밀번호는 환경변수 PCS_ID / PCS_PW 가 있으면 그쪽이 우선한다.
    """
    cfg_path = BASE / "config.json"
    if not cfg_path.exists():
        cfg_path = BASE / "config.example.json"
    if not cfg_path.exists():
        raise SystemExit("config.json 이 없습니다. config.example.json 을 복사해서 만드세요.")
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))

    env_id = (os.environ.get("PCS_ID") or "").strip()
    env_pw = (os.environ.get("PCS_PW") or "").strip()
    if env_id and env_pw:
        cfg.setdefault("login", {})
        cfg["login"]["enabled"] = True
        cfg["login"]["user_id"] = env_id
        cfg["login"]["password"] = env_pw
    return cfg


def months_from(today, count):
    """조회할 YYYYMM 목록 (이번 달 포함)."""
    out, y, m = [], today.year, today.month
    for _ in range(max(1, count)):
        out.append("%d%02d" % (y, m))
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


# --- 페이지 조작 ----------------------------------------------------------
def do_login(page, login_cfg):
    """로그인 시도. (성공여부, 실패사유) 를 돌려준다.

    로그인에 실패해도 예외를 던지지 않는다. 비로그인으로도 구간별 스케줄 조회는
    되므로, 수집 자체를 통째로 실패시키는 대신 사유를 기록하고 계속 진행한다.
    실패 사유는 스냅샷에 남아 뷰어 상단에 경고로 표시된다.
    """
    if not login_cfg.get("enabled"):
        log("로그인 설정이 꺼져 있어 비로그인으로 진행합니다.")
        return False, "로그인 설정이 꺼져 있음"
    uid = (login_cfg.get("user_id") or "").strip()
    pwd = (login_cfg.get("password") or "").strip()
    if not uid or not pwd or uid.startswith("여기에"):
        log("경고: 아이디/비밀번호가 없어 비로그인으로 진행합니다.")
        return False, "아이디/비밀번호가 설정되지 않음"

    try:
        page.click(LOGIN_BTN)
        page.wait_for_selector(LOGIN_ID, state="visible", timeout=20000)
        page.fill(LOGIN_ID, uid)
        page.fill(LOGIN_PW, pwd)
        page.click(LOGIN_GO)
        page.wait_for_timeout(4000)
    except Exception as exc:
        log("경고: 로그인 창을 다루지 못했습니다 (%s). 비로그인으로 계속합니다." % exc)
        return False, "로그인 창 오류: %s" % exc

    if page.locator(LOGIN_ERR).count() and page.locator(LOGIN_ERR).first.is_visible():
        log("경고: 로그인 실패 - 아이디/비밀번호 또는 회원승인 상태를 확인하세요. 비로그인으로 계속합니다.")
        return False, "아이디/비밀번호가 틀렸거나 회원승인이 되지 않음"
    if page.locator(LOGIN_BTN).count() and page.locator(LOGIN_BTN).first.is_visible():
        log("경고: 로그인 후에도 LOGIN 버튼이 남아 있습니다. 비로그인 상태일 수 있습니다.")
        return False, "로그인 후에도 LOGIN 버튼이 남아 있음"
    log("로그인 성공 (%s)" % uid)
    return True, None


def open_section_page(page):
    """스케줄 > 구간별 화면 열기."""
    try:
        page.click(SCHEDULE_MENU, timeout=5000)
        page.wait_for_timeout(700)
        page.click(SECTION_MENU, timeout=5000)
    except PWTimeout:
        # 드롭다운이 열리지 않으면 메뉴 anchor 가 호출하는 함수를 직접 실행
        page.evaluate(OPEN_SECTION_JS)
    page.wait_for_selector(INQUIRY, state="visible", timeout=30000)
    page.wait_for_timeout(1500)


def set_port(page, field, code):
    """출발지/도착지 입력. 팝업 입력칸에 코드를 치고 Enter 가 아닌 Tab 으로 확정한다."""
    for attempt in (1, 2):
        page.click(field)
        page.wait_for_selector(POPUP_INPUT, timeout=15000)
        page.wait_for_timeout(900)
        page.keyboard.type(code, delay=90)
        page.wait_for_timeout(1600)
        page.keyboard.press("Tab")
        page.wait_for_timeout(1600)
        value = page.input_value(field)
        if value.strip():
            return value.strip()
        log("  포트 입력 재시도 (%s, %d회차)" % (code, attempt))
    raise RuntimeError("포트 코드 '%s' 입력 실패" % code)


def inquiry(page):
    """Inquiry 버튼을 눌러 조회하고, 화면이 받는 응답을 그대로 가져온다."""
    with page.expect_response(
        lambda r: "/sch/selectScheList" in r.url, timeout=45000
    ) as info:
        page.click(INQUIRY)
    resp = info.value
    page.wait_for_timeout(2500)
    return json.loads(resp.text()), json.loads(resp.request.post_data)


# --- 데이터 변환 ----------------------------------------------------------
def fmt_dt(day, hhmm):
    """('20260904', '1600') -> ('2026-09-04', '2026-09-04 16:00')"""
    if not day or len(day) != 8:
        return None, None
    d = "%s-%s-%s" % (day[:4], day[4:6], day[6:])
    hhmm = (hhmm or "0000").zfill(4)
    return d, "%s %s:%s" % (d, hhmm[:2], hhmm[2:])


def to_vessel(row):
    """조회 결과 한 줄 -> 화면에 쓸 형태.

    선명 + 항차코드(9자리) 뒤 5자리 = 사용자가 원하는 표기.
    예) VSL_NAME='DONGJIN VENUS', VSLVOY='JDJV0335N' -> 'DONGJIN VENUS 0335N'
    """
    name = (row.get("VSL_NAME") or "").strip()
    vslvoy = (row.get("VSLVOY") or "").strip()
    voyage = vslvoy[-5:] if len(vslvoy) >= 5 else vslvoy
    etd_date, etd = fmt_dt(row.get("EXPORT_DT"), row.get("EXPORT_HH"))
    eta_date, eta = fmt_dt(row.get("IMPORT_DT"), row.get("IMPORT_HH"))
    return {
        "label": ("%s %s" % (name, voyage)).strip(),
        "vessel": name,
        "voyage": voyage,
        "vslvoy": vslvoy,
        "route": (row.get("ROT") or "").strip(),
        "etd_date": etd_date,
        "etd": etd,
        "eta_date": eta_date,
        "eta": eta,
        "pol_name": (row.get("LD_PORTD") or "").strip(),
        "pod_name": (row.get("DC_PORTD") or "").strip(),
        "direct": (row.get("TS_MIN") or "").strip() == "D",
        "closed": (row.get("VSL_CLOSE") or "").strip() == "T",
    }


def dedupe(vessels):
    seen, out = set(), []
    for v in vessels:
        key = (v["vslvoy"], v["etd"], v["pol_name"], v["pod_name"])
        if key in seen:
            continue
        seen.add(key)
        out.append(v)
    out.sort(key=lambda v: (v["etd_date"] or "", v["etd"] or "", v["label"]))
    return out


# --- 항로 1개 수집 --------------------------------------------------------
def collect_route(page, pol, pod, month_list):
    open_section_page(page)
    pol_name = set_port(page, POL, pol)
    pod_name = set_port(page, POD, pod)
    log("  %s(%s) -> %s(%s)" % (pol, pol_name, pod, pod_name))

    data, payload = inquiry(page)
    rows = list(data.get("dma_search", {}).get("vcursor") or [])
    base = dict(payload["dma_search"])
    first_month = base.get("inpmon")
    log("    %s: %d건" % (first_month, len(rows)))

    for mon in month_list:
        if mon == first_month:
            continue
        req = dict(base)
        req["inpmon"] = mon
        req.pop("vcursor", None)
        try:
            txt = page.evaluate(FETCH_MONTH_JS, [{"dma_search": req}])
            more = json.loads(txt).get("dma_search", {}).get("vcursor") or []
            rows.extend(more)
            log("    %s: %d건" % (mon, len(more)))
        except Exception as exc:
            log("    %s: 조회 실패 (%s)" % (mon, exc))
        page.wait_for_timeout(600)

    return {
        "pol": pol,
        "pod": pod,
        "pol_name": pol_name,
        "pod_name": pod_name,
        "ok": True,
        "error": None,
        "vessels": dedupe([to_vessel(r) for r in rows]),
    }


# --- 메인 ----------------------------------------------------------------
def dump_debug(page, tag):
    """실패 시점의 화면과 HTML 을 logs/ 에 남긴다 (GitHub Actions 에서 내려받아 확인)."""
    try:
        LOGS.mkdir(exist_ok=True)
        stamp = now_kst().strftime("%H%M%S")
        safe = re.sub(r"[^A-Za-z0-9_.-]", "_", tag)
        png = LOGS / ("debug_%s_%s.png" % (safe, stamp))
        htm = LOGS / ("debug_%s_%s.html" % (safe, stamp))
        page.screenshot(path=str(png), full_page=True)
        htm.write_text(page.content(), encoding="utf-8")
        log("  디버그 저장: %s , %s" % (png.name, htm.name))
    except Exception as exc:
        log("  디버그 저장 실패: %s" % exc)


def write_index():
    """data/index.json — 뷰어가 서버 없이(GitHub Pages) 목록을 읽을 수 있게 한다."""
    name_re = re.compile(r"^(\d{4}-\d{2}-\d{2})_(AM|PM)\.json$")
    items = []
    for f in sorted(DATA.glob("*.json")):
        m = name_re.match(f.name)
        if not m:
            continue
        date, session = m.group(1), m.group(2)
        kor = "오전" if session == "AM" else "오후"
        item = {
            "date": date, "session": session, "session_kor": kor,
            "label": "%s %s" % (date.replace("-", "."), kor),
            "vessels": 0, "collected_at": "",
        }
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
            item["vessels"] = sum(len(r.get("vessels", [])) for r in d.get("routes", []))
            item["collected_at"] = d.get("collected_at", "")
        except Exception:
            pass
        items.append(item)
    items.sort(key=lambda x: (x["date"], x["session"]), reverse=True)
    (DATA / "index.json").write_text(
        json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")
    return len(items)


def run(session=None, show=False):
    cfg = load_config()
    now = now_kst()
    if session is None:
        session = "AM" if now.hour < 12 else "PM"
    session = session.upper()
    kor = "오전" if session == "AM" else "오후"
    month_list = months_from(now.date(), cfg.get("months_ahead", 3))

    log("=" * 60)
    log("수집 시작: %s %s  (조회 월: %s)" % (now.strftime("%Y.%m.%d"), kor, ", ".join(month_list)))

    snapshot = {
        "snapshot": "%s %s" % (now.strftime("%Y.%m.%d"), kor),
        "date": now.strftime("%Y-%m-%d"),
        "session": session,
        "session_kor": kor,
        "collected_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "months": month_list,
        "logged_in": False,
        "login_error": None,
        "routes": [],
    }

    headless = cfg.get("headless", True) and not show
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        ctx = browser.new_context(viewport={"width": 1600, "height": 1000}, locale="ko-KR")
        page = ctx.new_page()
        page.on("dialog", lambda d: d.accept())
        try:
            # 러너 네트워크가 느릴 수 있어 접속은 몇 번 재시도한다.
            for attempt in (1, 2, 3):
                try:
                    page.goto(SITE, wait_until="load", timeout=90000)
                    break
                except Exception as exc:
                    log("사이트 접속 실패 (%d/3): %s" % (attempt, exc))
                    if attempt == 3:
                        raise
                    page.wait_for_timeout(5000)
            page.wait_for_timeout(3500)
            snapshot["logged_in"], snapshot["login_error"] = do_login(page, cfg.get("login", {}))

            for route in cfg["routes"]:
                pol, pod = route["pol"], route["pod"]
                log("[%s -> %s] 조회" % (pol, pod))
                try:
                    snapshot["routes"].append(collect_route(page, pol, pod, month_list))
                except Exception as exc:
                    log("  실패: %s" % exc)
                    traceback.print_exc()
                    dump_debug(page, "route_%s_%s" % (pol, pod))
                    snapshot["routes"].append({
                        "pol": pol, "pod": pod,
                        "pol_name": pol, "pod_name": pod,
                        "ok": False, "error": str(exc), "vessels": [],
                    })
        except Exception:
            dump_debug(page, "fatal")
            raise
        finally:
            browser.close()

    total = sum(len(r["vessels"]) for r in snapshot["routes"])
    failed = [r for r in snapshot["routes"] if not r["ok"]]

    if total == 0:
        # 전부 실패한 경우 기존에 잘 받아둔 스냅샷을 빈 파일로 덮어쓰지 않는다.
        log("수집된 모선이 0건이라 저장하지 않습니다. (실패 항로 %d개)" % len(failed))
        return 1

    DATA.mkdir(exist_ok=True)
    out = DATA / ("%s_%s.json" % (snapshot["date"], session))
    out.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8")
    n = write_index()

    log("저장 완료: %s  (모선 %d건, 실패 항로 %d개, 전체 스냅샷 %d개)"
        % (out.name, total, len(failed), n))
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="PCS LINE 구간별 스케줄 수집기")
    ap.add_argument("--session", choices=["AM", "PM", "am", "pm"], default=None,
                    help="오전(AM)/오후(PM) 스냅샷 지정. 생략하면 현재 시각으로 판정")
    ap.add_argument("--show", action="store_true", help="브라우저 창을 띄워서 실행")
    args = ap.parse_args()
    try:
        sys.exit(run(args.session, args.show))
    except Exception as e:
        log("치명적 오류: %s" % e)
        traceback.print_exc()
        sys.exit(2)
