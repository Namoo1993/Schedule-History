# -*- coding: utf-8 -*-
"""
ICON(icon.pcsline.co.kr) 3562 B/L Profitability Inquiry 실적 엑셀 내려받기.

사용법:
    python icon_actuals.py                 # 현재 시각 기준으로 오전/오후 자동 판정
    python icon_actuals.py --session AM    # 오전 이름으로 저장
    python icon_actuals.py --show          # 브라우저 창을 띄워서 동작 확인
    python icon_actuals.py --from 20260901 --to 20261231   # 기간 직접 지정

조회 기간은 전달 1일 ~ 이번 달 말일이다. (출항 전 항차는 실적이 없으므로 다음 달은 보지 않는다)
결과: actuals/YYYY-MM-DD_AM.xlsx       원본 엑셀 (git 에 올리지 않음)
      data/actuals.json                항차·POL·POD 별 TEU·금액 합계 (뷰어용)

엑셀에 운임·수익·화주명이 들어 있으므로 원본은 actuals/ 에만 두고(.gitignore),
뷰어에는 항차·POL·POD 별 합계만 올린다.

    python icon_actuals.py --build-only    # 이미 받은 엑셀로 합계 파일만 다시 만들기
"""
import argparse
import calendar
import json
import re
import sys
import time
import traceback
import warnings
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

from scraper import load_config, now_kst, log, dump_debug, describe_page

BASE = Path(__file__).resolve().parent
OUT = BASE / "actuals"
SUM_OUT = BASE / "data" / "actuals.json"
SITE = "https://icon.pcsline.co.kr/"
MENU_NO = "3562"
POL_LIST = "CN,VN,JP,TH,HK"
WAIT_MS = 20 * 60 * 1000   # 조회·다운로드는 길면 5분 넘게 걸린다고 해서 넉넉히

# --- 화면 element ---------------------------------------------------------
LOGIN_ID = "#mf_ibx_empCd"
LOGIN_PW = "#mf_sct_password"
LOGIN_GO = "#mf_btn_login"
# 메뉴 검색칸은 처음엔 숨어 있고, 위에 덮인 label 을 눌러야 나타난다.
MENU_LABEL = "#mf_wfm_header_ibx_menuNm_label"
MENU_INPUT = "#mf_wfm_header_ibx_menuNm_input"

P = "#mf_tac_layout_contents_00010640_body_"
DATE_FROM = P + "searchMapINPTM1_input"
DATE_TO = P + "searchMapINPTM2_input"
POL = P + "idpol"
POD = P + "idpod"
INQUIRY = P + "btn_search"
TOTAL_CNT = P + "totalCnt"
# 다운로드 버튼 id 가운데(wq_uuid_NNNN)는 화면을 열 때마다 바뀐다.
EXCEL = "a[id^='mf_tac_layout_contents_00010640_body_'][id$='_btnExcelDown']:visible"
ALERT_MSG = "[id^='mf_tac_layout_contents_00010640_body_alert'][id$='_tbx_message']"

# Inquiry 가 실제로 쏘는 조회 요청. 화면의 'Progress...' 창(processMsgLayer)은
# 자동 실행 브라우저에서는 DOM 에 잡히지 않아서, 대신 이 응답과 건수 표시를 기다린다.
QUERY_API = "/css/doc/selectCSSDOC3562R02_NEW"


def date_range(today):
    """2026-10-15 -> ('20260901', '20261031')  전달 1일 ~ 이번 달 말일"""
    y, m = today.year, today.month
    py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
    return "%d%02d01" % (py, pm), "%d%02d%02d" % (y, m, calendar.monthrange(y, m)[1])


def check_alert(page):
    """화면 안 Alert 창 문구가 있으면 돌려준다."""
    for el in page.locator(ALERT_MSG).all():
        if el.is_visible():
            return el.inner_text().strip()
    return None


def total_count(page):
    txt = page.locator(TOTAL_CNT).inner_text() if page.locator(TOTAL_CNT).count() else ""
    digits = re.sub(r"[^0-9]", "", txt)
    return int(digits) if digits else 0


def type_into(page, sel, text):
    page.click(sel)
    page.keyboard.press("Control+A")
    page.keyboard.type(text, delay=60)
    page.keyboard.press("Tab")
    page.wait_for_timeout(500)


def login(page, cfg):
    acc = cfg.get("icon_login") or {}
    uid, pwd = (acc.get("user_id") or "").strip(), (acc.get("password") or "").strip()
    if not uid or not pwd:
        raise RuntimeError("config.json 에 icon_login 아이디/비밀번호가 없습니다.")
    page.goto(SITE, wait_until="domcontentloaded", timeout=90000)
    page.wait_for_selector(LOGIN_ID, state="visible", timeout=60000)
    page.fill(LOGIN_ID, uid)
    page.fill(LOGIN_PW, pwd)
    page.click(LOGIN_GO)
    try:
        page.wait_for_selector(MENU_LABEL, state="visible", timeout=60000)
    except Exception:
        raise RuntimeError("로그인 실패 - 아이디/비밀번호를 확인하세요")
    page.wait_for_timeout(5000)   # 첫 화면 위젯들이 다 뜰 때까지
    log("ICON 로그인 성공 (%s)" % uid)


def open_menu(page):
    page.click(MENU_LABEL)
    page.wait_for_selector(MENU_INPUT, state="visible", timeout=10000)
    page.keyboard.type(MENU_NO, delay=120)
    page.wait_for_timeout(1500)
    page.keyboard.press("Enter")
    page.wait_for_selector(INQUIRY, state="visible", timeout=60000)
    page.wait_for_timeout(3000)
    log("  %s B/L Profitability Inquiry 화면 열림" % MENU_NO)


def inquiry(page, errors):
    """Inquiry 를 누르고 조회 결과 건수를 돌려준다. 서버 오류면 예외."""
    errors.clear()
    t0 = time.time()
    with page.expect_response(lambda r: r.url.endswith(QUERY_API), timeout=WAIT_MS) as info:
        page.click(INQUIRY)
    if info.value.status != 200:
        raise RuntimeError("조회 응답 HTTP %s" % info.value.status)
    # 응답을 받은 뒤에도 화면이 표를 채우고 보조 조회를 몇 번 더 한 다음에 건수를 바꾼다.
    for _ in range(180):
        if errors:
            raise RuntimeError("서버 오류: %s" % errors[-1])
        msg = check_alert(page)
        if msg:
            raise RuntimeError("조회 경고: %s" % msg)
        n = total_count(page)
        if n:
            log("  Inquiry 완료: %d건 (%.0f초)" % (n, time.time() - t0))
            return n
        page.wait_for_timeout(1000)
    raise RuntimeError("조회 결과가 0건입니다")


def read_rows(path):
    """엑셀을 (머리글 4행 이름 -> 열 번호, 6행부터의 데이터 행) 으로 읽는다."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")   # 사이트 엑셀에 기본 스타일이 없다는 경고
        ws = openpyxl.load_workbook(path).active
    head = {}
    for i, c in enumerate(ws[4]):
        if c.value and c.value not in head:
            head[c.value] = i
    return head, [r for r in ws.iter_rows(min_row=6, values_only=True) if any(r)]


def summarize(path):
    """항차(VVD) -> POL -> POD 별 TEU(U열) 합계와 Income S.TTL(AD열) 합계.

    POD 가 한국(KR) 이면 항구마다 따로, 그 밖은 나라별로 합친다 (RUVVO+RUVFP -> RU).
    RPT(= 금액 합계 / TEU 합계)는 뷰어에서 나눠 보여준다.
    Type(Q열) 이 FEE 인 행은 어떤 경우에도 넣지 않는다.
    """
    head, rows = read_rows(path)
    c_vvd, c_pol, c_pod = head.get("VVD", 6), head.get("POL", 13), head.get("POD", 14)
    c_teu = head.get("TEU", 20)
    c_type = head.get("Type", 16)
    c_obd = head.get("On Board\nDate", 10)
    c_amt = 29   # AD열 = Income 아래 S.TTL (5행 머리글이라 4행 이름으로는 못 찾는다)
    num = lambda v: float(v) if isinstance(v, (int, float)) else 0.0

    agg, first_obd = {}, {}
    for r in rows:
        vvd = (r[c_vvd] or "").strip()
        pol = (r[c_pol] or "").strip()
        pod = (r[c_pod] or "").strip()
        if not vvd or not pol:
            continue
        if str(r[c_type] or "").strip().upper() == "FEE":
            continue
        key = pod if pod.startswith("KR") else pod[:2]
        g = agg.setdefault(vvd, {}).setdefault((pol, key), {"teu": 0.0, "amt": 0.0, "bl": 0, "pods": set()})
        g["teu"] += num(r[c_teu])
        g["amt"] += num(r[c_amt])
        g["bl"] += 1
        g["pods"].add(pod)
        # POL 순서를 정할 때 스케줄에 없는 POL 은 가장 이른 On Board Date 로 대신한다
        obd = str(r[c_obd] or "")[:10]
        if obd and (vvd, pol) not in first_obd or obd and obd < first_obd[(vvd, pol)]:
            first_obd[(vvd, pol)] = obd

    out = {}
    for vvd, groups in agg.items():
        # POL 순, 같은 POL 안에서는 한국 항구 먼저, 그다음 나라 코드 순
        lines = sorted(groups.items(), key=lambda kv: (kv[0][0], not kv[0][1].startswith("KR"), kv[0][1]))
        out[vvd] = [{
            "pol": pol, "pod": key, "pods": sorted(g["pods"]),
            "teu": round(g["teu"], 2), "amt": round(g["amt"], 2), "bl": g["bl"],
            "obd": first_obd.get((vvd, pol), ""),
        } for (pol, key), g in lines]
    return out


def build_summary(xlsx, period):
    """뷰어용 항차별 합계 파일을 만든다. (비밀번호는 나중에 붙이기로 해서 지금은 평문)"""
    summary = summarize(xlsx)
    payload = {
        "updated": now_kst().strftime("%Y-%m-%d %H:%M"),
        "period": period,
        "source": xlsx.name,
        "vvd": summary,
    }
    SUM_OUT.parent.mkdir(exist_ok=True)
    SUM_OUT.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    log("  뷰어용 실적 저장: data/%s (항차 %d개)" % (SUM_OUT.name, len(summary)))
    return True


def latest_xlsx():
    files = sorted(OUT.glob("????-??-??_[AP]M.xlsx"))
    return files[-1] if files else None


def count_rows(path):
    """내려받은 엑셀의 B/L 건수. 머리글이 4~5행이고 6행부터 데이터다."""
    head, rows = read_rows(path)
    col = head.get("B/L No.", 7)
    return sum(1 for r in rows if r[col])


def run(session=None, show=False, d_from=None, d_to=None):
    cfg = load_config()
    now = now_kst()
    if session is None:
        session = "AM" if now.hour < 12 else "PM"
    session = session.upper()
    auto_from, auto_to = date_range(now.date())
    d_from, d_to = d_from or auto_from, d_to or auto_to

    log("=" * 60)
    log("ICON 실적 다운로드 시작: %s %s  (기간 %s ~ %s, POL %s)"
        % (now.strftime("%Y.%m.%d"), session, d_from, d_to, POL_LIST))

    OUT.mkdir(exist_ok=True)
    out = OUT / ("%s_%s.xlsx" % (now.strftime("%Y-%m-%d"), session))
    tmp = out.with_suffix(".part.xlsx")
    headless = cfg.get("headless", True) and not show

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        # 화면이 좁으면 조회조건 줄이 가로로 밀려 엉뚱한 칸이 눌린다
        ctx = browser.new_context(viewport={"width": 2300, "height": 1000},
                                  locale="ko-KR", accept_downloads=True)
        page = ctx.new_page()
        # 사이트가 서버 오류를 브라우저 경고창으로 띄운다 ("서버에러입니다(timeout)").
        # 첫 화면 위젯에서도 뜨므로 조회 직전에 비우고, 조회 중에 뜬 것만 실패로 본다.
        errors = []

        def on_dialog(d):
            m = d.message.strip()
            log("  브라우저 경고: %s" % m)
            if "에러" in m or "error" in m.lower():
                errors.append(m)
            d.accept()
        page.on("dialog", on_dialog)

        try:
            login(page, cfg)
            open_menu(page)

            type_into(page, DATE_FROM, d_from)
            type_into(page, DATE_TO, d_to)
            type_into(page, POL, POL_LIST)
            page.fill(POD, "")
            log("  조회조건: %s ~ %s, POL=%s, POD=(비움)"
                % (page.input_value(DATE_FROM), page.input_value(DATE_TO), page.input_value(POL)))

            # 서버가 가끔 timeout 을 낸다. 한 번 더 눌러 본다.
            for attempt in (1, 2):
                try:
                    n = inquiry(page, errors)
                    break
                except Exception as exc:
                    log("  조회 실패 (%d/2): %s" % (attempt, exc))
                    if attempt == 2:
                        raise
                    page.wait_for_timeout(10000)

            t0 = time.time()
            errors.clear()
            with page.expect_download(timeout=WAIT_MS) as dl:
                page.locator(EXCEL).first.click()
            dl.value.save_as(str(tmp))
            got = count_rows(tmp)
            log("  엑셀 다운로드 완료: %d건 (%.0f초)" % (got, time.time() - t0))
            if got == 0:
                raise RuntimeError("내려받은 엑셀이 비어 있습니다")
            if got != n:
                log("  경고: 화면 건수(%d)와 엑셀 건수(%d)가 다릅니다" % (n, got))
        except Exception:
            log("  현재 화면: %s" % describe_page(page))
            dump_debug(page, "icon_fatal")
            tmp.unlink(missing_ok=True)
            raise
        finally:
            browser.close()

    # 검증을 통과한 파일만 최종 이름으로 바꾼다. 실패하면 이전 파일이 그대로 남는다.
    tmp.replace(out)
    log("저장 완료: actuals/%s (%.0f KB)" % (out.name, out.stat().st_size / 1024))
    build_summary(out, [d_from, d_to])
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="ICON B/L Profitability 실적 엑셀 다운로드")
    ap.add_argument("--session", choices=["AM", "PM", "am", "pm"], default=None,
                    help="오전(AM)/오후(PM) 이름으로 저장. 생략하면 현재 시각으로 판정")
    ap.add_argument("--show", action="store_true", help="브라우저 창을 띄워서 실행")
    ap.add_argument("--from", dest="d_from", help="조회 시작일 YYYYMMDD (생략하면 스케줄 조회 월 기준)")
    ap.add_argument("--to", dest="d_to", help="조회 종료일 YYYYMMDD")
    ap.add_argument("--build-only", action="store_true",
                    help="조회하지 않고 가장 최근 엑셀로 뷰어용 합계 파일만 다시 만든다")
    args = ap.parse_args()
    if args.build_only:
        x = latest_xlsx()
        if not x:
            raise SystemExit("actuals/ 에 받은 엑셀이 없습니다.")
        sys.exit(0 if build_summary(x, None) else 1)
    try:
        sys.exit(run(args.session, args.show, args.d_from, args.d_to))
    except Exception as e:
        log("ICON 실적 치명적 오류: %s" % e)
        traceback.print_exc()
        sys.exit(2)
