# -*- coding: utf-8 -*-
"""
수집된 스케줄 스냅샷을 달력으로 보여주는 로컬 웹 뷰어.
표준 라이브러리만 사용한다 (추가 설치 불필요).

    python server.py            # http://127.0.0.1:8777 열림
    python server.py --port 9000
    python server.py --no-open  # 브라우저 자동 실행 안 함
"""
import argparse
import json
import re
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
INDEX = BASE / "static" / "index.html"
NAME_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})_(AM|PM)\.json$")


def list_snapshots():
    """저장된 스냅샷 목록 (최신순)."""
    out = []
    if DATA.exists():
        for f in DATA.glob("*.json"):
            m = NAME_RE.match(f.name)
            if not m:
                continue
            date, session = m.group(1), m.group(2)
            item = {
                "date": date,
                "session": session,
                "session_kor": "오전" if session == "AM" else "오후",
                "label": "%s %s" % (date.replace("-", "."), "오전" if session == "AM" else "오후"),
                "vessels": 0,
                "collected_at": "",
            }
            try:
                d = json.loads(f.read_text(encoding="utf-8"))
                item["vessels"] = sum(len(r.get("vessels", [])) for r in d.get("routes", []))
                item["collected_at"] = d.get("collected_at", "")
            except Exception:
                pass
            out.append(item)
    out.sort(key=lambda x: (x["date"], x["session"]), reverse=True)
    return out


def read_snapshot(date, session):
    m = re.match(r"^\d{4}-\d{2}-\d{2}$", date or "")
    if not m or session not in ("AM", "PM"):
        return None
    f = DATA / ("%s_%s.json" % (date, session))
    if not f.exists():
        return None
    return json.loads(f.read_text(encoding="utf-8"))


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass  # 콘솔 조용히

    def _send(self, code, body, ctype):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False), "application/json; charset=utf-8")

    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)

        if u.path in ("/", "/index.html"):
            if not INDEX.exists():
                return self._send(500, "static/index.html 이 없습니다.", "text/plain; charset=utf-8")
            return self._send(200, INDEX.read_text(encoding="utf-8"), "text/html; charset=utf-8")

        if u.path == "/api/snapshots":
            return self._json(list_snapshots())

        if u.path == "/api/snapshot":
            date = (q.get("date") or [""])[0]
            session = (q.get("session") or [""])[0].upper()
            data = read_snapshot(date, session)
            if data is None:
                return self._json({"error": "해당 날짜/시간대의 수집 데이터가 없습니다."}, 404)
            return self._json(data)

        return self._send(404, "not found", "text/plain; charset=utf-8")


def main():
    ap = argparse.ArgumentParser(description="PCS 스케줄 뷰어")
    ap.add_argument("--port", type=int, default=None)
    ap.add_argument("--no-open", action="store_true")
    args = ap.parse_args()

    port = args.port
    if port is None:
        try:
            port = json.loads((BASE / "config.json").read_text(encoding="utf-8")).get("viewer_port", 8777)
        except Exception:
            port = 8777

    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = "http://127.0.0.1:%d/" % port
    print("PCS 스케줄 뷰어 실행 중 →  %s" % url)
    print("종료하려면 이 창에서 Ctrl+C")
    if not args.no_open:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n종료합니다.")


if __name__ == "__main__":
    main()
