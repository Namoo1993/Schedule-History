# -*- coding: utf-8 -*-
"""
GitHub Pages 에 올릴 데이터를 비밀번호로 암호화한다.

    python encrypt_site.py

data/*.json (평문, git 에 올리지 않음)  ->  site_data/*.json (암호문, git 에 올림)

저장소와 사이트가 공개라서 평문을 올리면 주소만 알면 누구나 열 수 있다.
그래서 올리는 쪽은 암호문만 두고, 뷰어(static/index.html)가 비밀번호로 푼다.
로컬 뷰어(server.py)는 data/ 의 평문을 그대로 쓰므로 비밀번호를 묻지 않는다.

- 비밀번호: config.json 의 "site_password"
- 키: PBKDF2-SHA256(비밀번호, salt, ITER) -> AES-256-GCM
- 파일: gzip 으로 줄인 뒤 암호화. {"v":1, "iv": base64, "ct": base64}
- salt 와 반복 횟수는 site_data/key.json 에 둔다 (비밀이 아님)

내용이 바뀐 파일만 다시 암호화한다. 매번 새로 암호화하면 IV 가 달라져서
바뀌지 않은 스냅샷까지 커밋에 끼어든다. (.enc_cache.json 에 평문 해시를 기억)
"""
import base64
import gzip
import hashlib
import json
import os
import sys
from pathlib import Path

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

BASE = Path(__file__).resolve().parent
SRC = BASE / "data"
DST = BASE / "site_data"
KEY_FILE = DST / "key.json"
CACHE = BASE / ".enc_cache.json"
ITER = 200000


def b64(b):
    return base64.b64encode(b).decode("ascii")


def load_password():
    cfg = json.loads((BASE / "config.json").read_text(encoding="utf-8"))
    pw = (cfg.get("site_password") or "").strip()
    if not pw:
        raise SystemExit("config.json 에 site_password 가 없습니다.")
    return pw


def derive(pw, salt, it):
    return hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, it, 32)


def main():
    pw = load_password()
    DST.mkdir(exist_ok=True)

    # salt 는 처음 한 번 만들고 계속 쓴다. 비밀번호가 바뀌면 전부 다시 암호화한다.
    meta = json.loads(KEY_FILE.read_text(encoding="utf-8")) if KEY_FILE.exists() else None
    if not meta:
        meta = {"v": 1, "kdf": "PBKDF2-SHA256", "iter": ITER, "salt": b64(os.urandom(16))}
        KEY_FILE.write_text(json.dumps(meta, indent=1), encoding="utf-8")
    key = derive(pw, base64.b64decode(meta["salt"]), meta["iter"])
    fp = hashlib.sha256(key).hexdigest()   # 키가 바뀌었는지 알아보는 용도 (로컬에만 저장)

    cache = {}
    if CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}
    if cache.get("key") != fp:
        cache = {"key": fp, "files": {}}

    aes = AESGCM(key)
    names = sorted(f.name for f in SRC.glob("*.json"))
    done = 0
    for name in names:
        raw = (SRC / name).read_bytes()
        h = hashlib.sha256(raw).hexdigest()
        out = DST / name
        if cache["files"].get(name) == h and out.exists():
            continue
        iv = os.urandom(12)
        ct = aes.encrypt(iv, gzip.compress(raw, mtime=0), None)
        out.write_text(json.dumps({"v": 1, "iv": b64(iv), "ct": b64(ct)}), encoding="utf-8")
        cache["files"][name] = h
        done += 1

    # data/ 에서 사라진 파일은 사이트에서도 지운다
    for f in DST.glob("*.json"):
        if f.name != KEY_FILE.name and f.name not in names:
            f.unlink()
            cache["files"].pop(f.name, None)

    CACHE.write_text(json.dumps(cache, indent=1), encoding="utf-8")
    print("암호화 완료: %d개 새로 암호화 / 전체 %d개 -> site_data/" % (done, len(names)))


if __name__ == "__main__":
    sys.exit(main())
