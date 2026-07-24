"""계정·로그인 — 기획서 '로그인·권한' 항목

관리자가 아이디·비밀번호를 발급·배포한다(자율 가입 없음).
  · 관리자(admin) — 전체 조회·승인·리포트·계정 관리
  · 기관(org)     — 자기 기관 입력·업로드·조회만

비밀번호는 PBKDF2-SHA256 해시로만 저장한다(평문 저장 안 함).
세션은 메모리에 두므로 서버를 다시 켜면 모두 로그아웃된다.
"""

import hashlib
import json
import os
import secrets
from datetime import datetime

from storage import DATA_DIR

USERS_FILE = DATA_DIR / "users.json"
ITERATIONS = 200_000
SESSIONS = {}  # sid -> {"id":..., "role":..., "org":...}


# ── 비밀번호 ─────────────────────────────────────────────────────


def hash_pw(password, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), ITERATIONS)
    return salt, digest.hex()


def check_pw(password, salt, expected):
    _, got = hash_pw(password, salt)
    return secrets.compare_digest(got, expected)


# ── 계정 ─────────────────────────────────────────────────────────


def load_users():
    if not USERS_FILE.exists():
        return {}
    return json.loads(USERS_FILE.read_text(encoding="utf-8"))


def save_users(users):
    DATA_DIR.mkdir(exist_ok=True)
    USERS_FILE.write_text(json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8")


def create_user(user_id, password, role, org=""):
    """계정 발급. 이미 있으면 None."""
    user_id = user_id.strip()
    users = load_users()
    if not user_id or user_id in users:
        return None
    salt, digest = hash_pw(password)
    users[user_id] = {
        "아이디": user_id,
        "권한": role,  # admin | org
        "기관명": org.strip(),
        "salt": salt,
        "hash": digest,
        "발급시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    save_users(users)
    return users[user_id]


def set_password(user_id, password):
    users = load_users()
    if user_id not in users:
        return False
    users[user_id]["salt"], users[user_id]["hash"] = hash_pw(password)
    save_users(users)
    return True


def delete_user(user_id):
    users = load_users()
    if user_id not in users or users[user_id]["권한"] == "admin" and _admin_count(users) <= 1:
        return False  # 마지막 관리자는 지울 수 없음
    del users[user_id]
    save_users(users)
    for sid, s in list(SESSIONS.items()):
        if s["id"] == user_id:
            del SESSIONS[sid]
    return True


def _admin_count(users):
    return sum(1 for u in users.values() if u["권한"] == "admin")


def verify(user_id, password):
    u = load_users().get(user_id.strip())
    if not u or not check_pw(password, u["salt"], u["hash"]):
        return None
    return u


def ensure_admin():
    """첫 실행이면 관리자 계정을 만들고 임시 비밀번호를 돌려준다."""
    users = load_users()
    if _admin_count(users) > 0:
        return None
    pw = secrets.token_urlsafe(9)
    create_user("admin", pw, "admin")
    return pw


# ── 세션 ─────────────────────────────────────────────────────────


def login(user):
    sid = secrets.token_urlsafe(24)
    SESSIONS[sid] = {"id": user["아이디"], "role": user["권한"], "org": user.get("기관명", "")}
    return sid


def logout(sid):
    SESSIONS.pop(sid, None)


def session_of(cookie_header):
    """Cookie 헤더에서 세션을 찾는다. 없으면 None."""
    if not cookie_header:
        return None
    for chunk in cookie_header.split(";"):
        k, _, v = chunk.strip().partition("=")
        if k == "sid":
            return SESSIONS.get(v)
    return None
