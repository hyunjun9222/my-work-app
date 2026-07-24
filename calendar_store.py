"""일정 조사 저장소 — 캘린더 탭

관리자가 기간을 정해 "이 중 언제가 되나요" 를 물으면(예: 10월 2주차 워크숍),
각 기관 계정이 날짜마다 가능/불가로 답한다.

    data/calendar.json   조사 목록 + 기관별 응답

응답은 기본적으로 **관리자만** 본다. 조사마다 '전체 공개' 를 켜면
참여 기관들도 서로의 답을 볼 수 있다(기본값은 비공개).

주차 저장소(storage.py)와 같은 방식으로 사람이 열어 고칠 수 있는 JSON 한 개만 쓴다.
"""

import json
from datetime import date, datetime, timedelta

from storage import DATA_DIR

FILE = DATA_DIR / "calendar.json"
MAX_DAYS = 31  # 한 조사에서 물을 수 있는 최대 기간(날짜 칸이 너무 넓어지지 않게)
ANSWERS = ("가능", "불가", "미정")
요일들 = "월화수목금토일"


def _blank():
    return {"다음번호": 1, "조사": []}


def load():
    if not FILE.exists():
        return _blank()
    data = json.loads(FILE.read_text(encoding="utf-8"))
    data.setdefault("다음번호", 1)
    data.setdefault("조사", [])
    return data


def _save(data):
    DATA_DIR.mkdir(exist_ok=True)
    FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


# ── 날짜 ─────────────────────────────────────────────────────────


def parse_day(s):
    """'2026-10-05' → date. 날짜가 아니면 None."""
    try:
        return date.fromisoformat(str(s).strip())
    except (ValueError, TypeError, AttributeError):
        return None


def days_of(ev):
    """조사 기간의 날짜 목록(문자열). 기간이 잘못되어 있으면 빈 목록."""
    시작, 종료 = parse_day(ev.get("시작일")), parse_day(ev.get("종료일"))
    if not 시작 or not 종료 or 종료 < 시작:
        return []
    return [(시작 + timedelta(days=i)).isoformat() for i in range((종료 - 시작).days + 1)]


def day_label(iso):
    """'2026-10-05' → '10/5(월)'. 날짜가 아니면 원문 그대로."""
    d = parse_day(iso)
    return f"{d.month}/{d.day}({요일들[d.weekday()]})" if d else str(iso)


def is_weekend(iso):
    d = parse_day(iso)
    return bool(d) and d.weekday() >= 5


def is_closed(ev, today=None):
    """마감일이 지났는지. 마감일이 없으면 늘 열려 있다."""
    마감 = parse_day(ev.get("마감일"))
    return bool(마감) and (today or date.today()) > 마감


# ── 조사 ─────────────────────────────────────────────────────────


def create(제목, 시작일, 종료일, 설명="", 마감일="", 대상=None, 공개=False, 작성자=""):
    """조사를 만든다. 반환 (조사, 오류문구) — 조사가 None 이면 오류문구를 보여준다."""
    제목 = (제목 or "").strip()
    시작, 종료 = parse_day(시작일), parse_day(종료일)
    if not 제목:
        return None, "제목을 입력해 주세요."
    if not 시작 or not 종료:
        return None, "기간을 날짜로 입력해 주세요."
    if 종료 < 시작:
        return None, "종료일이 시작일보다 빠릅니다. 기간을 다시 확인해 주세요."
    if (종료 - 시작).days + 1 > MAX_DAYS:
        return None, f"한 번에 물을 수 있는 기간은 최대 {MAX_DAYS}일입니다."
    마감 = parse_day(마감일)
    if 마감일 and not 마감:
        return None, "마감일을 날짜로 입력해 주세요."

    data = load()
    ev = {
        "번호": data["다음번호"],
        "제목": 제목,
        "설명": (설명 or "").strip(),
        "시작일": 시작.isoformat(),
        "종료일": 종료.isoformat(),
        "마감일": 마감.isoformat() if 마감 else "",
        "대상": [t.strip() for t in (대상 or []) if t.strip()],
        "공개": bool(공개),
        "작성자": 작성자,
        "생성시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "응답": {},
    }
    data["조사"].append(ev)
    data["다음번호"] += 1
    _save(data)
    return ev, None


def get(번호):
    try:
        번호 = int(번호)
    except (TypeError, ValueError):
        return None
    return next((ev for ev in load()["조사"] if ev["번호"] == 번호), None)


def list_events(newest_first=True):
    조사 = load()["조사"]
    return list(reversed(조사)) if newest_first else 조사


def answer(번호, 기관, 날짜별, 메모="", 응답자=""):
    """기관 한 곳의 응답을 저장한다(다시 내면 덮어쓴다). 반환 (조사, 오류문구)."""
    기관 = (기관 or "").strip()
    if not 기관:
        return None, "소속 기관이 없는 계정은 응답할 수 없습니다."
    data = load()
    ev = next((x for x in data["조사"] if x["번호"] == int(번호)), None)
    if not ev:
        return None, "없는 일정 조사입니다."
    if is_closed(ev):
        return None, "응답 마감일이 지났습니다. 관리자에게 문의해 주세요."

    유효 = {d: v for d, v in (날짜별 or {}).items() if d in days_of(ev) and v in ANSWERS}
    ev.setdefault("응답", {})[기관] = {
        "기관명": 기관,
        "응답자": 응답자,
        "날짜별": 유효,
        "메모": (메모 or "").strip(),
        "응답시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    _save(data)
    return ev, None


def set_options(번호, 공개=None, 마감일=None):
    """공개 여부·마감일을 바꾼다(관리자). 반환 (조사, 오류문구)."""
    data = load()
    ev = next((x for x in data["조사"] if x["번호"] == int(번호)), None)
    if not ev:
        return None, "없는 일정 조사입니다."
    if 공개 is not None:
        ev["공개"] = bool(공개)
    if 마감일 is not None:
        마감 = parse_day(마감일)
        if 마감일 and not 마감:
            return None, "마감일을 날짜로 입력해 주세요."
        ev["마감일"] = 마감.isoformat() if 마감 else ""
    _save(data)
    return ev, None


def delete(번호):
    data = load()
    남길 = [x for x in data["조사"] if x["번호"] != int(번호)]
    if len(남길) == len(data["조사"]):
        return False
    data["조사"] = 남길
    _save(data)
    return True


# ── 집계 ─────────────────────────────────────────────────────────


def targets(ev, roster=None):
    """이 조사에 답해야 할 기관 목록. 대상을 비워 두면 대상 명단 전체."""
    대상 = [t for t in ev.get("대상", []) if t]
    if 대상:
        return 대상
    명단 = list(roster or [])
    for org in ev.get("응답", {}):
        if org not in 명단:
            명단.append(org)
    return 명단


def pending(ev, roster=None):
    """아직 답하지 않은 기관."""
    응답 = ev.get("응답", {})
    return [t for t in targets(ev, roster) if t not in 응답]


def tally(ev):
    """날짜별 인원. 반환 {날짜: {"가능": n, "불가": n}} — 미정은 세지 않는다."""
    out = {d: {"가능": 0, "불가": 0} for d in days_of(ev)}
    for a in ev.get("응답", {}).values():
        for d, v in (a.get("날짜별") or {}).items():
            if d in out and v in out[d]:
                out[d][v] += 1
    return out


def best_days(ev):
    """'가능' 이 가장 많은 날짜들. 아무도 가능하지 않으면 빈 목록."""
    표 = tally(ev)
    최다 = max((v["가능"] for v in 표.values()), default=0)
    return [d for d, v in 표.items() if v["가능"] == 최다] if 최다 > 0 else []


def is_target(ev, org, roster=None):
    """이 기관이 답할 차례인지(대상 명단에 있거나 이미 답했는지)."""
    return bool(org) and (org in targets(ev, roster) or org in ev.get("응답", {}))


def can_open(ev, role, org, roster=None):
    """조사 화면을 열어 볼 수 있는지. 대상이 아니어도 '전체 공개' 면 볼 수 있다."""
    return role == "admin" or bool(ev.get("공개")) or is_target(ev, org, roster)


def can_see_answers(ev, role, org, roster=None):
    """다른 기관의 응답까지 볼 수 있는지. 관리자는 항상, 기관은 '전체 공개' 일 때만."""
    return role == "admin" or bool(ev.get("공개"))


def open_for(org, roster=None, today=None):
    """이 기관이 아직 답하지 않은, 마감 전 조사 목록."""
    out = []
    for ev in list_events():
        if is_closed(ev, today) or org in ev.get("응답", {}):
            continue
        if org in targets(ev, roster):
            out.append(ev)
    return out
