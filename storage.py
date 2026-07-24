"""주차별 저장소 — specs/자동화_흐름도.md 의 '주차별 저장소'

기관이 화면에서 직접 입력하거나 엑셀로 올린 제출 건을 주차별로 쌓아 둔다.
JSON 파일이라 열어서 눈으로 확인·수정할 수 있다.

    data/roster.json      대상 훈련기관 명단
    data/2026-W30.json    해당 주차의 기관별 제출 건
"""

import json
import os
import re
from datetime import datetime
from pathlib import Path

# 저장 위치. 배포 호스트에서 영구 디스크에 쌓으려면 환경변수 DATA_DIR 로 그 경로를 준다
# (예: Render/Fly 의 마운트 경로 /var/data). 없으면 이 폴더의 data/ 를 쓴다.
DATA_DIR = Path(os.environ.get("DATA_DIR") or (Path(__file__).parent / "data"))
# 과정 한 건에 저장하는 필드 (정부 「지산맞」 양식 기준). 기관명은 제출 키에서 오므로 뺀다.
PERF_COLS = ["구분", "정기수시", "과정구분", "NCS대분류명", "KECO세분류명", "과정명",
             "훈련목표인원", "훈련실시인원", "중도탈락자", "훈련중", "훈련수료인원", "취업인원"]
# 기관 연간 목표(교육실적 시트) — 제출 건 단위로 저장하고 집계 때 각 행에 붙인다.
TARGET_KEYS = ["목표_총", "목표_양성", "목표_향상", "목표_수시"]
NOTE_COLS = ["과정명", "분류", "내용", "확인필요"]
PLAN_COLS = ["날짜", "구분", "내용"]
STATUSES = ("제출", "승인", "반려")


def week_key(year, week):
    return f"{int(year)}-W{int(week):02d}"


def parse_week(key):
    m = re.match(r"^(\d{4})-W(\d{1,2})$", str(key).strip())
    return (int(m.group(1)), int(m.group(2))) if m else None


def last_year_key(key):
    y, w = parse_week(key)
    return week_key(y - 1, w)


def _path(key):
    return DATA_DIR / f"{key}.json"


def _blank(key):
    y, w = parse_week(key)
    return {"주차": key, "연": y, "주차번호": w, "월": None, "제출": {}}


# ── 주차 데이터 ──────────────────────────────────────────────────


def load_week(key):
    """주차 데이터를 읽는다. 없으면 None."""
    if not parse_week(key):
        return None
    p = _path(key)
    if not p.exists():
        return None
    return json.loads(p.read_text(encoding="utf-8"))


def save_submission(key, org, month, perf, notes, plans, source, targets=None):
    """기관 한 곳의 제출 건을 저장한다. 같은 기관이 다시 내면 덮어쓴다(재입력).

    targets = 기관 연간 목표(교육실적 시트에서 읽은 목표_총/양성/향상/수시). 집계 때 분모로 쓴다.
    """
    DATA_DIR.mkdir(exist_ok=True)
    data = load_week(key) or _blank(key)
    if month:
        data["월"] = int(month)
    prev = data["제출"].get(org, {})
    data["제출"][org] = {
        "기관명": org,
        "실적": perf,
        "목표": {k: (targets or {}).get(k) for k in TARGET_KEYS},
        "특이사항": notes,
        "주요일정": plans,
        "출처": source,
        "제출시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "상태": "제출",  # 다시 내면 승인 상태는 초기화된다
        "이전상태": prev.get("상태"),
    }
    _path(key).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


def set_status(key, org, status):
    data = load_week(key)
    if not data or org not in data["제출"] or status not in STATUSES:
        return None
    data["제출"][org]["상태"] = status
    _path(key).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


def delete_submission(key, org):
    data = load_week(key)
    if not data or org not in data["제출"]:
        return None
    del data["제출"][org]
    _path(key).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


def list_weeks():
    """저장된 주차를 최신순으로."""
    if not DATA_DIR.exists():
        return []
    keys = [p.stem for p in DATA_DIR.glob("*.json") if parse_week(p.stem)]
    return sorted(keys, reverse=True)


def prev_week_key(key):
    """저장소에 있는 주차 중 이 주차 바로 직전 것 (없으면 None) — 흐름도 8단계."""
    y, w = parse_week(key)
    earlier = [k for k in list_weeks() if parse_week(k) < (y, w)]
    return earlier[0] if earlier else None


# ── 대상 명단 ────────────────────────────────────────────────────


def load_roster():
    p = DATA_DIR / "roster.json"
    if not p.exists():
        return []
    return json.loads(p.read_text(encoding="utf-8"))


def save_roster(names):
    DATA_DIR.mkdir(exist_ok=True)
    names = [n.strip() for n in names if n.strip()]
    (DATA_DIR / "roster.json").write_text(json.dumps(names, ensure_ascii=False, indent=2), encoding="utf-8")
    return names


# ── 집계용 변환 ──────────────────────────────────────────────────


def to_rows(data, only_approved=False):
    """저장된 제출 건 → feature1_aggregate.process_rows 가 받는 행 목록."""
    rows, n = [], 1
    for org, sub in data["제출"].items():
        if only_approved and sub.get("상태") != "승인":
            continue
        목표 = sub.get("목표") or {}
        for p in sub["실적"]:
            n += 1
            row = {"_행": n, "기관명": org, **{c: p.get(c) for c in PERF_COLS}}
            for k in TARGET_KEYS:  # 기관 연간 목표를 각 행에 붙여 집계기로 넘긴다
                row[k] = 목표.get(k)
            rows.append(row)
    return rows


def to_notes(data):
    out = []
    for org, sub in data["제출"].items():
        for x in sub.get("특이사항", []):
            out.append({"기관명": org, **{c: str(x.get(c, "") or "") for c in NOTE_COLS}})
    return out


def to_plans(data):
    out = []
    for org, sub in data["제출"].items():
        for x in sub.get("주요일정", []):
            out.append({"기관명": org, **{c: str(x.get(c, "") or "") for c in PLAN_COLS}})
    return out
