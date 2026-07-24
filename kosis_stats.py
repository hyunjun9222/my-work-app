"""고용 통계 수집 — KOSIS 공유서비스 OpenAPI

전국·시도별 고용률 / 실업률 / 생산가능인구(15세이상인구)를 월별로 받아
시계열 선 그래프에 쓸 형태로 정리한다.

통계표: 국가데이터처(구 통계청) 경제활동인구조사
        orgId=101, tblId=DT_1DA7004S "행정구역(시도)별 경제활동인구"
        → 이 표 하나에 15세이상인구·고용률·실업률이 전국/시도별로 다 들어 있다.

응답 한 행에 항목명(ITM_NM)·지역명(C1_NM)이 같이 오므로 코드표를 따로 받지 않는다.
itmId=ALL, objL1=ALL 로 통째로 받아 필요한 것만 골라 쓴다.

키는 .env 의 KOSIS_API_KEY 를 쓴다. 발급: https://kosis.kr/openapi/

사용법:
    python kosis_stats.py                 # 저장된 결과 보기 (없으면 수집)
    python kosis_stats.py --refresh       # 최근 36개월 새로 수집
    python kosis_stats.py --months 60     # 기간 지정해 수집
    python kosis_stats.py --region 경상북도 --region 계
    python kosis_stats.py --list          # 통계표의 항목·지역 이름 확인용
"""

import argparse
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

from imagegen import load_env

BASE = Path(__file__).parent
CACHE_FILE = BASE / "data" / "kosis_stats.json"

API_URL = "https://kosis.kr/openapi/Param/statisticsParameterData.do"
ORG_ID = "101"
TBL_ID = "DT_1DA7004S"  # 행정구역(시도)별 경제활동인구 (월)
PRD_SE = "M"  # 월간

# 화면에 쓸 지표 이름 → 통계표의 항목명(ITM_NM). 순서가 곧 그래프 탭 순서다.
INDICATORS = [
    ("고용률", "고용률"),
    ("실업률", "실업률"),
    ("생산가능인구", "15세이상인구"),
]

# 통계표는 전국을 "계" 로 적는다. 화면에는 "전국" 으로 보여준다.
REGION_ALIAS = {"계": "전국"}

# 통계표에 있는 지역 전부 (화면의 선택 목록에 쓴다). 순서는 통계표 순서 그대로.
ALL_REGIONS = [
    "전국", "서울특별시", "부산광역시", "대구광역시", "인천광역시", "광주광역시",
    "대전광역시", "울산광역시", "세종특별자치시", "경기도", "강원도", "충청북도",
    "충청남도", "전라북도", "전라남도", "경상북도", "경상남도", "제주도",
]

DEFAULT_REGIONS = ["전국", "경상북도", "대구광역시", "서울특별시"]
DEFAULT_MONTHS = 36

# 선 그래프 색은 4개까지만 검증돼 있다. 그 이상은 색을 돌려쓰지 않고 지역 수를 줄인다.
MAX_SERIES = 4


class StatError(Exception):
    """사용자에게 그대로 보여줄 수 있는 오류."""


# ── API 호출 ─────────────────────────────────────────────────────


def api_key():
    load_env()
    key = (os.environ.get("KOSIS_API_KEY") or "").strip()
    if not key:
        raise StatError(
            "KOSIS API 키가 없습니다. 이 폴더의 .env 파일에 KOSIS_API_KEY 를 채우세요.\n"
            "발급: https://kosis.kr/openapi/ → 로그인 후 '활용신청'(즉시 발급, 무료)"
        )
    return key


def fetch_rows(months):
    """최근 months 개월치를 항목·지역 전부 받아온다."""
    params = {
        "method": "getList",
        "apiKey": api_key(),
        "orgId": ORG_ID,
        "tblId": TBL_ID,
        "itmId": "ALL",
        "objL1": "ALL",
        "prdSe": PRD_SE,
        "newEstPrdCnt": str(months),
        "format": "json",
        "jsonVD": "Y",
    }
    url = API_URL + "?" + urllib.parse.urlencode(params)
    try:
        with urllib.request.urlopen(url, timeout=90) as resp:
            body = resp.read().decode("utf-8")
    except urllib.error.URLError as ex:
        raise StatError(f"KOSIS 접속에 실패했습니다 — {ex}")

    try:
        data = json.loads(body)
    except json.JSONDecodeError:
        raise StatError(f"KOSIS 응답을 읽지 못했습니다 — {body[:200]}")

    # 오류는 목록이 아니라 dict 로 온다: {"err":"20","errMsg":"..."}
    if isinstance(data, dict):
        raise StatError(f"KOSIS 오류 {data.get('err', '?')} — {data.get('errMsg', data)}")
    if not data:
        raise StatError("KOSIS 가 빈 결과를 돌려줬습니다. 통계표 코드나 기간을 확인하세요.")
    return data


# ── 정리 ─────────────────────────────────────────────────────────


def region_name(row):
    name = (row.get("C1_NM") or "").strip()
    return REGION_ALIAS.get(name, name)


def num(value):
    """숫자로 못 읽으면 None (결측)."""
    try:
        return float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None


def shape(rows, regions):
    """API 응답 → {시점: [...], 지표: {이름: {단위, 지역: {지역명: [값...]}}}}"""
    항목별 = {itm: label for label, itm in INDICATORS}
    원하는지역 = set(regions)

    값 = {label: {} for label, _ in INDICATORS}
    단위 = {}
    시점 = set()

    for row in rows:
        label = 항목별.get((row.get("ITM_NM") or "").strip())
        region = region_name(row)
        period = row.get("PRD_DE")
        if not label or not period or region not in 원하는지역:
            continue
        시점.add(period)
        값[label].setdefault(region, {})[period] = num(row.get("DT"))
        단위.setdefault(label, (row.get("UNIT_NM") or "").strip())

    if not 시점:
        raise StatError(
            "요청한 지역·지표에 해당하는 값이 없습니다. `python kosis_stats.py --list` 로 이름을 확인하세요."
        )

    시점 = sorted(시점)
    지표 = {}
    for label, _ in INDICATORS:
        per_region = 값[label]
        if not per_region:
            continue
        지표[label] = {
            "단위": 단위.get(label, ""),
            # 요청한 순서를 유지한다 (표·범례 순서가 화면마다 흔들리지 않도록)
            "지역": {r: [per_region[r].get(p) for p in 시점] for r in regions if r in per_region},
        }
    return {"시점": 시점, "지표": 지표}


def collect(months=DEFAULT_MONTHS, regions=None, progress=None):
    """수집 → 정리 → 저장."""
    regions = regions or DEFAULT_REGIONS
    if progress:
        progress(f"최근 {months}개월 수집 ({', '.join(regions)})")

    data = shape(fetch_rows(months), regions)
    data["수집시각"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    data["출처"] = {
        "이름": "국가데이터처 경제활동인구조사 · 행정구역(시도)별 경제활동인구",
        "url": f"https://kosis.kr/statHtml/statHtml.do?orgId={ORG_ID}&tblId={TBL_ID}",
    }
    save_cache(data)
    return data


# ── 캐시 ─────────────────────────────────────────────────────────


def load_cache():
    if not CACHE_FILE.exists():
        return {}
    try:
        return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_cache(data):
    CACHE_FILE.parent.mkdir(exist_ok=True)
    CACHE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def fmt_period(period):
    """202605 → 2026-05"""
    return f"{period[:4]}-{period[4:]}" if len(period) == 6 else period


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="KOSIS 고용 통계 수집")
    ap.add_argument("--refresh", action="store_true", help="새로 수집")
    ap.add_argument("--months", type=int, default=DEFAULT_MONTHS, help=f"최근 몇 개월 (기본 {DEFAULT_MONTHS})")
    ap.add_argument("--region", action="append", help="지역 지정 (여러 번 사용 가능)")
    ap.add_argument("--list", action="store_true", help="통계표의 항목·지역 이름 출력")
    args = ap.parse_args()

    try:
        if args.list:
            rows = fetch_rows(1)
            print("\n■ 항목")
            for name in dict.fromkeys((r.get("ITM_NM") or "").strip() for r in rows):
                print(f"  {name}")
            print("\n■ 지역")
            for name in dict.fromkeys(region_name(r) for r in rows):
                print(f"  {name}")
            print()
            sys.exit(0)

        data = load_cache()
        if args.refresh or args.region or args.months != DEFAULT_MONTHS or not data:
            data = collect(args.months, args.region, progress=lambda m: print(f"  {m}…"))
    except StatError as ex:
        print(f"\n오류: {ex}\n")
        sys.exit(1)

    시점 = data["시점"]
    print(f"\n고용 통계  (최종 수집 {data.get('수집시각', '-')} · {fmt_period(시점[0])}~{fmt_period(시점[-1])})\n")
    for label, block in data["지표"].items():
        print(f"■ {label} ({block['단위']})")
        for region, values in block["지역"].items():
            최근 = [f"{fmt_period(p)} {v if v is not None else '-'}" for p, v in list(zip(시점, values))[-4:]]
            print(f"  {region:6s} … {' | '.join(최근)}")
        print()
    print(f"출처: {data['출처']['이름']}\n      {data['출처']['url']}")
    print(f"저장: {CACHE_FILE}\n")
