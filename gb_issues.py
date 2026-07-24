"""경북 지역 이슈 수집 — OpenAI 웹 검색 모델

경상북도의 주요 산업(섹터)별 최신 이슈와 직업훈련·고용 이슈를
섹터마다 PER_TOPIC 건씩 모아 각 이슈에 출처 링크를 붙여 돌려준다.

지어낸 내용이 섞이지 않도록 웹 검색 결과(annotations)에서 온 URL만 링크로 쓴다.
결과는 data/gb_issues.json 에 저장하고, 버튼을 누를 때만 새로 수집한다.

사용법:
    python gb_issues.py              # 저장된 결과 보기 (없으면 수집)
    python gb_issues.py --refresh    # 새로 수집
    python gb_issues.py --topic 이차전지
"""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from imagegen import have_key, load_env

BASE = Path(__file__).parent
CACHE_FILE = BASE / "data" / "gb_issues.json"
MODEL = "gpt-4o-search-preview"
PER_TOPIC = 15  # 섹터(주제)마다 뽑을 이슈 수 — 뉴스에 많이 언급되는 것부터
REGION = {"type": "approximate", "approximate": {"country": "KR", "region": "Gyeongsangbuk-do", "city": "Pohang"}}

# 경북 주력 산업 + 직업훈련. (주제, 검색 힌트)
TOPICS = [
    ("이차전지", "포항·구미 이차전지 소재·양극재 투자, 공장 증설, 특화단지"),
    ("반도체·전자", "구미 반도체 소재·부품, 전자산업 투자와 고용"),
    ("철강·금속", "포항 철강산업 구조전환, 수소환원제철, 업황"),
    ("자동차부품", "경북 자동차부품 기업 전동화 전환, 미래차"),
    ("바이오·백신", "안동 바이오·백신 산업, 세포배양, 제약"),
    ("관광·문화", "경주·안동 관광 산업, 문화관광 정책"),
    ("직업훈련·고용", "경상북도 직업훈련 사업, 폴리텍·훈련기관, 인력양성, 채용"),
]


class IssueError(Exception):
    """사용자에게 그대로 보여줄 수 있는 오류."""


def _client():
    load_env()
    if not have_key():
        raise IssueError(
            "OpenAI API 키가 없습니다. 이 폴더의 .env 파일에 OPENAI_API_KEY 를 채우세요.\n"
            "(이미지 생성·메모 분류와 같은 키를 씁니다)"
        )
    try:
        from openai import OpenAI
    except ImportError:
        raise IssueError("openai 라이브러리가 없습니다. `pip install openai` 로 설치하세요.")
    return OpenAI()


# ── 응답 → 이슈 목록 ─────────────────────────────────────────────

CITE_RE = re.compile(r"\s*\(\[[^\]]+\]\([^)]+\)\)")  # 본문에 박히는 ([제목](url)) 표기


def split_issues(content, annotations):
    """본문을 줄 단위 이슈로 나누고, 각 줄 위치에 걸린 출처를 붙인다."""
    issues = []
    for m in re.finditer(r"^\s*(?:[-*•]|\d+[.)])\s*(.+)$", content, re.MULTILINE):
        issues.append({"_start": m.start(1), "_end": m.end(1), "내용": m.group(1).strip(), "출처": []})

    if not issues:  # 목록 형태가 아니면 문단 전체를 한 건으로
        issues = [{"_start": 0, "_end": len(content), "내용": content.strip(), "출처": []}]

    for a in annotations or []:
        cite = getattr(a, "url_citation", None)
        if not cite:
            continue
        # 인용 표기는 해당 줄 끝(때로는 줄바꿈 뒤)에 붙는다.
        # 위치를 넘어서지 않는 가장 마지막 이슈에 붙여야 한 칸씩 밀리지 않는다.
        pos = getattr(cite, "start_index", None)
        target = None
        if pos is not None:
            for it in issues:
                if it["_start"] <= pos:
                    target = it
                else:
                    break
        target = target or issues[-1]
        url = cite.url
        if not any(s["url"] == url for s in target["출처"]):
            target["출처"].append({"제목": (cite.title or url)[:120], "url": url})

    for it in issues:
        it["내용"] = CITE_RE.sub("", it["내용"]).strip()  # 본문 속 인용 표기는 링크로 대체하므로 제거
        del it["_start"], it["_end"]
    return [it for it in issues if it["내용"]]


def fetch_topic(client, topic, hint, count=PER_TOPIC):
    """한 주제(섹터)의 최신 이슈를 검색해 이슈 목록으로."""
    prompt = (
        f"경상북도의 '{topic}' 분야에서 최근 뉴스에 자주 언급되는 핵심 이슈를 {count}건 찾아 주세요.\n"
        f"참고 키워드: {hint}\n\n"
        "조건:\n"
        "- 여러 매체에서 반복해 다뤄진 사안을 먼저 쓰고, 단발성 보도는 뒤로 미루세요.\n"
        f"- 같은 사안을 표현만 바꿔 여러 줄로 쓰지 말고, 서로 다른 사안 {count}건으로 채우세요.\n"
        "- 각 이슈를 '- ' 로 시작하는 한 줄로 쓰세요. 한 줄에 한 이슈만.\n"
        "- 각 줄은 무엇이 언제 일어났는지 한국어 두 문장 이내로 구체적으로 씁니다.\n"
        "- 검색으로 확인한 내용만 쓰고, 확인되지 않은 수치나 계획은 쓰지 마세요.\n"
        "- 최근 것을 우선하고, 경상북도와 직접 관련된 것만 고르세요.\n"
        f"- 확인된 사안이 {count}건에 못 미치면 확인된 것만 쓰고 억지로 채우지 마세요.\n"
        "- 머리말이나 맺음말 없이 목록만 출력하세요."
    )
    try:
        resp = client.chat.completions.create(
            model=MODEL,
            web_search_options={"user_location": REGION},
            messages=[{"role": "user", "content": prompt}],
        )
    except Exception as ex:
        raise IssueError(f"'{topic}' 검색에 실패했습니다 — {type(ex).__name__}: {ex}")

    msg = resp.choices[0].message
    return split_issues(msg.content or "", getattr(msg, "annotations", None))


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


def collect(topics=None, progress=None):
    """주제별로 수집해 저장한다. topics 를 주면 그 주제만 갱신."""
    client = _client()
    cache = load_cache()
    targets = [t for t in TOPICS if not topics or t[0] in topics]
    if not targets:
        raise IssueError(f"알 수 없는 주제입니다: {', '.join(topics)}")

    for topic, hint in targets:
        if progress:
            progress(topic)
        cache[topic] = {
            "주제": topic,
            "이슈": fetch_topic(client, topic, hint),
            "수집시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "모델": MODEL,
        }
    save_cache(cache)
    return cache


def ordered(cache):
    """TOPICS 순서대로 (저장된 것만)."""
    return [cache[t] for t, _ in TOPICS if t in cache]


def last_updated(cache):
    times = [v.get("수집시각") for v in cache.values() if v.get("수집시각")]
    return max(times) if times else None


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="경북 산업·직업훈련 최신 이슈 수집")
    ap.add_argument("--refresh", action="store_true", help="새로 수집")
    ap.add_argument("--topic", action="append", help="특정 주제만 (여러 번 사용 가능)")
    args = ap.parse_args()

    cache = load_cache()
    if args.refresh or args.topic or not cache:
        try:
            cache = collect(args.topic, progress=lambda t: print(f"  검색 중… {t}"))
        except IssueError as ex:
            print(f"\n오류: {ex}\n")
            sys.exit(1)

    print(f"\n경상북도 산업·직업훈련 이슈  (최종 수집 {last_updated(cache) or '-'})\n")
    for block in ordered(cache):
        print(f"■ {block['주제']}  ({len(block['이슈'])}건 · {block['수집시각']})")
        for it in block["이슈"]:
            print(f"  · {it['내용']}")
            for s in it["출처"]:
                print(f"      ↳ {s['제목']}")
                print(f"        {s['url']}")
        print()
    print(f"저장: {CACHE_FILE}\n")
