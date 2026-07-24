"""메모 자동 분류 — OpenAI 챗 API

practice/memos/ 폴더의 메모를 읽어 각 메모에 어울리는 카테고리를 AI가 붙인다.

API 키는 imagegen 과 같은 곳(.env 의 OPENAI_API_KEY)에서 읽는다.
결과는 data/memo_tags.json 에 저장하고, 내용이 바뀌지 않은 메모는
다시 호출하지 않는다(같은 메모를 반복 분류해 요금이 새지 않게).

사용법:
    python memo_tagger.py            # 새 메모·수정된 메모만 분류
    python memo_tagger.py --all      # 전부 다시 분류
"""

import argparse
import hashlib
import json
import sys
from datetime import datetime
from pathlib import Path

from imagegen import ImageGenError, have_key, load_env

BASE = Path(__file__).parent
MEMOS_DIR = BASE / "practice" / "memos"  # 연습용 메모는 practice/ 로 옮겼다
CACHE_FILE = BASE / "data" / "memo_tags.json"
MODEL = "gpt-5-mini"
MAX_CHARS = 6000  # 메모가 길면 앞부분만 보낸다

# AI 가 고를 기본 카테고리. 어디에도 안 맞으면 새로 지어낼 수 있게 열어 둔다.
CATEGORIES = ["회의록", "고객 문의", "아이디어", "할 일", "장애·이슈", "계약·비용", "자료 조사", "기타"]

SYSTEM = (
    "당신은 업무 메모를 분류하는 사서입니다. "
    "메모 전체를 읽고 그 메모의 성격에 가장 잘 맞는 카테고리 하나를 고릅니다. "
    "주어진 후보 중에 맞는 것이 없을 때만 새 카테고리를 짧은 한국어 명사구로 만들어 씁니다. "
    "메모에 적혀 있지 않은 내용은 지어내지 않습니다."
)

SCHEMA = {
    "type": "json_schema",
    "json_schema": {
        "name": "memo_category",
        "strict": True,
        "schema": {
            "type": "object",
            "properties": {
                "카테고리": {"type": "string", "description": "메모의 성격을 나타내는 카테고리 하나"},
                "후보에_있음": {"type": "boolean", "description": "주어진 후보 목록에서 골랐으면 true"},
                "태그": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "메모의 주제를 나타내는 키워드 2~4개",
                },
                "한줄요약": {"type": "string", "description": "메모 내용을 한 문장으로"},
                "확신도": {"type": "string", "enum": ["높음", "보통", "낮음"]},
                "근거": {"type": "string", "description": "그 카테고리로 본 이유를 짧게"},
            },
            "required": ["카테고리", "후보에_있음", "태그", "한줄요약", "확신도", "근거"],
            "additionalProperties": False,
        },
    },
}


class TagError(Exception):
    """사용자에게 그대로 보여줄 수 있는 오류."""


# ── 메모 읽기 ────────────────────────────────────────────────────


def list_memos(directory=None):
    d = Path(directory) if directory else MEMOS_DIR
    if not d.exists():
        return []
    return sorted(p for p in d.glob("*.md") if p.is_file())


def digest(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]


# ── 캐시 ─────────────────────────────────────────────────────────


def load_cache():
    if not CACHE_FILE.exists():
        return {}
    try:
        return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_cache(cache):
    CACHE_FILE.parent.mkdir(exist_ok=True)
    CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


# ── 분류 ─────────────────────────────────────────────────────────


def _client():
    load_env()
    if not have_key():
        raise TagError(
            "OpenAI API 키가 없습니다. 이 폴더의 .env 파일에 OPENAI_API_KEY 를 채우세요.\n"
            "(이미지 생성 기능과 같은 키를 씁니다)"
        )
    try:
        from openai import OpenAI
    except ImportError:
        raise TagError("openai 라이브러리가 없습니다. `pip install openai` 로 설치하세요.")
    return OpenAI()


def classify_one(client, name, text):
    """메모 한 건을 분류한다."""
    body = text[:MAX_CHARS]
    user = (
        f"카테고리 후보: {', '.join(CATEGORIES)}\n\n"
        f"아래는 '{name}' 파일의 내용입니다.\n\n----\n{body}\n----"
    )
    try:
        resp = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "system", "content": SYSTEM}, {"role": "user", "content": user}],
            response_format=SCHEMA,
        )
    except Exception as ex:
        raise TagError(f"분류 요청에 실패했습니다 — {type(ex).__name__}: {ex}")

    try:
        return json.loads(resp.choices[0].message.content)
    except (json.JSONDecodeError, AttributeError, IndexError) as ex:
        raise TagError(f"응답을 읽지 못했습니다 — {type(ex).__name__}: {ex}")


def tag_memos(directory=None, force=False, progress=None):
    """메모 폴더 전체를 분류한다. (결과목록, 새로_호출한_건수) 반환."""
    files = list_memos(directory)
    if not files:
        raise TagError(f"분류할 메모가 없습니다. {MEMOS_DIR.name}/ 폴더에 .md 파일을 넣어 주세요.")

    cache = load_cache()
    results, called = [], 0
    client = None

    for p in files:
        text = p.read_text(encoding="utf-8")
        h = digest(text)
        hit = cache.get(p.name)

        if hit and hit.get("_해시") == h and not force:
            results.append(hit)
            continue

        if client is None:
            client = _client()
        if progress:
            progress(p.name)

        item = classify_one(client, p.name, text)
        item["_파일"] = p.name
        item["_해시"] = h
        item["_분류시각"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        item["_모델"] = MODEL
        cache[p.name] = item
        results.append(item)
        called += 1

    # 사라진 메모는 캐시에서 정리
    names = {p.name for p in files}
    for stale in [k for k in cache if k not in names]:
        del cache[stale]

    save_cache(cache)
    return results, called


def grouped(results):
    """카테고리별로 묶어 (카테고리, 항목들) 목록으로."""
    order, groups = [], {}
    for r in results:
        c = r["카테고리"]
        if c not in order:
            order.append(c)
        groups.setdefault(c, []).append(r)
    return [(c, groups[c]) for c in sorted(order, key=lambda c: (-len(groups[c]), c))]


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="memos/ 폴더의 메모에 AI가 카테고리를 붙입니다")
    ap.add_argument("--all", action="store_true", help="캐시를 무시하고 전부 다시 분류")
    ap.add_argument("--dir", help="메모 폴더 (기본 memos/)")
    args = ap.parse_args()

    try:
        results, called = tag_memos(args.dir, force=args.all, progress=lambda n: print(f"  분류 중… {n}"))
    except TagError as ex:
        print(f"\n오류: {ex}\n")
        sys.exit(1)

    print(f"\n메모 {len(results)}건 (새로 분류 {called}건 / 캐시 {len(results) - called}건) · 모델 {MODEL}\n")
    for cat, items in grouped(results):
        print(f"■ {cat}  ({len(items)}건)")
        for r in items:
            mark = "" if r["확신도"] == "높음" else f"  [확신도 {r['확신도']}]"
            print(f"  · {r['_파일']}{mark}")
            print(f"      {r['한줄요약']}")
            print(f"      태그: {', '.join(r['태그'])}")
        print()
    print(f"결과 저장: {CACHE_FILE}\n")
