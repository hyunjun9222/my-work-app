"""AI 이미지 생성 — OpenAI 이미지 API

짧은 설명을 주면 이미지를 만들어 images/ 폴더에 png 로 저장한다.

API 키는 이 폴더의 .env 파일 또는 환경변수 OPENAI_API_KEY 에서 읽는다.
코드에 직접 적지 않는다.

    .env 파일:   OPENAI_API_KEY=sk-...
    PowerShell:  $env:OPENAI_API_KEY = "sk-..."
    Git Bash  :  export OPENAI_API_KEY="sk-..."

사용법:
    python imagegen.py "내 서비스 마스코트 - 물방울 모양 귀여운 캐릭터"
    python imagegen.py "설명" --size 1024x1536 --n 2
    python imagegen.py "모자에 이 로고를 넣어줘" --ref inputs/logo.png
"""

import argparse
import base64
import os
import re
import sys
from datetime import datetime
from pathlib import Path

IMAGES_DIR = Path(__file__).parent / "images"
ENV_FILE = Path(__file__).parent / ".env"
MODEL = "gpt-image-1"
SIZES = ["1024x1024", "1024x1536", "1536x1024", "auto"]
QUALITIES = ["low", "medium", "high", "auto"]
MAX_N = 4
DEFAULT_N = 2  # 한 번에 두 장
REF_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp"}


class ImageGenError(Exception):
    """사용자에게 그대로 보여줄 수 있는 오류."""


def load_env(path=None):
    """.env 파일의 KEY=VALUE 를 환경변수로 올린다.

    이미 설정된 환경변수가 우선한다(.env 가 덮어쓰지 않는다).
    라이브러리 없이 동작하도록 직접 읽는다.
    """
    path = Path(path) if path else ENV_FILE
    if not path.exists():
        return {}
    loaded = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key, value = key.strip(), value.strip().strip("'\"")
        if not key or not value or value.startswith("여기에"):
            continue  # 안내 문구가 그대로 남아 있으면 건너뛴다
        loaded[key] = value
        os.environ.setdefault(key, value)
    return loaded


def api_key():
    load_env()
    return os.environ.get("OPENAI_API_KEY", "").strip()


def key_source():
    """키를 어디서 읽었는지 — 화면 안내용."""
    from_file = load_env().get("OPENAI_API_KEY")
    value = os.environ.get("OPENAI_API_KEY", "").strip()
    if not value:
        return None
    return ".env 파일" if value == from_file else "환경변수"


def have_key():
    return bool(api_key())


def slugify(text, limit=40):
    """설명을 파일명에 쓸 수 있게 다듬는다. 한글은 그대로 둔다."""
    s = re.sub(r"[^\w가-힣\s-]", "", str(text), flags=re.UNICODE)
    s = re.sub(r"[\s_]+", "-", s.strip()).strip("-")
    return (s[:limit].rstrip("-") or "image")


def unique_path(directory, stem, ext=".png"):
    """같은 이름이 있으면 -2, -3 을 붙인다. 기존 파일을 덮어쓰지 않는다."""
    p = directory / f"{stem}{ext}"
    i = 2
    while p.exists():
        p = directory / f"{stem}-{i}{ext}"
        i += 1
    return p


def generate(prompt, size="1024x1024", n=DEFAULT_N, quality="medium", out_dir=None, reference=None):
    """이미지를 만들어 저장하고 저장된 경로 목록을 돌려준다.

    reference 에 이미지 경로를 주면 그 그림을 재료로 삼아 편집·합성한다
    (예: 로고를 건네주고 "모자에 이 로고를 넣어줘").
    """
    prompt = (prompt or "").strip()
    if not prompt:
        raise ImageGenError("설명(프롬프트)이 비어 있습니다.")
    if size not in SIZES:
        raise ImageGenError(f"지원하지 않는 크기입니다: {size} (가능: {', '.join(SIZES)})")
    if quality not in QUALITIES:
        raise ImageGenError(f"지원하지 않는 품질입니다: {quality} (가능: {', '.join(QUALITIES)})")
    n = max(1, min(int(n), MAX_N))

    refs = [Path(r) for r in ([reference] if isinstance(reference, (str, Path)) else (reference or []))]
    for r in refs:
        if not r.is_file():
            raise ImageGenError(f"참고 이미지를 찾을 수 없습니다: {r}")
        if r.suffix.lower() not in REF_SUFFIXES:
            raise ImageGenError(f"참고 이미지는 png·jpg·webp 만 됩니다: {r.name}")

    if not have_key():
        raise ImageGenError(
            "OpenAI API 키가 없습니다. 이 폴더의 .env 파일을 열어 아래 한 줄을 채우세요.\n"
            "  OPENAI_API_KEY=sk-여기에-키를-붙여넣기\n"
            "  (저장한 뒤 서버를 다시 실행하면 적용됩니다)"
        )

    try:
        from openai import OpenAI
    except ImportError:
        raise ImageGenError("openai 라이브러리가 없습니다. `pip install openai` 로 설치하세요.")

    client = OpenAI()
    handles = []
    try:
        if refs:  # 참고 이미지가 있으면 편집(합성) 방식으로
            handles = [r.open("rb") for r in refs]
            resp = client.images.edit(
                model=MODEL,
                image=handles if len(handles) > 1 else handles[0],
                prompt=prompt,
                size=size,
                n=n,
                quality=quality,
            )
        else:
            resp = client.images.generate(model=MODEL, prompt=prompt, size=size, n=n, quality=quality)
    except Exception as ex:  # 인증·요금·정책 거부 등을 그대로 전달
        raise ImageGenError(f"이미지 생성에 실패했습니다 — {type(ex).__name__}: {ex}")
    finally:
        for h in handles:
            h.close()

    directory = Path(out_dir) if out_dir else IMAGES_DIR
    directory.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    slug = slugify(prompt)

    saved = []
    for i, item in enumerate(resp.data, 1):
        data = getattr(item, "b64_json", None)
        if not data:
            raise ImageGenError("응답에 이미지 데이터가 없습니다.")
        suffix = f"-{i}" if len(resp.data) > 1 else ""
        path = unique_path(directory, f"{stamp}-{slug}{suffix}")
        path.write_bytes(base64.b64decode(data))
        saved.append(path)

    # 어떤 설명으로 만든 이미지인지 함께 남긴다
    note = prompt + (f"  [참고 이미지: {', '.join(r.name for r in refs)}]" if refs else "")
    (directory / "prompts.log").open("a", encoding="utf-8").write(
        f"{datetime.now():%Y-%m-%d %H:%M}\t{', '.join(p.name for p in saved)}\t{note}\n"
    )
    return saved


def list_images(directory=None):
    """저장된 이미지를 최신순으로."""
    directory = Path(directory) if directory else IMAGES_DIR
    if not directory.exists():
        return []
    return sorted(directory.glob("*.png"), key=lambda p: p.stat().st_mtime, reverse=True)


def load_prompts(directory=None):
    """파일명 → 만들 때 쓴 설명."""
    directory = Path(directory) if directory else IMAGES_DIR
    log = directory / "prompts.log"
    out = {}
    if not log.exists():
        return out
    for line in log.read_text(encoding="utf-8").splitlines():
        parts = line.split("\t")
        if len(parts) == 3:
            for name in parts[1].split(", "):
                out[name.strip()] = parts[2]
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="OpenAI 이미지 생성 → images/ 폴더에 저장")
    ap.add_argument("prompt", nargs="+", help="만들고 싶은 이미지 설명")
    ap.add_argument("--size", default="1024x1024", choices=SIZES)
    ap.add_argument("--quality", default="medium", choices=QUALITIES)
    ap.add_argument("--n", type=int, default=DEFAULT_N, help=f"장수 (기본 {DEFAULT_N}, 최대 {MAX_N})")
    ap.add_argument("--ref", action="append", help="참고 이미지 경로 (로고 등). 여러 번 쓸 수 있음")
    args = ap.parse_args()

    try:
        paths = generate(" ".join(args.prompt), args.size, args.n, args.quality, reference=args.ref)
    except ImageGenError as ex:
        print(f"\n오류: {ex}\n")
        sys.exit(1)
    print("\n저장했습니다:")
    for p in paths:
        print(f"  {p}  ({p.stat().st_size // 1024} KB)")
    print()
