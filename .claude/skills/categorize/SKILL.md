---
name: categorize
description: Classify memo files in practice/memos/ with memo_tagger.py, caching results in data/memo_tags.json. Use when the user asks to sort, tag, or categorize 메모.
---

# categorize — 메모 분류

## 무엇을

[memo_tagger.py](../../../memo_tagger.py) 로 `practice/memos/` 의 메모를 분류·태깅한다. 결과는 `data/memo_tags.json` 에 캐시된다. 앱 화면은 없애서 CLI 로만 쓴다.

## 입력

- `practice/memos/` 폴더의 메모 파일들. 새 메모만 처리하는 것이 기본이다.
- 선택: 전체 재분류 여부(`--all`).
- `.env` 의 `OPENAI_API_KEY`.

## 순서

1. `practice/memos/` 에 파일이 있는지 확인한다. 비어 있으면 실행하지 않고 알린다.
2. `python memo_tagger.py` 를 돌린다. 전체를 다시 분류해야 할 때만 `--all` 을 붙인다.
3. 내용 해시로 캐시하므로 바뀌지 않은 메모는 다시 호출되지 않는다. 재실행 전에 이 점을 감안한다.
4. 결과에서 분류·태그·확신도와, 후보에 없던 새 분류가 생겼는지 확인한다.

## 출력

처리한 메모 수, 분류별 건수, 새로 생긴 분류, 확신도 낮은 항목을 전한다.

## 규칙

- `data/memo_tags.json` 을 손으로 고치지 않는다. 다시 만들려면 스크립트로 돌린다.
- 메모 원본은 읽기만 한다. 수정하지 않는다.
- 분류가 애매한 메모는 임의로 정하지 말고 확신도 낮음으로 남겨 사용자에게 알린다.
