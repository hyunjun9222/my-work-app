---
name: issues
description: Collect Gyeongbuk industry and vocational-training news with gb_issues.py (OpenAI web-search model), cached in data/. Use when the user asks for 경북 이슈, 산업 동향, or 훈련 관련 뉴스 정리.
---

# issues — 경북 산업·직업훈련 이슈 수집

## 무엇을

[gb_issues.py](../../../gb_issues.py) 로 경북 산업·직업훈련 이슈를 모은다. 결과는 `data/gb_issues.json` 에 캐시되고 화면에서는 `/issues` 로 본다.

## 입력

- 선택: 주제(`--topic 이차전지` 처럼), 새로 받을지 여부(`--refresh`).
- `.env` 의 `OPENAI_API_KEY`.

## 순서

1. 캐시부터 본다. 최근 자료로 충분하면 다시 부르지 않는다.
2. 새로 받아야 하면 `python gb_issues.py --refresh [--topic 주제]` 를 돌린다.
3. 각 항목의 제목·요약·링크를 확인한다. 링크는 응답 `annotations` 에서 온 URL만 쓴다.
4. 훈련 과정·기관 운영과 관련 있는 것만 골라 정리한다.

## 출력

이슈 목록을 제목·한 줄 요약·출처 링크로 전하고, 캐시를 쓴 것인지 새로 받은 것인지 밝힌다.

## 규칙

- 모델이 지어낸 주소를 링크로 올리지 않는다. 출처가 없으면 링크 없이 적는다.
- 기사에 없는 수치·배경을 덧붙이지 않는다.
- 외부 호출은 필요할 때만 한다. 같은 주제를 반복해서 부르지 않는다.
