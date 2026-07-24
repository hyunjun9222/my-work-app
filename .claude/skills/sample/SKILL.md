---
name: sample
description: Create a fake practice workbook (실적/특이사항/주요일정 sheets) under inputs/ for a given week, optionally with deliberate input errors, then run it through the aggregator. Use when the user wants 연습용 샘플 데이터 to test screens or rules.
---

# sample — 연습용 샘플 주차 파일 만들기

## 무엇을

가짜 데이터로 주차 엑셀 하나를 만들어 화면·집계를 연습한다. 실제 기관명·사내 자료는 절대 쓰지 않는다.

## 입력

- 주차 키(예: `2026-W31`). 없으면 물어본다.
- 선택: 기관 수·과정 수, 일부러 넣을 오류 건수, 일정 날짜.

## 순서

1. 기존 샘플 [inputs/sample-training-data-2026-W30.xlsx](../../../inputs/sample-training-data-2026-W30.xlsx) 의 시트·컬럼 구성을 확인한다.
2. 생성 스크립트를 스크래치패드에 쓴다. 프로젝트 폴더에 임시 스크립트를 남기지 않는다.
3. 세 시트를 채운다.
   - `실적` — 기관명·과정명·NCS분류·KECO분류·훈련목표인원·훈련실시인원·훈련수료인원·중도탈락자.
   - `특이사항` — 기관명·과정명·분류·내용·확인필요(Y/N).
   - `주요일정` — 날짜(YYYY-MM-DD)·기관명·구분·내용.
4. 검증을 확인하려면 오류를 일부러 섞는다. 예: 숫자 자리에 한글, 필수값 빈칸, 수료 인원이 실시 인원보다 많음.
5. `inputs/sample-training-data-<주차>.xlsx` 로 저장한다. 같은 이름이 이미 있으면 덮어쓰기 전에 물어본다.
6. `PYTHONIOENCODING=utf-8 python feature1_aggregate.py <저장경로>` 로 돌려 결과와 오류 목록을 확인한다.

## 출력

만든 파일 경로와, 돌려서 나온 합계·이수율·탈락률·오류 목록을 함께 전한다.

## 규칙

- 기관명·강사명·과정명은 모두 지어낸 이름을 쓴다. 실명 금지.
- `data/` 에는 쓰지 않는다. 샘플은 `inputs/` 에만 둔다.
- 오류를 섞었으면 어떤 행에 무엇을 넣었는지 사용자에게 밝힌다.
