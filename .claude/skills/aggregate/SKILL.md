---
name: aggregate
description: Validate and aggregate a training-performance Excel file with feature1_aggregate.py, then report totals, rates, and input errors. Use when the user hands over a 실적 엑셀 and asks to check, validate, or aggregate it (집계, 검증, 이수율, 탈락률).
---

# aggregate — 실적 엑셀 집계·검증

## 무엇을

훈련기관 실적 엑셀을 [feature1_aggregate.py](../../../feature1_aggregate.py) 로 돌려 검증·표준화·지표·합계를 뽑는다. 이수율·탈락률은 여기서만 계산하고 다른 곳에서 다시 계산하지 않는다.

## 입력

- 엑셀 경로 하나. 지시에 없으면 `inputs/` 의 샘플을 쓰고, 어느 파일인지 애매하면 물어본다.
- 시트 `실적` 과 8개 컬럼(`기관명`·`과정명`·`NCS분류`·`KECO분류`·`훈련목표인원`·`훈련실시인원`·`훈련수료인원`·`중도탈락자`)이 필요하다.

## 순서

1. `PYTHONIOENCODING=utf-8 python feature1_aggregate.py <엑셀경로>` 로 돌린다. 콘솔이 cp949 라 이 환경변수가 없으면 한글에서 깨진다.
2. 시트·컬럼이 없으면 `SystemExit` 이 난다. 무엇이 없는지 그대로 사용자에게 전한다.
3. 전체 합계와 구분별(기관·NCS·KECO) 합계를 확인한다.
4. 오류 목록을 행 번호·컬럼·사유까지 확인한다. 오류 행은 합계에서 빠지고 지표는 `검증 필요` 로 남는다.

## 출력

답변으로 전체 이수율·탈락률, 정상/제외 과정 수, 오류 목록을 전한다. 파일은 만들지 않는다.

## 규칙

- 표의 행 순서와 구분 순서는 입력에 나온 순서 그대로 둔다. 정렬하지 않는다.
- 숫자로 읽지 못한 값은 원본 그대로 전한다. 임의로 고치거나 0으로 바꾸지 않는다.
- 결과 수치는 실행해서 나온 값만 적는다.
