---
name: approve
description: Check submission and approval status for a week (제출·승인 현황판) and tell what is blocking report publication — 미제출, 반려 대상, 명단 외 기관. Use when the user asks who has not submitted or whether the weekly report can be published.
---

# approve — 제출·승인 현황 확인

## 무엇을

한 주차의 제출·승인 상태를 보고, 리포트를 확정할 수 있는지 판단할 재료를 준다. 화면으로는 `/admin` 이다.

## 입력

- 주차 키. 없으면 저장소 최신 주차를 쓴다.
- 대상 명단은 `storage.load_roster()` 에서 온다. 명단이 비어 있으면 제출한 기관이 대상이 된다.

## 순서

1. `storage.load_week(주차)` 로 제출 내역을 읽는다. 없으면 그대로 알린다.
2. `report_engine.submission_board(집계결과, 명단)` 으로 현황판을 만든다.
3. [app.py](../../../app.py) 의 `apply_approval()` 로 저장소의 승인·반려 상태를 덮어 반영한다.
4. 상태별로 센다 — 승인, 검토대기, 반려 대상, 미제출, 명단 외.

## 출력

승인 완료율(승인 수/대상 수)과 상태별 기관 목록을 전한다. 미제출·반려 대상이 하나라도 있으면 "리포트 발행 보류, 지금 수치는 승인 전 잠정 계산" 이라고 분명히 적는다.

## 규칙

- 승인·반려 처리는 사람이 한다. 대신 눌러 상태를 바꾸지 않는다.
- 같은 기관이 다시 제출하면 상태가 `제출` 로 초기화되어 이전 승인이 무효가 된다는 점을 감안한다.
- `data/` 의 주차 파일을 직접 고치지 않는다.
