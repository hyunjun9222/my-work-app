---
name: todos
description: Extract only the action items (오늘 할 일) from a weekly report — 미제출 독촉, 재입력 요청, 이상치 확인, 특이사항 확인, 오늘 일정 — sorted by priority. Use when the user asks what to handle today or wants the to-do list instead of the whole report.
---

# todos — 오늘 할 일만 뽑기

## 무엇을

리포트 결과에서 조치가 필요한 항목만 추린다. `report_engine.today_todos()` 가 만드는 목록이며, 별도 탭은 없애서, 주간 리포트(`/report`) 맨 위 섹션으로만 나온다.

## 입력

- 주차 키 또는 엑셀 경로. 없으면 저장소의 최신 주차, 저장소가 비어 있으면 `inputs/` 의 샘플을 쓴다.
- 선택: 기준일. '오늘 일정' 항목이 이 날짜로 잡힌다.

## 순서

1. [app.py](../../../app.py) 의 `week_report(week, today, src)` 로 리포트를 만든다. 저장소에 제출이 없으면 샘플 엑셀로 자동 전환된다.
2. `rep["오늘할일"]` 을 읽는다. 각 항목은 `우선순위`·`구분`·`대상`·`할일` 네 값이다.
3. 우선순위(높음·보통·낮음)별로 몇 건인지 센다.
4. 사용자가 특정 우선순위만 원하면 그 값으로 거른다.

## 출력

높음부터 순서대로 나열한다. 한 줄에 대상 기관과 할 일 하나씩이고, 맨 아래에 전체 건수와 높음 건수를 적는다. 화면에서 보려면 `/report?week=<주차>` 맨 위 '오늘 할 일' 카드를 안내한다.

## 규칙

- 여기서 새로 판단하지 않는다. 이미 계산된 상태값을 행동 문장으로 옮기기만 한다.
- 승인 상태가 반영된 리포트를 써야 한다. 승인 완료 기관이 "승인 검토" 로 남아 있으면 `apply_approval()` 을 거치지 않은 것이다.
- 목록이 비면 비었다고 그대로 말한다. 채워 넣지 않는다.
