---
name: report
description: Build the full weekly report dict with report_engine.build_report and summarize it (totals, week-over-week, outliers, notes, schedule, draft summary). Use when the user asks for 주간 리포트 for a given week or excel file.
---

# report — 주간 리포트 만들기

## 무엇을

[report_engine.py](../../../report_engine.py) 의 `build_report()` 로 한 주차의 리포트 전체를 만든다. 집계 위에 취합·전주 대비 비교·작년 동기 이상치·특이사항·일정·요약초안을 얹은 결과다.

## 입력

- 주차 키(`2026-W31` 형식) 또는 엑셀 경로 하나.
- 선택: 기준일(`today`). 없으면 오늘 날짜를 쓴다.

## 순서

1. 출처를 정한다. 저장소면 `storage.load_week(주차)`, 파일이면 엑셀 경로를 그대로 넘긴다.
2. `build_report(this, last, ly, roster, week, today=기준일)` 로 만든다. 전주·작년 자료는 `storage.prev_week_key()` / `last_year_key()` 로 찾는다.
3. 저장소 자료면 [app.py](../../../app.py) 의 `apply_approval(rep, data)` 로 승인·반려 상태를 반영한다. 이걸 빼면 승인된 기관이 계속 "승인 검토" 로 남는다.
4. 결과 dict 에서 `합계_전체`, `비교`, `이상치`, `특이사항`, `주요일정`, `요약초안`, `오늘할일` 을 확인한다.

## 출력

답변으로 전체 이수율·탈락률, 전주 대비 증감, 이상치 건수, 확인 필요 특이사항 수, 요약초안을 전한다. 엑셀 파일이 필요하면 화면의 `/export?week=<주차>&kind=report` 를 안내한다.

## 규칙

- `비교` 와 `이상치` 는 비교 자료가 없으면 `None` 이다. 없는 것을 있다고 쓰지 말고 "자료 없음" 으로 적는다.
- 지표를 여기서 다시 계산하지 않는다. 값은 결과 dict 에 있는 것만 옮긴다.
- `data/` 를 지우거나 덮어쓰지 않는다.
