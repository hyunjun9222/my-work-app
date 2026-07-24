---
name: stats
description: Fetch and summarize KOSIS employment statistics (고용률·실업률·생산가능인구) by region and month with kosis_stats.py, cached in data/. Use when the user asks about 고용 통계 or wants regional employment numbers.
---

# stats — 고용 통계 보기

## 무엇을

[kosis_stats.py](../../../kosis_stats.py) 로 KOSIS(국가통계포털) 행정구역별 경제활동인구에서 고용률·실업률·생산가능인구를 월별로 받는다. 결과는 `data/kosis_stats.json` 에 캐시되고 화면에서는 `/stats` 로 본다.

## 입력

- 선택: 기간(`--months 60`), 새로 받을지 여부(`--refresh`), 사용 가능한 지역·항목 확인(`--list`).
- `.env` 의 `KOSIS_API_KEY`.

## 순서

1. 캐시를 먼저 본다. 월 단위 통계라 자주 바뀌지 않는다.
2. 필요하면 `python kosis_stats.py --refresh [--months 60]` 를 돌린다.
3. 어떤 지역·항목을 볼지 정한다. 항목·지역은 응답의 `ITM_NM`/`C1_NM` 으로 고르며 코드를 하드코딩하지 않는다.
4. 최근 값과 흐름(증가·감소)을 확인한다.

## 출력

지역별 최근 고용률·실업률과 눈에 띄는 변화를 전한다. 어느 시점 기준인지 반드시 함께 적는다.

## 규칙

- 선 그래프를 그릴 때 색 4개를 돌려쓰지 않는다. 지역 선택은 4개까지다.
- 통계에 없는 기간·지역을 추정해 채우지 않는다.
- 수치는 받은 값 그대로 쓰고 반올림 기준을 밝힌다.
