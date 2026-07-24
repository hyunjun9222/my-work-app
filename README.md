# 주간 훈련기관 교육실적 취합 App

훈련기관이 주차별 교육실적(정부 「지산맞 훈련실적」 양식)을 화면 입력 또는 엑셀 업로드로 내면,
관리자가 승인하고 주간 리포트·결과표 엑셀을 뽑는 앱입니다. 집계·검증·지표는 모두
[feature1_aggregate.py](feature1_aggregate.py) 한 곳에서 계산합니다(단일 진실 공급원).

## 두 개의 실행 진입점

이 저장소에는 성격이 다른 두 화면이 있습니다.

| 진입점 | 설명 | 실행 |
|---|---|---|
| **`app.py`** | 로그인·업로드·승인·리포트·챗봇이 있는 **본 포털** (표준 라이브러리 `http.server` 기반) | `python app.py` → http://localhost:8000 |
| **`streamlit_app.py`** | 실적을 **월별·연도별로 보는 읽기 전용 화면** (Streamlit) | `streamlit run streamlit_app.py` → http://localhost:8501 |

> **배포 참고 —** Streamlit Community Cloud / Codespaces 는 `streamlit run streamlit_app.py` 로
> **Streamlit 화면만** 띄웁니다(본 포털 `app.py` 는 별도 호스트가 필요합니다). 배포 환경에는
> `data/` 가 없으므로, 화면은 함께 담긴 **`demo_data/` 의 예시(가짜) 데이터**로 자동 채워집니다.
> 로컬에서 `python app.py` 로 실제 제출이 `data/` 에 쌓이면 그 실제 자료를 씁니다.

## 설치

```bash
pip install -r requirements.txt
```

의존성은 `openpyxl`(엑셀), `streamlit`·`pandas`(월별·연도별 화면), `openai`(챗봇·이미지 등 AI 기능, 지연 import) 4개뿐입니다.

## 첫 실행

`python app.py` 를 처음 실행하면 `admin` 계정과 임시 비밀번호가 콘솔에 한 번 출력됩니다.
`data/` 는 사람이 열어 고칠 수 있는 JSON 저장소이며, API 키는 `.env`(비공개)에 두고 `.env.example` 을 서식으로 씁니다.

## 폴더

- `feature1_aggregate.py` — 검증·표준화·지표·합계 (양성/향상 현황 시트 + 교육실적 목표)
- `report_engine.py` — 취합·전주 비교·이상치·요약
- `storage.py` / `calendar_store.py` — JSON 저장소
- `app.py` — 포털 화면과 라우팅 · `streamlit_app.py` — 월별·연도별 보기
- `demo_data/` — 개인정보 없는 예시 데이터(배포 화면용)
- `specs/` — 요구사항 원본
