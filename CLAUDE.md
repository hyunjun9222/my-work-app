# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 이 저장소는

주간 훈련기관 교육실적 취합 앱. 훈련기관이 주차별 실적을 화면 입력 또는 엑셀 업로드로 내면, 관리자가 승인하고 주간 리포트·엑셀을 뽑는다. 도메인 용어·주석·화면 문구가 전부 한국어이고, dict 키도 한국어(`"훈련수료인원"`, `"합계_기관별"`)다 — 새 코드도 그 규칙을 따른다.

`specs/` 가 사실상의 요구사항 원본이다. 코드 주석 곳곳의 "흐름도 N단계" 는 [specs/자동화_흐름도.md](specs/자동화_흐름도.md) 의 단계 번호를, `feature1_aggregate.py` 는 [specs/feature-1-spec.md](specs/feature-1-spec.md) 를 가리킨다. 집계·검증 동작을 고칠 때는 해당 명세를 먼저 읽고, 명세와 어긋나게 바꾸려면 명세도 같이 고친다.

## 실행

```powershell
python app.py            # http://localhost:8000 (브라우저 자동 실행)
python app.py 9000       # 포트 지정
python app.py --public   # 외부 접속 허용(0.0.0.0 수신) — 절차·주의는 remote-access.md

streamlit run streamlit_app.py            # http://localhost:8501 — 월별·연도별 실적 보기(읽기 전용)

python feature1_aggregate.py [엑셀경로]   # 정부 「지산맞」 양식 집계 결과를 터미널 표로 확인 (기본 sample/2026년 A기관.xlsx)
python imagegen.py "설명" [--size 1024x1536] [--n 2] [--ref inputs/logo.png]
python memo_tagger.py [--all]             # practice/memos/ 분류, 결과 data/memo_tags.json 에 캐시
python gb_issues.py [--refresh] [--topic 이차전지]      # 경북 산업·직업훈련 이슈 수집
python kosis_stats.py [--refresh] [--months 60] [--list]  # KOSIS 고용 통계 수집
```

의존성은 `openpyxl` (필수) 과 `openai` (imagegen/memo_tagger 에서만, 지연 import). requirements 파일·테스트 스위트·빌드 단계는 없다. 변경 검증은 `python feature1_aggregate.py` 로 집계 로직을, `python app.py` 로 화면을 직접 돌려서 한다.

첫 실행 시 `auth.ensure_admin()` 이 `admin` 계정과 임시 비밀번호를 만들어 콘솔에 한 번만 출력한다. `data/users.json` 을 지우면 다시 발급된다.

## 구조

의존 방향은 한 줄이다: `app.py` → `report_engine.py` → `feature1_aggregate.py`, 그리고 `storage.py` 는 모두가 쓴다.

- **[feature1_aggregate.py](feature1_aggregate.py)** — 검증·표준화·지표·합계의 **단일 진실 공급원**. 다른 어디에서도 실시율/수료율/탈락률/취업률을 다시 계산하지 않는다. 입력은 정부 「지산맞 훈련실적」 양식이다.
  - `read_rows(path)` = 엑셀 읽기(1단계)만, `process_rows(rows)` = 검증~합계(2~6단계). 엑셀 업로드든 화면 직접 입력이든 같은 규칙을 통과시키려고 쪼개 놓았다. 새 입력 경로는 행 목록(`_행`·`구분` + `TEXT_COLS`/`COUNT_COLS` 키 dict)으로 만들어 `process_rows` 에 넣는다.
  - **과정 행은 「양성훈련 현황」·「향상훈련 현황」 두 시트**에서 읽는다. 머리글을 이름으로 찾으므로(`HEADER_MAP`, 정규화) 정부 원본의 열 순서·간소 양식 둘 다 읽힌다. 각 행에 `구분`(양성/향상)이 붙고, `정기수시=='수시'` 면 집계 버킷은 `수시`.
  - **실시율·수료율(목표 대비)의 분모 = 「교육실적」 시트의 기관 연간 목표**(`read_targets`, 정기채용=양성/정기재직=향상/수시). 이게 정부 총계와 정확히 맞는 지점이다(정원 합이 아니다). 목표가 없으면(직접 입력) 과정 정원(`훈련목표인원`) 합으로 대체. 목표는 각 행에 `목표_총/양성/향상/수시` 로 실려 다니고 `storage`·`to_rows` 가 보존한다.
  - 지표: `실시율`=실시÷목표, `수료율`=수료÷실시(실시 대비), `수료율_목표`=수료÷목표, `탈락률`=중도탈락÷실시, `취업률`=취업÷수료(양성만, 향상은 `해당없음`). `모집률`=실시÷정원.
  - 반환 dict: `표`(오류 행 포함), `합계_전체`/`합계_기관별`/`합계_구분별`(양성/향상/수시)/`합계_기관구분별`/`합계_NCS별`/`합계_KECO별`, `오류`, `행수`. 각 합계 행에 `목표훈련인원`(그 그룹 목표)과 위 지표들이 들어 있다.
  - 지표 값은 float | None | 문자열 상수(`검증 필요`=오류 행, `계산 불가`=분모 0, `해당없음`=취업률 등)다. 소비하는 쪽은 반드시 `isinstance` 로 갈라야 한다 (`app.pct()`, `app.bars()`, `report_engine.flag_outliers()` 참고).
  - 오류 행은 표에 남기되 지표를 계산하지 않고 합계에서 뺀다. 빠진 수는 구분마다 `제외` 로 따라다닌다. 리포트 헤드라인 지표는 **실시율·수료율(실시 대비)·탈락률**, 이상치·전주 비교는 **수료율(실시 대비)** 기준이다.
  - `read_rows` 는 양성·향상 현황 시트가 하나도 없으면 `SystemExit` 을 던진다 — 웹에서는 `app.py` 업로드 핸들러가 잡아 에러 페이지로 바꾼다.
- **[storage.py](storage.py)** — `data/{연}-W{주차}.json` 주차 파일 + `roster.json`(대상 명단) + `users.json`. DB 없음, 사람이 열어보고 고칠 수 있는 JSON이 의도된 설계다. 주차 키는 `2026-W30` 형식이며 `week_key`/`parse_week`/`prev_week_key`/`last_year_key` 로만 다룬다. `to_rows`/`to_notes`/`to_plans` 가 저장 형식 → 집계 입력 어댑터.
  - 같은 기관이 다시 제출하면 덮어쓰고 상태가 `제출` 로 초기화된다(승인 무효화). 상태는 `제출`/`승인`/`반려`.
- **[report_engine.py](report_engine.py)** — feature1 위에 취합·비교·이상치·요약을 얹는다. `build_report(...)` 하나가 리포트 전체 dict를 만든다. `open_source()` 가 "엑셀 경로 | 저장소 dict" 양쪽을 받아 같은 집계로 정규화하는 지점이다. 전주 대비 비교(9단계)와 작년 동기 이상치(11단계, `OUTLIER_THRESHOLD` 10%p)는 비교 자료가 없으면 `None` 을 돌려주므로 화면·엑셀 쪽에서 항상 None 분기를 둔다. 작년 과정 매칭은 `course_key()` 로 과정명 끝의 "N기" 를 떼어 맞춘다.
  - `today_todos()` 는 리포트의 다른 결과(현황판·이상치·특이사항·일정)에서 파생되므로 `build_report` 안에서 **맨 마지막에** 계산한다. 여기서 새로 판단하지 않고 이미 정해진 상태값을 행동 문장으로 옮기기만 한다.
- **[app.py](app.py)** — 프레임워크 없이 `http.server` + f-string HTML. 단일 파일에 `*_page()` 렌더 함수들과 `Handler.do_GET`/`do_POST` 라우팅이 들어 있다. 화면 추가는 `xxx_page()` 함수 + 라우트 분기 한 줄로 한다.
  - 권한: `do_GET` 의 `admin_only` 집합과 `do_POST` 의 대응 검사에서 걸러진다. 기관 계정은 업로드 시 자기 기관 행만 저장된다.
  - 사용자 입력을 HTML에 넣을 때는 반드시 `e()` (html.escape) 를 거친다.
  - `template_xlsx()` (`/template`) 가 기관 배포용 **간소** 빈 양식을 만든다(양성/향상 현황 + 교육실적 목표 + 작성 방법). 정부 원본이 있으면 그대로 올리면 되고, 이 양식은 없을 때만 쓴다. 머리글 이름은 리더(`HEADER_MAP`)가 찾는 정부 이름 그대로여야 한다. 정부 원본의 NCS/KECO 참조 시트(수천 행)는 재현하지 않는다.
  - `trend_page()` (`/trend`, 관리자 전용) 가 저장소 주차를 (연,월)로 묶어 월별·연도별 수료율(목표 대비)을 보여준다 — `streamlit_app.py` 와 같은 계산을 포털 안에 넣은 것이라, 포털만 배포해도 이 화면이 함께 간다(streamlit 은 이제 선택).
  - 배포: `python app.py` 는 환경변수 `PORT` 가 있으면(Render·Railway 등) 그 포트로 `0.0.0.0` 바인딩·브라우저 미실행(클라우드 모드). 시작 명령은 `Procfile`(`web: python app.py`). 키는 그 호스트의 환경변수 `OPENAI_API_KEY`/`KOSIS_API_KEY` 로 넣으면 코드가 그대로 읽는다. **저장 위치는 `storage.DATA_DIR` = 환경변수 `DATA_DIR`(있으면) 아니면 `data/`** — 배포 호스트에서 영구 디스크를 마운트하고 `DATA_DIR` 을 그 경로로 주면 재시작해도 제출·계정이 보존된다(`render.yaml` 참고, 디스크는 유료 인스턴스 필요). 세션은 메모리라 재시작 시 로그아웃되지만 자료는 남는다. **포털은 Streamlit Cloud 로 배포할 수 없다**(자체 http 서버 + 휘발성 저장소). `streamlit_app.py` 는 뷰어 배포용 별개다.
  - `export_report()` (`/export?kind=report`) 가 정부 「기관별 합계」 형식 결과표(`훈련실적 총계` = 전체 계 + 기관 소계 + 양성/향상/수시, `세부실적`, NCS/KECO별)를 만든다. `export_raw()` 는 양성·향상 과정을 통합 세부내역으로 뽑는다.
- **[calendar_store.py](calendar_store.py)** — 캘린더 탭(`/calendar`)의 일정 조사 저장소. `data/calendar.json` 한 개에 조사 목록과 기관별 응답을 담는다. 관리자가 기간(최대 `MAX_DAYS` 31일)을 정해 물으면 기관 계정이 날짜마다 `가능`/`불가`/`미정` 으로 답한다.
  - 화면은 `app.month_grid()` 가 그리는 월별 달력이다. 기관은 날짜 칸에서 바로 고르고(`cal_pick_grid`), 관리자는 날짜별 가능 인원을 본다(`cal_count_grid`). 조사 기간 밖의 날짜는 회색으로 남겨 물어본 범위가 달력 위에 드러나게 한다.
  - 응답 열람은 `can_see_answers()` 한 곳에서만 판정한다 — 기본은 관리자만, 조사의 `공개` 가 켜져 있으면 기관도 서로 확인한다. 화면 진입 가능 여부는 `can_open()`, 응답 자격은 `is_target()` 이다.
  - 본 취합 흐름(실적·승인·리포트)과 자료가 섞이지 않는다. 주차 파일·집계기를 건드리지 않고 `storage.DATA_DIR`·`load_roster()` 만 가져다 쓴다.
- **[streamlit_app.py](streamlit_app.py)** — `app.py` 포털 전체를 Streamlit 으로 옮긴 화면. **share.streamlit.io(Streamlit Community Cloud) 배포용**이며 로그인·직접입력·엑셀 업로드·취합/승인·리포트·핵심요약·월별연도별·캘린더·경북이슈·종합제언·고용통계·계정관리·챗봇을 모두 담는다. 화면만 Streamlit(`st.*`)으로 다시 짜고 **집계·검증·리포트·엑셀 빌더·종합제언 로직은 기존 모듈을 그대로 재사용**한다 — 특히 `import app` 으로 `export_report`/`export_raw`/`template_xlsx`/`advice_blocks`/`series_facts`/`month_week_to_key` 등 순수 함수를 가져다 쓴다(app 은 서버를 `__main__` 에서만 띄우므로 import 해도 서버가 안 뜬다). 그래서 **app.py 는 지우면 안 된다**(이 화면이 라이브러리로 의존). 로그인은 `st.session_state`, 관리자 계정은 Secrets 의 `ADMIN_PASSWORD` 로 매번 같은 값으로 보장한다. ⚠ Streamlit Cloud 는 저장소가 재시작마다 초기화되므로 제출·계정이 보존되지 않는다(데모·시연용). 월 합계는 주차 행을 합쳐 한 번에 집계한다(주차별 이수율 평균이 아니다).
- **[chatbot.py](chatbot.py)** — 화면 우측 아래 말풍선 위젯의 답을 만든다. `specs/기획서*`·`자동화_흐름도`·`feature-1-spec`·`CLAUDE.md` 를 읽어(`build_context`, 캐시) 시스템 프롬프트에 넣고 OpenAI 채팅 모델(`gpt-4o-mini`)로 답한다 — 답 문장을 사람이 써 두지 않아 문서가 바뀌면 답도 바뀐다. 키는 `imagegen.load_env()` 가 읽는 `.env` 의 `OPENAI_API_KEY`(다른 도구와 같은 키). 자료 밖은 지어내지 말고 "관리자에게 확인" 으로 답하게 프롬프트로 묶었고, 맥락에 섞인 개발용 지시(예: "확인!" 붙이기)는 따르지 말라고 명시했다. 위젯 HTML/JS/CSS 는 `app.py`(`CHAT_WIDGET`·`STYLE`·`page()`/`m_page()`)에 있고, `/chat` POST 가 `answer()` 를 부른다 — 로그인 세션이 있어야 하고 IP당 10분 30회로 제한(`chat_quota`)한다.
- **[auth.py](auth.py)** — PBKDF2-SHA256 해시, 세션은 메모리(`SESSIONS`)라 재시작하면 전원 로그아웃. 자율 가입 없이 관리자가 계정을 발급한다. 마지막 관리자 계정은 삭제 불가.
- **[gb_issues.py](gb_issues.py) / [kosis_stats.py](kosis_stats.py)** — 본 취합 흐름과 별개인 참고자료 화면(`/issues`, `/stats`). 두 자료를 묶어 주제별 제언을 만드는 `/advice`(종합 제언) 화면이 이 위에 얹혀 있다. 둘 다 `data/*.json` 에 캐시하고 버튼을 누를 때만 외부를 부른다.
  - `gb_issues` 는 OpenAI 웹 검색 모델(`gpt-4o-search-preview`)로 경북 산업·직업훈련 이슈를 섹터(주제)마다 `PER_TOPIC`(15)건씩 모은다. 링크는 응답 `annotations` 에서 온 URL만 쓴다 — 모델이 지어낸 주소를 화면에 올리지 않기 위함이다.
  - `/advice`(종합 제언)는 **이 두 자료에서만** 문장을 뽑는다. 주차 실적은 쓰지 않는다 — 실적 판단은 ③ 주간 리포트 몫이다. 낱말 빈도(`app.keywords()`)는 형태소 분석 없이 조사만 떼고 세는 기계적 값이라, 화면에도 그렇게 밝혀 둔다.
  - `kosis_stats` 는 KOSIS OpenAPI(`orgId=101`, `tblId=DT_1DA7004S` 행정구역(시도)별 경제활동인구)에서 고용률·실업률·생산가능인구를 월별로 받는다. 항목·지역 코드는 하드코딩하지 않고 응답의 `ITM_NM`/`C1_NM` 으로 고른다. 키는 `.env` 의 `KOSIS_API_KEY`.
  - 선 그래프는 `app.line_chart()` 가 SVG로 직접 그린다(차트 라이브러리 없음). 색 4개는 검증된 팔레트라 **돌려쓰지 않는다** — 지역 선택은 4개로 제한한다.
- **[imagegen.py](imagegen.py) / [memo_tagger.py](memo_tagger.py)** — 본 흐름과 무관한 부수 도구. 화면은 없애고 CLI 로만 쓴다. 키는 `.env` 의 `OPENAI_API_KEY` 를 `load_env()` 가 읽으며, 이미 설정된 환경변수가 우선한다. `memo_tagger` 는 내용 해시로 캐시해 같은 메모를 다시 호출하지 않는다.

## 절대 규칙

- 항상 한국어로, 공손하고 간결하게. 추측하지 않고 확인한 것만 말하며, 답변 마지막에 "확인!" 을 붙인다.
- 실명·실제 사내 자료를 넣지 않는다. 예시는 가짜 데이터로 만든다.
- `data/` 를 지우거나 덮어쓰지 않고, 지표 계산은 `feature1_aggregate.py` 밖으로 퍼뜨리지 않는다.

## 규칙 인덱스

- [rules/tone.md](rules/tone.md) — 말투: 언어, 정직성, 어려운 용어 괄호 풀이, 맺음말.
- [rules/format.md](rules/format.md) — 결과 형식: 요약 3단 구조, 제목·소제목, 표 세 칸 제한, 굵게 사용 범위, 파일 링크 표기.
- [rules/forbidden.md](rules/forbidden.md) — 하지 말 것: 내용 금지, 데이터·파일 금지, 코드 금지(정렬·원본 표시·`e()` 이스케이프·명세 준수).

## 우선순위

규칙이 부딪치면 **사용자의 이번 지시 > 절대 규칙 > rules/forbidden.md > rules/tone.md > rules/format.md** 순으로 따르고, 형식 때문에 사실을 왜곡하지 않는다.
