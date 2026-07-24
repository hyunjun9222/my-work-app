"""주간 훈련기관 교육실적 취합 — Streamlit 포털

app.py(자체 http 서버 포털)의 기능을 Streamlit 한 화면으로 옮긴 것.
share.streamlit.io(Streamlit Community Cloud)에서 그대로 돈다.

    streamlit run streamlit_app.py

집계·검증·리포트·엑셀 빌더·종합제언 로직은 기존 모듈을 그대로 재사용한다
(storage·feature1_aggregate·report_engine·auth·calendar_store·chatbot·app).
app 모듈은 서버를 __main__ 에서만 띄우므로 import 해도 서버가 뜨지 않는다.

⚠ Streamlit Community Cloud 는 저장소(data/)가 재시작·재배포마다 초기화된다.
   → 계정·제출·승인이 사라진다. 관리자 계정은 Secrets 의 ADMIN_PASSWORD 로 매번
   같은 값으로 다시 만들어, 재시작 뒤에도 같은 비밀번호로 로그인되게 한다.
"""

import io
from datetime import date

import pandas as pd
import streamlit as st

import app  # 순수 로직·엑셀 빌더 재사용 (import 안전: 서버는 __main__ 에서만)
import auth
import calendar_store
import chatbot
import feature1_aggregate as f1
import gb_issues
import kosis_stats
import report_engine as rpt
import storage

st.set_page_config(page_title="주간 훈련기관 교육실적 취합", page_icon="🏫", layout="wide")


# ── 지표 표기 (float | None | 문자열 상수) ───────────────────────


def pct(v):
    if isinstance(v, str):
        return v
    return "—" if v is None else f"{v * 100:.1f}%"


def num0(v):
    return 0 if v is None else v


# ── 관리자 계정 보장 (Cloud 는 재시작마다 초기화) ────────────────


def ensure_admin():
    """관리자 계정이 없으면 만든다. Secrets 의 ADMIN_PASSWORD 를 쓰고, 없으면 랜덤(로그에만)."""
    if any(u.get("권한") == "admin" for u in auth.load_users().values()):
        return None
    try:
        pw = str(st.secrets["ADMIN_PASSWORD"])
    except Exception:
        pw = None
    if pw:
        auth.create_user("admin", pw, "admin")
        return None
    return auth.ensure_admin()  # 랜덤 — Streamlit 앱 로그(Manage app → Logs)에 한 번 찍힘


_first_pw = ensure_admin()


# ── 로그인 ───────────────────────────────────────────────────────


def login_gate():
    """로그인 안 됐으면 로그인 화면을 보이고 True 를 돌려준다(=여기서 멈춤)."""
    if st.session_state.get("sess"):
        return False
    st.title("🏫 주간 훈련기관 교육실적 취합")
    st.caption("관리자가 발급한 아이디로 접속하세요. 자율 가입은 없습니다.")
    if _first_pw:
        st.warning(
            "관리자 계정을 새로 만들었습니다. Secrets 에 `ADMIN_PASSWORD` 를 설정한 뒤 재시작하면 "
            "그 비밀번호로 로그인됩니다. (지금 임시 비밀번호는 앱 로그에서 확인)"
        )
    uid = st.text_input("아이디", key="login_id")
    pw = st.text_input("비밀번호", type="password", key="login_pw")
    if st.button("로그인", type="primary"):
        u = auth.verify(uid, pw)
        if u:
            st.session_state["sess"] = {"id": u["아이디"], "role": u["권한"], "org": u.get("기관명", "")}
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 맞지 않습니다.")
    st.info("🧪 로그인 후 관리자 화면의 '예시 데이터 채우기' 로 데모를 빠르게 채울 수 있습니다.")
    return True


# ── 주차 선택 ────────────────────────────────────────────────────


def week_picker(label="주차", key="wk"):
    """저장된 주차 중 하나를 고른다. 없으면 None."""
    weeks = storage.list_weeks()
    if not weeks:
        st.info("저장된 주차가 없습니다. 먼저 ① 입력/업로드로 실적을 받아 주세요.")
        return None
    return st.selectbox(label, weeks, key=key)


def new_week_inputs(prefix=""):
    """연·월·(그 달의) 주차 → 주차 키. app.month_week_to_key 재사용."""
    c1, c2, c3 = st.columns(3)
    y = c1.number_input("연", 2000, 2100, date.today().year, key=prefix + "y")
    m = c2.number_input("월", 1, 12, date.today().month, key=prefix + "m")
    n = c3.number_input("그 달의 몇 주차", 1, 6, 1, key=prefix + "n")
    key, 기준일 = app.month_week_to_key(int(y), int(m), int(n))
    st.caption(f"연중 주차 **{key}** · 기준일(일요일) {기준일 or '-'}")
    return key, int(m)


# ── 화면: 홈 ─────────────────────────────────────────────────────


def page_home(sess):
    st.header("대시보드")
    st.caption("기관이 직접 입력하거나 엑셀로 올리면 같은 저장소에 쌓이고, 관리자가 승인하면 주간 리포트가 만들어집니다.")
    weeks = storage.list_weeks()
    roster = storage.load_roster()
    if not weeks:
        st.info("아직 저장된 제출이 없습니다. 왼쪽 메뉴에서 ① 직접 입력 또는 ① 엑셀 업로드로 시작하세요.")
        return
    rows = []
    for k in weeks[:12]:
        d = storage.load_week(k)
        subs = d["제출"]
        승인 = sum(1 for s in subs.values() if s["상태"] == "승인")
        과정 = sum(len(s["실적"]) for s in subs.values())
        rows.append({"주차": k, "제출 기관": len(subs), "과정": 과정,
                     "승인": f'{승인}/{len(roster) or len(subs)}'})
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")


# ── 화면: 직접 입력 ──────────────────────────────────────────────

PERF_UI = ["구분", "정기수시", "과정구분", "NCS대분류명", "KECO세분류명", "과정명",
           "훈련목표인원", "훈련실시인원", "중도탈락자", "훈련중", "훈련수료인원", "취업인원"]


def page_submit(sess):
    st.header("① 실적 직접 입력")
    st.caption("한 줄 = 한 과정. 수료율=수료÷실시, 탈락률=중도탈락÷실시, 취업은 양성 과정만.")
    관리자 = sess["role"] == "admin"
    roster = storage.load_roster()
    if 관리자:
        org = st.selectbox("훈련기관", roster or ["(명단 없음 — 계정 관리에서 등록)"])
    else:
        org = sess["org"]
        st.text_input("훈련기관", org, disabled=True)

    key, month = new_week_inputs("s_")
    기존 = (storage.load_week(key) or {}).get("제출", {}).get(org) if key else None
    초기 = pd.DataFrame(기존["실적"] if 기존 else [], columns=PERF_UI)
    if 초기.empty:
        초기 = pd.DataFrame([{c: ("양성" if c == "구분" else ("정기" if c == "정기수시" else None)) for c in PERF_UI}])

    ed = st.data_editor(
        초기, num_rows="dynamic", width="stretch", key="perf_editor",
        column_config={
            "구분": st.column_config.SelectboxColumn(options=["양성", "향상"], required=True),
            "정기수시": st.column_config.SelectboxColumn(options=["정기", "수시"]),
            "훈련목표인원": st.column_config.NumberColumn(min_value=0, format="%d"),
            "훈련실시인원": st.column_config.NumberColumn(min_value=0, format="%d"),
            "중도탈락자": st.column_config.NumberColumn(min_value=0, format="%d"),
            "훈련중": st.column_config.NumberColumn(min_value=0, format="%d"),
            "훈련수료인원": st.column_config.NumberColumn(min_value=0, format="%d"),
            "취업인원": st.column_config.NumberColumn(min_value=0, format="%d"),
        },
    )
    if st.button("제출하고 잠정 계산 보기", type="primary"):
        if not org or "명단 없음" in str(org):
            st.error("훈련기관을 먼저 정해 주세요(계정 관리에서 명단 등록).")
            return
        perf = []
        for _, r in ed.iterrows():
            if not str(r.get("과정명") or "").strip():
                continue
            perf.append({c: (int(r[c]) if c in f1.COUNT_COLS and pd.notna(r[c]) else
                             (str(r[c]).strip() if pd.notna(r[c]) else "")) for c in PERF_UI})
        if not perf:
            st.error("실적이 한 건도 입력되지 않았습니다(과정명 필수).")
            return
        storage.save_submission(key, org, month, perf, [], [], "직접 입력", targets=None)
        st.success(f"{org} · {key} 제출 저장 완료. 아래 잠정 계산을 확인하세요.")
        _preview(storage.load_week(key), [org])


def _preview(data, orgs):
    only = {"제출": {o: data["제출"][o] for o in orgs if o in data["제출"]}}
    res = f1.process_rows(storage.to_rows(only))
    t = res["합계_전체"][0]
    c1, c2, c3 = st.columns(3)
    c1.metric("실시율(목표 대비)", pct(t["실시율"]))
    c2.metric("수료율(실시 대비)", pct(t["수료율"]))
    c3.metric("집계 과정", f'{t["과정수"]}개' + (f' · 제외 {t["제외"]}' if t["제외"] else ""))
    df = pd.DataFrame([{
        "구분": r["구분"], "기관": r["기관명"], "과정": r["과정명"],
        "목표": r["훈련목표인원"], "실시": r["훈련실시인원"], "수료": r["훈련수료인원"],
        "탈락": r["중도탈락자"], "수료율": pct(r["수료율"]), "오류": "⚠" if r["_오류"] else "",
    } for r in res["표"]])
    st.dataframe(df, hide_index=True, width="stretch")
    if res["오류"]:
        st.warning(f"입력 오류 {len(res['오류'])}건 — 오류 행은 지표 계산·합계에서 제외됩니다.")


# ── 화면: 엑셀 업로드 ────────────────────────────────────────────


def page_upload(sess):
    st.header("① 엑셀 업로드")
    st.caption("정부 「지산맞 훈련실적」 양식(양성훈련 현황·향상훈련 현황 시트)을 그대로 올립니다.")
    st.download_button("간소 빈 양식 내려받기 (.xlsx)", app.template_xlsx(),
                       file_name=app.TEMPLATE_NAME,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    up = st.file_uploader("실적 엑셀 (.xlsx)", type=["xlsx", "xlsm"])
    key, month = new_week_inputs("u_")
    if up and st.button("업로드하고 잠정 계산 보기", type="primary"):
        import tempfile
        from pathlib import Path
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "up.xlsx"
            p.write_bytes(up.getvalue())
            try:
                rows = f1.read_rows(p)
                targets = f1.read_targets(app.load_workbook(p, data_only=True))
            except SystemExit as ex:
                st.error(f"엑셀을 읽지 못했습니다 — {ex}")
                return
            except Exception as ex:
                st.error(f"파일을 읽지 못했습니다 — {type(ex).__name__}: {ex}")
                return
        by_org = {}
        for r in rows:
            o = (r.get("기관명") or "(기관명 미기재)").strip() or "(기관명 미기재)"
            by_org.setdefault(o, []).append({c: r.get(c) for c in storage.PERF_COLS})
        skipped = []
        if sess["role"] != "admin":
            mine = sess["org"]
            skipped = [o for o in by_org if o != mine]
            by_org = {o: v for o, v in by_org.items() if o == mine}
            if not by_org:
                st.error(f"파일에 '{mine}' 기관(훈련센터명)의 과정이 없습니다.")
                return
        for o, perf in by_org.items():
            storage.save_submission(key, o, month, perf, [], [], "엑셀 업로드",
                                    targets=targets if len(by_org) == 1 else None)
        st.success(f"저장 완료: {', '.join(by_org)}" + (f" · 다른 기관 {len(skipped)}곳 제외" if skipped else ""))
        _preview(storage.load_week(key), list(by_org))


# ── 화면: 취합·승인 (관리자) ─────────────────────────────────────


def page_admin(sess):
    st.header("② 취합·승인")
    with st.expander("대상 훈련기관 명단"):
        cur = "\n".join(storage.load_roster())
        new = st.text_area("한 줄에 한 기관", cur, height=140)
        if st.button("명단 저장"):
            storage.save_roster(new.splitlines())
            st.success("명단을 저장했습니다."); st.rerun()
    key = week_picker("주차", "admin_wk")
    if not key:
        return
    data = storage.load_week(key)
    roster = storage.load_roster()
    subs = data["제출"]
    승인 = sum(1 for s in subs.values() if s["상태"] == "승인")
    base = len(roster) or len(subs)
    st.metric("승인 완료율", f"{승인}/{base}")
    미제출 = [r for r in roster if r not in subs]
    if 미제출:
        st.warning("미제출: " + ", ".join(미제출))
    for org, s in subs.items():
        badge = {"승인": "🟢", "제출": "🟡", "반려": "🔴"}.get(s["상태"], "⚪")
        with st.container(border=True):
            c1, c2 = st.columns([3, 2])
            c1.markdown(f"**{org}** {badge} {s['상태']} · 과정 {len(s['실적'])}개 · {s['제출시각']} · {s['출처']}")
            b1, b2, b3 = c2.columns(3)
            if b1.button("승인", key=f"ap_{org}"):
                storage.set_status(key, org, "승인"); st.rerun()
            if b2.button("반려", key=f"rj_{org}"):
                storage.set_status(key, org, "반려"); st.rerun()
            if b3.button("삭제", key=f"dl_{org}"):
                storage.delete_submission(key, org); st.rerun()


# ── 화면: 주간 리포트 ────────────────────────────────────────────


def _build(week):
    prev = storage.prev_week_key(week)
    ly = storage.last_year_key(week)
    return rpt.build_report(storage.load_week(week), storage.load_week(prev) if prev else None,
                            storage.load_week(ly), storage.load_roster(), week)


def page_report(sess):
    st.header("③ 주간 리포트")
    key = week_picker("주차", "rep_wk")
    if not key:
        return
    rep = _build(key)
    res = rep["실적"]
    t = res["합계_전체"][0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("실시율(목표 대비)", pct(t["실시율"]))
    c2.metric("수료율(실시 대비)", pct(t["수료율"]))
    c3.metric("전체 탈락률", pct(t["탈락률"]))
    c4.metric("집계 과정", f'{t["과정수"]}개')

    st.subheader("제출·승인 현황판")
    st.dataframe(pd.DataFrame([{"기관": b["기관명"], "상태": b["상태"], "오류행": b["오류행"], "비고": b["비고"]}
                              for b in rep["현황판"]], ), hide_index=True, width="stretch")

    st.subheader("기관별 수료율(실시 대비)")
    g = pd.DataFrame([{"기관": x["구분"], "수료율(%)": (x["수료율"] * 100 if isinstance(x["수료율"], float) else None)}
                      for x in res["합계_기관별"]]).set_index("기관")
    st.bar_chart(g, height=320)

    st.subheader("통합 교육실적표")
    st.dataframe(pd.DataFrame([{
        "구분": r["구분"], "기관": r["기관명"], "과정": r["과정명"], "목표": r["훈련목표인원"],
        "실시": r["훈련실시인원"], "수료": r["훈련수료인원"], "탈락": r["중도탈락자"],
        "수료율": pct(r["수료율"]), "탈락률": pct(r["탈락률"]), "오류": "⚠" if r["_오류"] else "",
    } for r in res["표"]]), hide_index=True, width="stretch")

    if rep["이상치"] and rep["이상치"]["플래그"]:
        st.subheader("작년 동기 대비 이상치")
        for fl in rep["이상치"]["플래그"]:
            st.write(f"{'🔻' if fl['방향']=='급감' else '🔺'} **{fl['기관명']} · {fl['과정명']}** — "
                     f"작년 {fl['작년']*100:.1f}% → 이번 {fl['이번']*100:.1f}% ({fl['차이']:+.1f}%p)")

    st.subheader("요약 초안")
    st.text_area("문구를 고쳐 확정하세요(화면 편집만)", rep["요약초안"], height=170)

    d1, d2 = st.columns(2)
    d1.download_button("리포트 엑셀 받기", app.export_report(key), file_name=f"{key}-report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("입력값 엑셀 받기", app.export_raw(key), file_name=f"{key}-input.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_brief(sess):
    st.header("③-1 핵심 요약")
    key = week_picker("주차", "brief_wk")
    if not key:
        return
    rep = _build(key)
    t = rep["실적"]["합계_전체"][0]
    board = rep["현황판"]
    승인 = sum(1 for b in board if b["상태"] == "승인")
    막힘 = [b["기관명"] for b in board if b["상태"] in ("미제출", "반려 대상")]
    todos = rep["오늘할일"]
    c1, c2, c3 = st.columns(3)
    c1.metric("실시율(목표 대비)", pct(t["실시율"]))
    c2.metric("수료율(실시 대비)", pct(t["수료율"]))
    c3.metric("제출·승인", f"{승인}/{len(board)}")
    st.write(f"**리포트 발행**: {'보류 — ' + ', '.join(막힘) if 막힘 else '가능(미제출·반려 없음)'}")
    st.subheader(f"오늘 할 일 {len(todos)}건")
    for x in todos[:12]:
        st.write(f"- [{x['우선순위']}] **{x['대상']}** · {x['할일']}")


# ── 화면: 월별·연도별 ────────────────────────────────────────────


def _ym(key):
    p = storage.parse_week(key)
    if not p:
        return None
    y, w = p
    try:
        d = date.fromisocalendar(y, w, 7)
        return d.year, d.month
    except ValueError:
        return None


def _totals(keys, approved, org):
    찾 = (org or "").strip().lower()
    rows = []
    for k in keys:
        d = storage.load_week(k)
        if not d:
            continue
        for r in storage.to_rows(d, only_approved=approved):
            if 찾 and 찾 not in str(r.get("기관명", "")).lower():
                continue
            rows.append({**r, "_행": len(rows) + 2})
    return f1.process_rows(rows)["합계_전체"][0] if rows else None


def page_trend(sess):
    st.header("③-2 월별·연도별 보기")
    idx = {}
    for k in storage.list_weeks():
        ym = _ym(k)
        if ym:
            idx.setdefault(ym, []).append(k)
    if not idx:
        st.info("저장된 주차 실적이 없습니다.")
        return
    years = sorted({y for y, _ in idx}, reverse=True)
    c1, c2, c3, c4 = st.columns(4)
    단위 = c1.radio("보기 단위", ["월별", "연도별"], horizontal=True)
    연도 = c2.selectbox("연도", years)
    승인만 = c3.checkbox("승인분만")
    비교 = c4.checkbox("작년 비교", value=True)
    org = st.text_input("훈련기관(일부, 비우면 전체)", "")

    if 단위 == "월별":
        구간 = [(f"{연도}-{m:02d}", idx.get((연도, m), []), idx.get((연도 - 1, m), [])) for m in range(1, 13)]
        구간 = [x for x in 구간 if x[1] or x[2]]
        비교이름 = f"{연도 - 1}년 같은 달"
    else:
        구간 = [(f"{y}년",
                 [k for (yy, _), ks in idx.items() if yy == y for k in ks],
                 [k for (yy, _), ks in idx.items() if yy == y - 1 for k in ks]) for y in sorted(years)]
        비교이름 = "직전 연도"

    표 = []
    for 이름, 이번, 작년 in 구간:
        t = _totals(이번, 승인만, org)
        p = _totals(작년, 승인만, org) if 비교 else None
        수 = t["수료율_목표"] if t else None
        작 = p["수료율_목표"] if p else None
        표.append({"구간": 이름, "과정": t["과정수"] if t else 0,
                   "목표": (t.get("목표훈련인원") or 0) if t else 0,
                   "수료": t["훈련수료인원"] if t else 0,
                   "수료율(목표대비)": (수 * 100 if isinstance(수, float) else None),
                   "작년 수료율": (작 * 100 if isinstance(작, float) else None)})
    df = pd.DataFrame(표)
    있는 = df[df["과정"] > 0]
    if 있는.empty:
        st.warning("고른 범위에 집계된 과정이 없습니다.")
        return
    총목표, 총수료 = int(있는["목표"].sum()), int(있는["수료"].sum())
    st.metric("전체 수료율(목표 대비)", pct(총수료 / 총목표 if 총목표 else None))
    칼럼 = ["수료율(목표대비)"] + (["작년 수료율"] if 비교 and df["작년 수료율"].notna().any() else [])
    st.bar_chart(df.set_index("구간")[칼럼], height=340)
    st.dataframe(df, hide_index=True, width="stretch")


# ── 화면: 캘린더(일정 조사) ──────────────────────────────────────


def page_calendar(sess):
    st.header("캘린더 — 일정 조사")
    roster = storage.load_roster()
    관리자 = sess["role"] == "admin"
    if 관리자:
        with st.expander("새 일정 조사 만들기"):
            제목 = st.text_input("제목", key="cal_title")
            c1, c2, c3 = st.columns(3)
            시작 = c1.date_input("시작일", date.today(), key="cal_s")
            종료 = c2.date_input("종료일", date.today(), key="cal_e")
            공개 = c3.checkbox("전체 공개", key="cal_pub")
            대상 = st.multiselect("대상 기관(비우면 전체)", roster, key="cal_tg")
            if st.button("조사 만들기"):
                ev, err = calendar_store.create(제목, 시작.isoformat(), 종료.isoformat(),
                                                대상=대상, 공개=공개, 작성자=sess["id"])
                if err:
                    st.error(err)
                else:
                    st.success("만들었습니다."); st.rerun()
    events = calendar_store.list_events()
    events = [e for e in events if calendar_store.can_open(e, sess["role"], sess["org"], roster)]
    if not events:
        st.info("등록된 일정 조사가 없습니다.")
        return
    ev = st.selectbox("조사 선택", events, format_func=lambda e: f'{e["제목"]} ({e["시작일"]}~{e["종료일"]})')
    days = calendar_store.days_of(ev)
    if sess["role"] != "admin" and calendar_store.is_target(ev, sess["org"], roster) and not calendar_store.is_closed(ev):
        st.subheader("내 응답")
        mine = (ev.get("응답", {}).get(sess["org"]) or {}).get("날짜별", {})
        골라 = {}
        for d in days:
            골라[d] = st.radio(calendar_store.day_label(d), calendar_store.ANSWERS,
                               index=calendar_store.ANSWERS.index(mine.get(d, "미정")),
                               horizontal=True, key=f"cal_{ev['번호']}_{d}")
        if st.button("응답 제출"):
            calendar_store.answer(ev["번호"], sess["org"], 골라, 응답자=sess["id"])
            st.success("응답을 저장했습니다."); st.rerun()
    if calendar_store.can_see_answers(ev, sess["role"], sess["org"], roster):
        st.subheader("응답 현황 (날짜별 '가능' 인원)")
        표 = calendar_store.tally(ev)
        st.dataframe(pd.DataFrame([{"날짜": calendar_store.day_label(d),
                                    "가능": 표[d]["가능"], "불가": 표[d]["불가"]} for d in days]),
                     hide_index=True, width="stretch")
        best = calendar_store.best_days(ev)
        if best:
            st.success("가장 많이 가능한 날 — " + ", ".join(calendar_store.day_label(d) for d in best))
    if 관리자 and st.button("이 조사 삭제"):
        calendar_store.delete(ev["번호"]); st.rerun()


# ── 화면: 경북 이슈 ──────────────────────────────────────────────


def page_issues(sess):
    st.header("경북 산업·직업훈련 이슈")
    cache = gb_issues.load_cache()
    st.caption(f"섹터 {len(gb_issues.TOPICS)}개 · 최종 수집 {gb_issues.last_updated(cache) or '없음'}")
    if sess["role"] == "admin" and st.button("웹에서 새로 수집 (API 요금 발생)"):
        try:
            with st.spinner("수집 중…"):
                cache = gb_issues.collect()
            st.success("수집 완료."); st.rerun()
        except Exception as ex:
            st.error(f"수집 실패 — {ex}")
    for block in gb_issues.ordered(cache):
        출처 = sum(1 for it in block["이슈"] if it["출처"])
        st.subheader(f'{block["주제"]}  ({len(block["이슈"])}건 · 출처 {출처})')
        if block["이슈"] and not 출처:
            st.warning("이 섹터는 출처가 없습니다 — 지어냈을 수 있으니 다시 수집하세요.")
        for it in block["이슈"]:
            links = " · ".join(f'[{s["제목"][:30]}]({s["url"]})' for s in it["출처"])
            st.markdown(f"- {it['내용']}" + (f"  \n  {links}" if links else ""))


# ── 화면: 종합 제언 ──────────────────────────────────────────────


def page_advice(sess):
    st.header("종합 제언")
    st.caption("「경북 이슈」와 「고용 통계」에서만 뽑아 산업·고용·직업훈련 관점으로 정리합니다.")
    facts = app.series_facts(kosis_stats.load_cache())
    issues = gb_issues.ordered(gb_issues.load_cache())
    blocks = app.advice_blocks(facts, issues)
    for idx, (name, desc) in enumerate(app.ADVICE_TOPICS, 1):
        b = blocks[name]
        st.subheader(f"{idx}. {name}")
        st.caption(desc)
        st.markdown("**근거**\n" + "\n".join(f"- {x}" for x in b["근거"]))
        if b["중점"]:
            st.markdown("**주요 중점사항**\n" + "\n".join(f"- {x}" for x in b["중점"]))
        if b["전략"]:
            st.markdown("**향후 전략**\n" + "\n".join(f"- {x}" for x in b["전략"]))


# ── 화면: 고용 통계 ──────────────────────────────────────────────


def page_stats(sess):
    st.header("고용 통계 (KOSIS)")
    data = kosis_stats.load_cache()
    st.caption(f"최종 수집 {data.get('수집시각', '없음')}")
    if sess["role"] == "admin" and st.button("KOSIS 에서 새로 불러오기"):
        try:
            with st.spinner("불러오는 중…"):
                data = kosis_stats.collect()
            st.success("수집 완료."); st.rerun()
        except Exception as ex:
            st.error(f"수집 실패 — {ex}")
    시점 = [kosis_stats.fmt_period(p) for p in data.get("시점", [])]
    for 이름, block in (data.get("지표") or {}).items():
        st.subheader(f'{이름} ({block.get("단위","")})')
        df = pd.DataFrame({지역: vs for 지역, vs in block.get("지역", {}).items()}, index=시점)
        st.line_chart(df, height=300)
    if not data.get("지표"):
        st.info("아직 받아온 통계가 없습니다.")


# ── 화면: 계정 관리 (관리자) ─────────────────────────────────────


def page_users(sess):
    st.header("계정 관리")
    users = auth.load_users()
    st.dataframe(pd.DataFrame([{"아이디": uid, "권한": u["권한"], "소속": u.get("기관명", ""),
                                "발급": u.get("발급시각", "")} for uid, u in users.items()]),
                 hide_index=True, width="stretch")
    with st.expander("새 계정 발급"):
        c1, c2, c3, c4 = st.columns(4)
        uid = c1.text_input("아이디", key="nu_id")
        pw = c2.text_input("초기 비밀번호", key="nu_pw")
        role = c3.selectbox("권한", ["org", "admin"], key="nu_role")
        org = c4.selectbox("소속 기관", [""] + storage.load_roster(), key="nu_org")
        if st.button("발급"):
            if auth.create_user(uid, pw, role, org):
                st.success(f"{uid} 발급 완료."); st.rerun()
            else:
                st.error("이미 있는 아이디이거나 값이 비었습니다.")
    with st.expander("데모용 예시 데이터 채우기"):
        st.caption("개인정보 없는 가짜 실적을 저장소에 채웁니다(빈 화면 방지용).")
        if st.button("예시 데이터 채우기"):
            n = _load_demo()
            st.success(f"예시 {n}개 주차를 채웠습니다."); st.rerun()


def _load_demo():
    """demo_data/ 의 예시 주차·명단을 현재 저장소로 복사한다."""
    from pathlib import Path
    import json
    src = Path(__file__).parent / "demo_data"
    if not src.exists():
        return 0
    storage.DATA_DIR.mkdir(exist_ok=True)
    n = 0
    for f in src.glob("*.json"):
        (storage.DATA_DIR / f.name).write_text(f.read_text(encoding="utf-8"), encoding="utf-8")
        if f.stem not in ("roster",):
            n += 1
    return n


# ── 화면: 챗봇 ───────────────────────────────────────────────────


def page_chat(sess):
    st.header("도움 챗봇")
    st.caption("이 서비스에 대해 물어보세요. 자료 밖은 지어내지 않습니다.")
    hist = st.session_state.setdefault("chat", [])
    for m in hist:
        st.chat_message(m["role"]).write(m["content"])
    if q := st.chat_input("예: 실적은 어떻게 제출하나요?"):
        hist.append({"role": "user", "content": q})
        st.chat_message("user").write(q)
        try:
            with st.spinner("답변 작성 중…"):
                a = chatbot.answer(hist)
        except chatbot.ChatError as ex:
            a = str(ex)
        except Exception as ex:
            a = f"답변 중 문제가 생겼습니다 — {type(ex).__name__}"
        hist.append({"role": "assistant", "content": a})
        st.chat_message("assistant").write(a)


# ── 라우팅 ───────────────────────────────────────────────────────

PAGES_ADMIN = {
    "홈": page_home, "① 직접 입력": page_submit, "① 엑셀 업로드": page_upload,
    "② 취합·승인": page_admin, "③ 주간 리포트": page_report, "③-1 핵심 요약": page_brief,
    "③-2 월별·연도별": page_trend, "캘린더": page_calendar, "경북 이슈": page_issues,
    "종합 제언": page_advice, "고용 통계": page_stats, "계정 관리": page_users, "도움 챗봇": page_chat,
}
PAGES_ORG = {
    "홈": page_home, "① 직접 입력": page_submit, "① 엑셀 업로드": page_upload,
    "캘린더": page_calendar, "경북 이슈": page_issues, "종합 제언": page_advice,
    "고용 통계": page_stats, "도움 챗봇": page_chat,
}


def main():
    st.session_state.setdefault("sess", None)
    if login_gate():
        return
    sess = st.session_state["sess"]
    pages = PAGES_ADMIN if sess["role"] == "admin" else PAGES_ORG
    st.sidebar.markdown("### 🏫 교육실적 취합")
    역할 = "관리자" if sess["role"] == "admin" else (sess["org"] or "기관")
    st.sidebar.caption(f"{sess['id']} · {역할}")
    choice = st.sidebar.radio("메뉴", list(pages), label_visibility="collapsed")
    if st.sidebar.button("로그아웃"):
        st.session_state["sess"] = None
        st.rerun()
    pages[choice](sess)


main()
