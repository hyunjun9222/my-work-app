"""훈련실적 월별·연도별 보기 — Streamlit 화면

터미널 없이 실적을 훑어보려고 기존 앱 위에 얹은 별도 화면이다.
본 흐름(app.py 의 제출·승인·리포트)은 건드리지 않고, 저장소(data/*.json)를
읽기만 한다. 저장·수정 기능은 일부러 넣지 않았다.

    streamlit run streamlit_app.py      # http://localhost:8501

지표(수료율·탈락률)는 여기서 다시 계산하지 않는다. 주차 파일을 행 목록으로
펼쳐 feature1_aggregate.process_rows 에 넣고, 그 결과의 합계만 가져다 쓴다 —
화면이 달라도 숫자는 app.py·CLI 와 같아야 하기 때문이다.
"""

from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import storage
from feature1_aggregate import UNCOMPUTABLE, UNVERIFIED, process_rows

MONTHS = list(range(1, 13))

# 배포(클라우드)에는 data/ 가 없다(gitignore). 그럴 때는 함께 올린 예시 데이터로 화면을 채운다.
# 로컬에서 실제 제출이 쌓여 있으면(data/ 가 비지 않으면) 그대로 실제 자료를 쓴다.
DEMO_DIR = Path(__file__).parent / "demo_data"
DEMO_MODE = not storage.list_weeks() and DEMO_DIR.exists()
if DEMO_MODE:
    storage.DATA_DIR = DEMO_DIR


# ── 주차 → 월 ────────────────────────────────────────────────────


def week_month(key):
    """주차 키 → (연, 월). 기준은 app.week_label 과 같은 그 주의 일요일이다."""
    parsed = storage.parse_week(key)
    if not parsed:
        return None
    y, w = parsed
    try:
        일요일 = date.fromisocalendar(y, w, 7)
    except ValueError:
        return None
    return 일요일.year, 일요일.month


@st.cache_data(show_spinner=False)
def load_index():
    """저장된 주차를 (연, 월) 로 묶어 둔다. {(연,월): [주차키...]}"""
    index = {}
    for key in storage.list_weeks():
        ym = week_month(key)
        if ym:
            index.setdefault(ym, []).append(key)
    for keys in index.values():
        keys.sort()
    return index


# ── 집계 ─────────────────────────────────────────────────────────


def rows_of(keys, only_approved, 기관=""):
    """여러 주차의 제출을 한 행 목록으로 합친다. _행 은 다시 매긴다(오류 표시용).

    기관 문구를 주면 기관명에 그 말이 들어간 행만 남긴다(대소문자·앞뒤 공백 무시).
    """
    찾을말 = (기관 or "").strip().lower()
    rows = []
    for key in keys:
        data = storage.load_week(key)
        if not data:
            continue
        for r in storage.to_rows(data, only_approved=only_approved):
            if 찾을말 and 찾을말 not in str(r.get("기관명", "")).lower():
                continue
            rows.append({**r, "_행": len(rows) + 2, "_주차": key})
    return rows


def totals_of(keys, only_approved, 기관=""):
    """주차 묶음의 전체 합계 한 줄. 자료가 없으면 None."""
    rows = rows_of(keys, only_approved, 기관)
    if not rows:
        return None
    result = process_rows([{k: v for k, v in r.items() if k != "_주차"} for r in rows])
    t = dict(result["합계_전체"][0])
    t["오류수"] = result["행수"]["오류"]
    t["주차수"] = len(keys)
    return t


@st.cache_data(show_spinner=False)
def 기관예시():
    """입력칸 안내에 쓸 실제 기관명 하나. 저장소가 비면 일반 예시를 쓴다."""
    for key in storage.list_weeks():
        제출 = (storage.load_week(key) or {}).get("제출") or {}
        for org in 제출:
            return org
    return "가온직업전문학교"


def num(v):
    """지표 값 → 그래프·계산에 쓸 실수. 문자열 상수와 None 은 값이 없는 것으로 본다."""
    return None if v is None or isinstance(v, str) else float(v)


def pct_text(v):
    """지표 값 → 사람이 읽는 문구. 문자열 상수는 그대로 살려 둔다."""
    if isinstance(v, str):
        return v  # '검증 필요' · '계산 불가'
    return "자료 없음" if v is None else f"{v * 100:.1f}%"


# ── 화면 ─────────────────────────────────────────────────────────

st.set_page_config(page_title="훈련실적 월별·연도별 보기", page_icon="📊", layout="wide")
st.title("훈련실적 월별·연도별 보기")
st.caption(
    "저장소(data/)에 쌓인 주차 실적을 월·연도 단위로 묶어 봅니다. "
    "숫자는 기존 집계기(feature1_aggregate)를 그대로 통과한 값입니다."
)
if DEMO_MODE:
    st.info("🧪 예시 데이터로 보여 주는 데모 화면입니다. 실제 제출 자료가 아니며, 기관명·수치는 모두 가짜입니다.")

index = load_index()
if not index:
    st.warning("저장된 주차 실적이 없습니다. 먼저 본 앱(python app.py)에서 제출을 받아 주세요.")
    st.stop()

years = sorted({y for y, _ in index}, reverse=True)

# ── 입력칸
with st.form("설정"):
    st.subheader("살펴볼 범위")
    c1, c2, c3 = st.columns([1, 1, 2])
    단위 = c1.radio("보기 단위", ["월별", "연도별"], help="월별은 고른 연도의 달마다, 연도별은 연도마다 묶어 봅니다.")
    연도 = c2.selectbox("연도", years, help="월별로 볼 때 기준이 되는 연도입니다.")
    승인만 = c3.checkbox(
        "관리자 승인분만 집계", value=False,
        help="켜면 상태가 '승인' 인 제출만 셉니다. 끄면 제출된 것을 모두 셉니다.",
    )
    기관 = st.text_input(
        "훈련기관 이름으로 걸러 보기",
        value="", placeholder=f"예: {기관예시()}  (비워 두면 전체 기관)",
        help="기관명에 이 말이 들어간 제출만 집계합니다. 일부만 적어도 됩니다.",
    )
    비교 = st.checkbox(
        "작년 같은 시점과 비교", value=True,
        help="월별이면 작년 같은 달, 연도별이면 각 연도의 직전 연도와 견줍니다.",
    )
    실행 = st.form_submit_button("설정하기", type="primary")

if not 실행:
    st.info("범위를 고른 뒤 **설정하기** 를 누르면 아래에 요약·도표·그래프가 나옵니다.")
    st.stop()

# ── 구간 만들기
if 단위 == "월별":
    구간 = [(f"{연도}-{m:02d}", index.get((연도, m), []), index.get((연도 - 1, m), [])) for m in MONTHS]
    구간 = [x for x in 구간 if x[1] or x[2]]
    비교이름 = f"{연도 - 1}년 같은 달"
else:
    구간 = []
    for y in sorted(years):
        이번 = [k for (yy, _), ks in index.items() if yy == y for k in ks]
        작년 = [k for (yy, _), ks in index.items() if yy == y - 1 for k in ks]
        구간.append((f"{y}년", 이번, 작년))
    비교이름 = "직전 연도"

if not 구간:
    st.warning(f"{연도}년에 해당하는 주차 실적이 없습니다. 다른 연도를 골라 주세요.")
    st.stop()

# ── 집계
표 = []
for 이름, 이번키, 작년키 in 구간:
    t = totals_of(이번키, 승인만, 기관)
    p = totals_of(작년키, 승인만, 기관) if 비교 else None
    행 = {
        "구간": 이름,
        "주차수": len(이번키),
        "집계 과정수": t["과정수"] if t else 0,
        "목표인원": (t.get("목표훈련인원") if t else 0) or 0,
        "훈련수료인원": t["훈련수료인원"] if t else 0,
        "중도탈락자": t["중도탈락자"] if t else 0,
        "수료율": num(t["수료율_목표"]) if t else None,   # 목표 대비 수료
        "탈락률": num(t["탈락률"]) if t else None,
        "제외": t["제외"] if t else 0,
        "_수료율원문": t["수료율_목표"] if t else None,
    }
    if 비교:
        행["작년 수료율"] = num(p["수료율_목표"]) if p else None
        행["작년 수료인원"] = p["훈련수료인원"] if p else 0
        행["증감(%p)"] = (
            round((행["수료율"] - 행["작년 수료율"]) * 100, 1)
            if 행["수료율"] is not None and 행["작년 수료율"] is not None
            else None
        )
    표.append(행)

df = pd.DataFrame(표)
있는구간 = df[df["집계 과정수"] > 0]

# ── 요약문
st.subheader("요약")
걸러냄 = f"기관명에 '{기관.strip()}' 이(가) 들어간 제출만" if 기관.strip() else ""
if 있는구간.empty:
    st.warning(
        "고른 범위에 집계된 과정이 없습니다. "
        + (f"'{기관.strip()}' 로 거른 결과가 비었는지 확인해 보세요. " if 기관.strip() else "")
        + "승인만 보기를 끄거나 다른 연도를 골라 보셔도 됩니다."
    )
else:
    총목표 = int(있는구간["목표인원"].sum())
    총수료 = int(있는구간["훈련수료인원"].sum())
    총과정 = int(있는구간["집계 과정수"].sum())
    총제외 = int(있는구간["제외"].sum())
    전체수료율 = 총수료 / 총목표 if 총목표 else None
    최고 = 있는구간.loc[있는구간["수료율"].idxmax()] if 있는구간["수료율"].notna().any() else None
    최저 = 있는구간.loc[있는구간["수료율"].idxmin()] if 있는구간["수료율"].notna().any() else None

    문장 = [
        f"**{단위}** 기준으로 {len(있는구간)}개 구간, 과정 {총과정:,}개를 집계했습니다."
        + (f" ({걸러냄})" if 걸러냄 else ""),
        f"목표 {총목표:,}명 대비 수료 {총수료:,}명 — 전체 수료율(목표 대비) **{pct_text(전체수료율)}** 입니다.",
    ]
    if 최고 is not None and 최저 is not None and len(있는구간) > 1:
        문장.append(
            f"가장 높은 구간은 {최고['구간']}({최고['수료율'] * 100:.1f}%), "
            f"가장 낮은 구간은 {최저['구간']}({최저['수료율'] * 100:.1f}%) 입니다."
        )
    if 비교 and "증감(%p)" in 있는구간 and 있는구간["증감(%p)"].notna().any():
        비교가능 = 있는구간[있는구간["증감(%p)"].notna()]
        오른곳 = int((비교가능["증감(%p)"] > 0).sum())
        내린곳 = int((비교가능["증감(%p)"] < 0).sum())
        문장.append(
            f"{비교이름}과 견줄 수 있는 구간은 {len(비교가능)}개이며, "
            f"수료율이 오른 곳 {오른곳}개 · 내린 곳 {내린곳}개입니다."
        )
    elif 비교:
        문장.append(f"{비교이름} 자료가 저장소에 없어 비교는 생략했습니다.")
    if 총제외:
        문장.append(f"입력 오류로 합계에서 빠진 과정이 {총제외}개 있어, 위 수치는 확정 전 값입니다.")

    st.markdown("\n\n".join(문장))

    m1, m2, m3 = st.columns(3)
    m1.metric("전체 수료율(목표 대비)", pct_text(전체수료율))
    m2.metric("수료인원", f"{총수료:,}명")
    m3.metric("집계 과정수", f"{총과정:,}개", delta=f"제외 {총제외}개" if 총제외 else None, delta_color="inverse")

# ── 도표
st.subheader("도표")
보일칸 = ["구간", "주차수", "집계 과정수", "목표인원", "훈련수료인원", "중도탈락자", "수료율", "탈락률", "제외"]
if 비교:
    보일칸 += ["작년 수료율", "증감(%p)"]
st.dataframe(
    df[보일칸],
    hide_index=True,
    width="stretch",
    column_config={
        "수료율": st.column_config.NumberColumn("수료율(목표대비)", format="percent"),
        "탈락률": st.column_config.NumberColumn("탈락률", format="percent"),
        "작년 수료율": st.column_config.NumberColumn("작년 수료율", format="percent"),
        "증감(%p)": st.column_config.NumberColumn("증감(%p)", format="%.1f"),
    },
)
if (df["_수료율원문"] == UNVERIFIED).any() or (df["_수료율원문"] == UNCOMPUTABLE).any():
    st.caption("일부 구간은 입력 오류(검증 필요)나 분모 0(계산 불가)이라 수료율이 비어 있습니다.")

# ── 세로 막대 그래프
st.subheader("세로 막대 그래프")
탭1, 탭2 = st.tabs(["수료율", "수료인원"])

with 탭1:
    칼럼 = ["수료율"] + (["작년 수료율"] if 비교 and df["작년 수료율"].notna().any() else [])
    그래프 = df.set_index("구간")[칼럼].mul(100).round(1)
    st.bar_chart(그래프, height=380, y_label="수료율(목표 대비) (%)", x_label="구간")
    if len(칼럼) == 1 and 비교:
        st.caption(f"{비교이름} 자료가 없어 이번 값만 그렸습니다.")

with 탭2:
    칼럼 = ["훈련수료인원"] + (["작년 수료인원"] if 비교 and df.get("작년 수료인원", pd.Series()).sum() else [])
    st.bar_chart(df.set_index("구간")[칼럼], height=380, y_label="수료인원 (명)", x_label="구간")

with st.expander("이 화면이 쓰는 자료"):
    st.write(
        f"- 저장소 주차 파일 {sum(len(v) for v in index.values())}개 "
        f"({min(f'{y}-{m:02d}' for y, m in index)} ~ {max(f'{y}-{m:02d}' for y, m in index)})"
    )
    st.write(f"- 집계 범위: {'승인된 제출만' if 승인만 else '제출된 모든 건'}")
    st.write(f"- 기관 걸러내기: {걸러냄 or '전체 기관'}")
    st.write("- 수료율(목표 대비) = 훈련수료인원 ÷ 목표(연간 목표 또는 정원), 탈락률 = 중도탈락 ÷ 실시 (집계기에서 계산)")
    st.write("- 이 화면은 저장소를 읽기만 하며, 제출·승인 자료를 고치지 않습니다.")
