"""핵심 기능 1 — 실적 엑셀 검증·표준화·집계

specs/feature-1-spec.md 명세 구현.
실적 엑셀 1건을 읽어 형식 확인 → 표준화 → 지표 계산 → 합계 산출 후
통합 실적표 · 합계 · 오류 목록을 출력한다.

사용법:
    python feature1_aggregate.py [엑셀파일경로]
    (경로 생략 시 inputs/sample-training-data-2026-W30.xlsx)
"""

import sys
import unicodedata
from pathlib import Path

from openpyxl import load_workbook

SHEET = "실적"
COLS = [
    "기관명",
    "과정명",
    "NCS분류",
    "KECO분류",
    "훈련목표인원",
    "훈련실시인원",
    "훈련수료인원",
    "중도탈락자",
]
COUNT_COLS = COLS[4:]  # 인원 4개 항목
HEADER_ROW = 1  # 1행 = 머리글, 데이터는 2행부터

BLANK = "(미기재)"  # 분류값이 비어 있을 때 합계에서 쓰는 이름
UNVERIFIED = "검증 필요"  # 오류 행의 지표 (계산해도 신뢰할 수 없음)
UNCOMPUTABLE = "계산 불가"  # 분모가 0이라 계산 자체가 안 되는 경우


# ── 표 출력 보조 (한글 폭 보정) ──────────────────────────────────


def width(s):
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in str(s))


def pad(s, w, align="left"):
    gap = " " * max(0, w - width(s))
    return gap + str(s) if align == "right" else str(s) + gap


def table(headers, rows, aligns=None, rule_before=None):
    """rule_before: 해당 인덱스 행 앞에 구분선을 넣는다 (합계 행 구분용)."""
    aligns = aligns or ["left"] * len(headers)
    cells = [headers] + [[str(c) for c in r] for r in rows]
    widths = [max(width(row[i]) for row in cells) for i in range(len(headers))]
    line = "─┼─".join("─" * w for w in widths)

    out = [" │ ".join(pad(h, w) for h, w in zip(headers, widths)), line]
    for i, r in enumerate(rows):
        if rule_before is not None and i == rule_before:
            out.append(line)
        out.append(" │ ".join(pad(c, w, a) for c, w, a in zip(r, widths, aligns)))
    return "\n".join(out)


def pct(v):
    """지표 표시. 문자열(검증 필요/계산 불가)은 그대로 통과시킨다."""
    if isinstance(v, str):
        return v
    return "입력 없음" if v is None else f"{v * 100:.1f}%"


def num(v, raw=None):
    """인원 표시. 읽지 못한 값은 원본을 그대로 보여준다."""
    if v is not None:
        return v
    return f"'{raw}'" if raw not in (None, "") else ""


# ── 1. 읽기 ──────────────────────────────────────────────────────


def read_rows(path):
    """엑셀 「실적」 시트를 행 단위로 읽는다. (머리글 검증 포함)"""
    wb = load_workbook(path, data_only=True)
    if SHEET not in wb.sheetnames:
        raise SystemExit(f"오류: 「{SHEET}」 시트가 없습니다. (시트 목록: {', '.join(wb.sheetnames)})")

    ws = wb[SHEET]
    try:
        first = next(ws.iter_rows(values_only=True))
    except StopIteration:
        raise SystemExit(f"오류: 「{SHEET}」 시트가 비어 있습니다. 머리글 행부터 채워 주세요.")
    header = [str(c).strip() if c is not None else "" for c in first]

    missing = [c for c in COLS if c not in header]
    if missing:
        raise SystemExit(f"오류: 필수 컬럼이 없습니다 → {', '.join(missing)}")

    idx = {c: header.index(c) for c in COLS}
    rows = []
    for n, values in enumerate(ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True), start=HEADER_ROW + 1):
        if all(v is None or str(v).strip() == "" for v in values):
            continue  # 빈 행은 건너뜀
        rows.append({"_행": n, **{c: values[idx[c]] for c in COLS}})
    if not rows:
        raise SystemExit(f"오류: 「{SHEET}」 시트에 데이터 행이 없습니다. 머리글 아래에 실적을 채워 주세요.")
    return rows


# ── 2~3. 형식 확인 + 표준화 ──────────────────────────────────────


def to_int(v):
    """인원 값을 정수로 변환. 불가하면 None."""
    if v is None or str(v).strip() == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else None
    s = str(v).strip().replace(",", "").replace("명", "")
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if f.is_integer() else None


def check_and_normalize(row):
    """한 행을 검사·표준화한다. (표준화된 행, 오류목록) 반환."""
    errors = []
    n = row["_행"]
    out = {"_행": n, "_원본": {}}

    # 문자 항목: 공백 정리 + 필수값 확인
    for c in COLS[:4]:
        v = "" if row[c] is None else str(row[c]).strip()
        if not v:
            errors.append((n, c, "필수값 누락"))
        out[c] = v

    # 인원 항목: 숫자 변환 + 음수 확인
    for c in COUNT_COLS:
        raw = row[c]
        out["_원본"][c] = raw
        if raw is None or str(raw).strip() == "":
            errors.append((n, c, "필수값 누락"))
            out[c] = None
            continue
        v = to_int(raw)
        if v is None:
            errors.append((n, c, f"숫자로 읽을 수 없음 (값: '{raw}')"))
        elif v < 0:
            errors.append((n, c, f"음수는 올 수 없음 (값: {v})"))
            v = None
        out[c] = v

    # 논리 검사: 값이 둘 다 정상일 때만
    sil, su, tal = out["훈련실시인원"], out["훈련수료인원"], out["중도탈락자"]
    if sil is not None and su is not None and su > sil:
        errors.append((n, "훈련수료인원", f"실시인원({sil})보다 많음 (수료 {su})"))
    if sil is not None and tal is not None and tal > sil:
        errors.append((n, "중도탈락자", f"실시인원({sil})보다 많음 (탈락 {tal})"))

    return out, errors


# ── 4. 지표 계산 ─────────────────────────────────────────────────


def ratio(numerator, denominator):
    """분모가 0이면 '계산 불가', 값 자체가 없으면 None.

    둘을 구분해야 화면에서 '나눌 수 없음'과 '입력 없음'이 섞이지 않는다.
    """
    if denominator == 0:
        return UNCOMPUTABLE
    if numerator is None or denominator is None:
        return None
    return numerator / denominator


def add_metrics(row):
    """오류 행은 값이 신뢰할 수 없으므로 지표를 계산하지 않는다."""
    if row["_오류"]:
        row["이수율"] = UNVERIFIED
        row["탈락률"] = UNVERIFIED
    else:
        row["이수율"] = ratio(row["훈련수료인원"], row["훈련목표인원"])
        row["탈락률"] = ratio(row["중도탈락자"], row["훈련실시인원"])
    return row


# ── 5. 합계 산출 ─────────────────────────────────────────────────


def summarize(clean, dropped, key):
    """key 기준 합계. clean=집계 대상, dropped=오류로 빠진 행.

    구분 순서는 파일에 나온 순서를 따른다.
    """
    order, groups, drops = [], {}, {}
    for r in clean + dropped:
        name = r.get(key) or BLANK
        if name not in order:
            order.append(name)
    for r in clean:
        groups.setdefault(r.get(key) or BLANK, []).append(r)
    for r in dropped:
        drops[r.get(key) or BLANK] = drops.get(r.get(key) or BLANK, 0) + 1

    out = []
    for name in order:
        items = groups.get(name, [])
        s = {c: sum(i[c] for i in items) for c in COUNT_COLS}
        out.append(
            {
                "구분": name,
                "과정수": len(items),
                "제외": drops.get(name, 0),
                **s,
                "이수율": ratio(s["훈련수료인원"], s["훈련목표인원"]),
                "탈락률": ratio(s["중도탈락자"], s["훈련실시인원"]),
            }
        )
    return out


# ── 6. 결과 반환 ─────────────────────────────────────────────────


def process(path):
    """명세의 동작 1~6을 순서대로 수행하고 결과를 돌려준다."""
    return process_rows(read_rows(path))  # 1. 읽기


def process_rows(raw):
    """읽어들인 행 목록에 동작 2~6을 수행한다.

    입력 출처(엑셀 파일 / 화면 직접 입력)와 무관하게 같은 규칙을 적용하기 위해
    읽기(1단계)와 분리했다. 각 행은 _행 과 COLS 키를 가진 dict.
    """
    rows, errors = [], []
    for r in raw:
        norm, errs = check_and_normalize(r)  # 2. 형식 확인 + 3. 표준화
        norm["_오류"] = bool(errs)
        rows.append(add_metrics(norm))  # 4. 지표 계산
        errors.extend(errs)

    clean = [r for r in rows if not r["_오류"]]  # 5. 합계는 오류 없는 행만
    dropped = [r for r in rows if r["_오류"]]

    total = summarize([{**r, "_전체": "전체"} for r in clean], [{**r, "_전체": "전체"} for r in dropped], "_전체")
    if not total:  # 행이 한 건도 없을 때도 합계 한 줄은 있어야 소비하는 쪽이 안 깨진다
        total = [{"구분": "전체", "과정수": 0, "제외": 0, **{c: 0 for c in COUNT_COLS},
                  "이수율": UNCOMPUTABLE, "탈락률": UNCOMPUTABLE}]

    return {
        "표": rows,  # 오류 행 포함 (지표는 '검증 필요')
        "합계_전체": total,
        "합계_기관별": summarize(clean, dropped, "기관명"),
        "합계_NCS별": summarize(clean, dropped, "NCS분류"),
        "합계_KECO별": summarize(clean, dropped, "KECO분류"),
        "오류": errors,
        "행수": {"전체": len(rows), "정상": len(clean), "오류": len(dropped)},
    }


# ── 출력 ─────────────────────────────────────────────────────────


def show(result, path):
    R = "right"
    c = result["행수"]
    print(f"\n입력 파일: {path}")
    print(f"읽은 행: {c['전체']}행 (정상 {c['정상']} / 오류 {c['오류']})")

    # ① 통합 실적표 (맨 아래에 합계 행)
    print("\n" + "=" * 60)
    print("■ 통합 실적표")
    print("=" * 60)
    rows = [
        [
            "⚠" if r["_오류"] else "",
            r["_행"],
            r["기관명"],
            r["과정명"],
            r["NCS분류"],
            r["KECO분류"],
            *(num(r[col], r["_원본"].get(col)) for col in COUNT_COLS),
            pct(r["이수율"]),
            pct(r["탈락률"]),
        ]
        for r in result["표"]
    ]
    t = result["합계_전체"][0]
    rows.append(
        [
            "",
            "",
            "합계",
            f"정상 {t['과정수']}행" + (f" · 제외 {t['제외']}행" if t["제외"] else ""),
            "",
            "",
            *(t[col] for col in COUNT_COLS),
            pct(t["이수율"]),
            pct(t["탈락률"]),
        ]
    )
    print(
        table(
            ["", "행", "훈련기관", "과정명", "NCS", "KECO", "목표", "실시", "수료", "탈락", "이수율", "탈락률"],
            rows,
            ["left", R, "left", "left", "left", "left", R, R, R, R, R, R],
            rule_before=len(rows) - 1,
        )
    )
    if c["오류"]:
        print(f"\n⚠ 표시 {c['오류']}행은 오류가 있어 지표를 계산하지 않고 합계에서도 제외됨 (오류 목록 참고)")

    # ② 구분별 합계
    for title, key in [
        ("기관별 합계", "합계_기관별"),
        ("NCS분류별 합계", "합계_NCS별"),
        ("KECO분류별 합계", "합계_KECO별"),
    ]:
        print("\n" + "=" * 60)
        print(f"■ {title}")
        print("=" * 60)
        rows = [
            [
                g["구분"],
                g["과정수"],
                g["제외"] or "",
                *(g[col] for col in COUNT_COLS),
                pct(g["이수율"]),
                pct(g["탈락률"]),
            ]
            for g in result[key]
        ]
        print(
            table(
                ["구분", "집계 과정수", "제외", "목표", "실시", "수료", "탈락", "이수율", "탈락률"],
                rows,
                ["left", R, R, R, R, R, R, R, R],
            )
        )
        if any(g["제외"] for g in result[key]):
            print("* '제외'는 오류로 집계에서 빠진 과정 수. 값이 있는 구분은 다른 구분과 단순 비교 시 주의.")

    # ③ 오류 목록
    print("\n" + "=" * 60)
    print(f"■ 오류 목록  ({len(result['오류'])}건)")
    print("=" * 60)
    if not result["오류"]:
        print("오류 없음")
    else:
        print(table(["행", "컬럼", "사유"], [[n, col, m] for n, col, m in result["오류"]], [R, "left", "left"]))
    print()


if __name__ == "__main__":
    default = Path(__file__).parent / "inputs" / "sample-training-data-2026-W30.xlsx"
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else default
    if not target.exists():
        raise SystemExit(f"오류: 파일을 찾을 수 없습니다 → {target}")
    show(process(target), target)
