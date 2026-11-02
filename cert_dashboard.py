"""인증 프로젝트 현황 대시보드 (교육용 가상 데이터)

실행: streamlit run cert_dashboard.py
필요 패키지: streamlit, pandas, openpyxl
"""
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill

# ---------------------------------------------------------------- 상수
DATA_DIR = Path(__file__).parent / "data"
PROJECT_FILE = DATA_DIR / "SB_project_demo.csv"
# 요청서에는 SB_resources_demo.csv 로 적혀 있어 두 이름 모두 찾는다
RESOURCE_FILES = [DATA_DIR / "SB_resources_demo.csv", DATA_DIR / "SB_resource_demo.csv"]

FILTER_COLUMNS = ["시험분야", "적용규격", "담당엔지니어", "진행상태", "시료상태"]
INITIAL_FIELDS = ["EMC", "무선(RF)", "인증서류", "SAR"]
MISSING = "(미입력)"

STATUS_DONE = "완료"
STATUS_DELAY = "지연 후보(일정 확인 필요)"
STATUS_TODAY = "오늘 마감"
STATUS_PLANNED = "진행 예정"
STATUS_CHECK = "일정 확인 필요"
SCHEDULE_STATUSES = [STATUS_DELAY, STATUS_TODAY, STATUS_PLANNED, STATUS_CHECK, STATUS_DONE]

EMPTY_MESSAGE = "조건에 맞는 인증 프로젝트가 없습니다"
EXCEL_NOTICE = "교육용 검토 파일 — 공식 보고서 아님"
SESSION_NOTICE = "카테고리·메모 변경 내용은 현재 브라우저 세션에만 보관되며, 새로고침하거나 세션이 끝나면 사라질 수 있습니다."
RESOURCE_NOTICE = "현황 조회용이며 예약·시험 재개를 승인하는 기능이 아닙니다."

TABLE_COLUMNS = [
    "의뢰번호", "고객사", "제품명", "모델명", "시험분야", "적용규격", "담당엔지니어",
    "접수일", "완료예정일", "잔여일수", "일정 상태", "진행상태", "시료상태", "메모",
]

RESOURCE_STYLE = {  # 상태별 카드 색상 (배경, 글자)
    "사용 가능": ("#e3f4e8", "#1b6b34"),
    "사용 중": ("#e3edfb", "#1d4f91"),
    "점검 중": ("#fdf0d9", "#8a5300"),
}

ss = st.session_state


# ---------------------------------------------------------------- 데이터 읽기
@st.cache_data
def read_csv(path: Path) -> pd.DataFrame:
    """CSV를 문자열로 읽고 열 이름·값의 앞뒤 공백을 제거한다."""
    df = pd.read_csv(path, encoding="utf-8-sig", dtype=str)
    df.columns = df.columns.str.strip()
    for col in df.columns:
        df[col] = df[col].str.strip()
    return df


def load_projects() -> pd.DataFrame:
    """의뢰 데이터를 읽고, 필터 열의 빈 값을 '(미입력)'으로 채운다."""
    df = read_csv(PROJECT_FILE).copy()
    for col in FILTER_COLUMNS:
        df[col] = df[col].fillna(MISSING).replace("", MISSING)
    df["메모"] = df["메모"].fillna("")
    return df


def load_resources() -> pd.DataFrame | None:
    """시험 자원 CSV를 읽는다. 파일이 없으면 None."""
    for path in RESOURCE_FILES:
        if path.exists():
            return read_csv(path)
    return None


def init_state() -> None:
    """세션 최초 실행 시 편집용 데이터·카테고리·필터 선택값을 준비한다."""
    if "projects" not in ss:
        ss.projects = load_projects()
    if "fields" not in ss:
        extra = [f for f in ss.projects["시험분야"].unique() if f not in INITIAL_FIELDS]
        ss.fields = INITIAL_FIELDS + sorted(extra)
    for col in FILTER_COLUMNS:
        key = filter_key(col)
        if key not in ss:
            ss[key] = filter_options(col)


# ---------------------------------------------------------------- 일정 계산
def schedule_status(progress: str, days) -> str:
    """진행상태와 잔여일수로 일정 상태를 정한다. (완료가 최우선)"""
    if progress == "완료":
        return STATUS_DONE
    if pd.isna(days):
        return STATUS_CHECK
    if days < 0:
        return STATUS_DELAY
    if days == 0:
        return STATUS_TODAY
    return STATUS_PLANNED


def add_schedule_columns(df: pd.DataFrame, base_date: date) -> pd.DataFrame:
    """잔여일수(완료예정일 - 기준일)와 일정 상태 열을 추가한다."""
    df = df.copy()
    # 비어 있거나 YYYY-MM-DD 형식이 아니면 NaT → 잔여일수 비어 있음
    due = pd.to_datetime(df["완료예정일"], format="%Y-%m-%d", errors="coerce")
    df["잔여일수"] = (due - pd.Timestamp(base_date)).dt.days.astype("Int64")
    df["일정 상태"] = [schedule_status(p, d) for p, d in zip(df["진행상태"], df["잔여일수"])]
    return df


# ---------------------------------------------------------------- 필터
def filter_key(col: str) -> str:
    """필터 위젯의 session_state 키."""
    return f"flt_{col}"


def filter_options(col: str) -> list[str]:
    """필터 선택지. 시험분야는 관리 중인 카테고리 목록을 쓴다."""
    if col == "시험분야":
        return list(ss.fields)
    return sorted(ss.projects[col].unique())


def apply_filters(df: pd.DataFrame, selections: dict[str, list[str]]) -> pd.DataFrame:
    """모든 필터 조건(AND)을 만족하는 행만 남긴다."""
    mask = pd.Series(True, index=df.index)
    for col, selected in selections.items():
        mask &= df[col].isin(selected)
    return df[mask]


def describe_selection(col: str, selected: list[str]) -> str:
    """필터 선택값을 사람이 읽을 문장으로 바꾼다 (Excel 분석조건 시트용)."""
    options = filter_options(col)
    if set(selected) >= set(options):
        return "전체"
    if not selected:
        return "선택 없음"
    return ", ".join(selected)


# ---------------------------------------------------------------- 카테고리 관리 (버튼 콜백)
def flash(kind: str, text: str) -> None:
    """다음 화면 갱신 때 보여 줄 알림을 저장한다."""
    ss.flash = (kind, text)


def field_usage(name: str) -> int:
    """해당 시험분야를 쓰는 의뢰 건수."""
    return int((ss.projects["시험분야"] == name).sum())


def on_add_field() -> None:
    """시험분야 등록."""
    name = ss.cat_add_name.strip()
    if not name:
        flash("warning", "등록할 시험분야 이름을 입력하세요.")
    elif name in ss.fields:
        flash("warning", f"'{name}'은(는) 이미 있는 시험분야입니다.")
    else:
        ss.fields.append(name)
        ss[filter_key("시험분야")].append(name)  # 새 분야도 필터에 포함
        ss.cat_add_name = ""
        flash("success", f"시험분야 '{name}'을(를) 등록했습니다.")


def on_rename_field() -> None:
    """시험분야 이름 변경 + 관련 의뢰 데이터·필터 선택값 동시 갱신."""
    old, new = ss.cat_rename_old, ss.cat_rename_new.strip()
    if not new:
        flash("warning", "새 이름을 입력하세요.")
        return
    if new in ss.fields:
        flash("warning", f"'{new}'은(는) 이미 있는 시험분야입니다.")
        return
    ss.fields[ss.fields.index(old)] = new
    moved = field_usage(old)
    ss.projects.loc[ss.projects["시험분야"] == old, "시험분야"] = new
    key = filter_key("시험분야")
    ss[key] = [new if v == old else v for v in ss[key]]
    ss.cat_rename_old = new
    ss.cat_rename_new = ""
    flash("success", f"'{old}' → '{new}'로 변경했습니다. (의뢰 {moved}건 함께 갱신)")


def on_reclassify_field() -> None:
    """삭제하려는 분야의 의뢰를 다른 분야로 옮긴다."""
    src, dst = ss.cat_delete, ss.cat_reclass_to
    moved = field_usage(src)
    ss.projects.loc[ss.projects["시험분야"] == src, "시험분야"] = dst
    flash("success", f"'{src}' 의뢰 {moved}건을 '{dst}'(으)로 재분류했습니다. 이제 삭제할 수 있습니다.")


def on_delete_field() -> None:
    """사용 중이 아닌 시험분야만 삭제한다."""
    name = ss.cat_delete
    if field_usage(name) > 0:
        flash("error", f"'{name}'은(는) 사용 중이라 삭제할 수 없습니다. 먼저 재분류하세요.")
        return
    ss.fields.remove(name)
    key = filter_key("시험분야")
    ss[key] = [v for v in ss[key] if v != name]
    flash("success", f"시험분야 '{name}'을(를) 삭제했습니다.")


def on_save_memo(request_id: str) -> None:
    """상세 영역에서 수정한 메모를 의뢰 데이터에 저장한다."""
    text = ss[f"memo_input_{request_id}"].strip()
    ss.projects.loc[ss.projects["의뢰번호"] == request_id, "메모"] = text
    ss.memo_flash = f"{request_id} 메모를 저장했습니다."


# ---------------------------------------------------------------- 집계
def status_summary(df: pd.DataFrame) -> pd.DataFrame:
    """일정 상태·진행상태별 건수와 비율 표."""
    parts = []
    for col in ["일정 상태", "진행상태"]:
        counts = df[col].value_counts().rename_axis("상태").reset_index(name="건수")
        counts.insert(0, "구분", col)
        parts.append(counts)
    out = pd.concat(parts, ignore_index=True)
    total = len(df)
    out["비율"] = (out["건수"] / total).round(3) if total else 0
    return out


def engineer_summary(df: pd.DataFrame) -> pd.DataFrame:
    """담당엔지니어별 전체 건수와 일정 상태별 건수 표."""
    table = pd.crosstab(df["담당엔지니어"], df["일정 상태"])
    table = table.reindex(columns=SCHEDULE_STATUSES, fill_value=0)
    table.insert(0, "전체", table.sum(axis=1))
    return table.sort_values("전체", ascending=False).reset_index()


# ---------------------------------------------------------------- Excel
def build_excel(df: pd.DataFrame, scope: str, base_date: date,
                conditions: dict[str, str], created_at: datetime) -> bytes:
    """4개 시트(프로젝트목록/상태별집계/담당자별집계/분석조건)로 된 xlsx를 만든다."""
    info = [("데이터 범위", scope), ("기준일", f"{base_date:%Y-%m-%d}"),
            ("파일 생성 일시", f"{created_at:%Y-%m-%d %H:%M:%S}"), ("건수", f"{len(df)}건")]
    info += [(f"필터: {k}", v) for k, v in conditions.items()]
    info.append(("데이터 출처", "교육용 가상 데이터"))
    sheets = {
        "프로젝트목록": df[TABLE_COLUMNS],
        "상태별집계": status_summary(df),
        "담당자별집계": engineer_summary(df),
        "분석조건": pd.DataFrame(info, columns=["항목", "값"]),
    }
    header_fill = PatternFill("solid", fgColor="E7ECF3")
    delay_fill = PatternFill("solid", fgColor="FDE2E2")

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, data in sheets.items():
            data.to_excel(writer, sheet_name=name, index=False, startrow=2)
            ws = writer.sheets[name]
            ws["A1"] = EXCEL_NOTICE
            ws["A1"].font = Font(bold=True, color="C00000", size=12)
            for cell in ws[3]:  # 머리글 행
                cell.font = Font(bold=True)
                cell.fill = header_fill
            for i, col in enumerate(data.columns, start=1):
                width = max([len(str(col))] + [len(str(v)) for v in data[col]]) * 1.8 + 2
                ws.column_dimensions[ws.cell(row=3, column=i).column_letter].width = min(width, 50)
            ws.freeze_panes = "A4"
        # 지연 후보 행 강조
        ws = writer.sheets["프로젝트목록"]
        status_col = TABLE_COLUMNS.index("일정 상태")
        for row in ws.iter_rows(min_row=4):
            if row[status_col].value == STATUS_DELAY:
                for cell in row:
                    cell.fill = delay_fill
    return buf.getvalue()


# ---------------------------------------------------------------- 화면 구성 요소
def card(title: str, value: str, bg: str = "", fg: str = "", note: str = "") -> None:
    """테두리 있는 상태 카드 하나를 그린다."""
    with st.container(border=True):
        st.caption(title)
        if bg:
            st.markdown(
                f"<span style='background:{bg};color:{fg};padding:4px 12px;"
                f"border-radius:999px;font-weight:600'>{value}</span>",
                unsafe_allow_html=True,
            )
        else:
            st.markdown(f"#### {value}")
        if note:
            st.caption(note)


def render_resource_cards(resources: pd.DataFrame | None) -> None:
    """시험 자원 상태 카드. CSV 상태를 그대로 보여 주며 코드로 바꾸지 않는다."""
    st.subheader("시험 자원 현황")
    if resources is None:
        st.warning("자원 데이터 파일(SB_resources_demo.csv)을 찾을 수 없습니다.")
        return
    cols = st.columns(len(resources))
    for col, (_, row) in zip(cols, resources.iterrows()):
        status = row["상태"] if pd.notna(row["상태"]) else MISSING
        bg, fg = RESOURCE_STYLE.get(status, ("#eeeeee", "#444444"))
        with col:
            card(row["자원"], status, bg, fg, row.get("비고", "") or "")
    st.caption(f"ⓘ {RESOURCE_NOTICE}")


def render_summary_cards(df: pd.DataFrame) -> None:
    """필터 결과 기준 요약 지표."""
    counts = df["일정 상태"].value_counts()
    items = [("프로젝트 수", len(df)), ("완료", counts.get(STATUS_DONE, 0)),
             ("지연 후보", counts.get(STATUS_DELAY, 0)), ("오늘 마감", counts.get(STATUS_TODAY, 0)),
             ("일정 확인 필요", counts.get(STATUS_CHECK, 0))]
    for col, (label, value) in zip(st.columns(len(items)), items):
        col.metric(label, f"{value}건")


def render_charts(df: pd.DataFrame) -> None:
    """시험분야별 막대그래프와 진행상태 도넛 차트."""
    left, right = st.columns(2)
    with left:
        st.subheader("시험분야별 건수")
        data = df["시험분야"].value_counts().rename_axis("시험분야").reset_index(name="건수")
        bar = (
            alt.Chart(data)
            .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("시험분야:N", sort="-y", axis=alt.Axis(labelAngle=0)),
                y=alt.Y("건수:Q", axis=alt.Axis(tickMinStep=1)),
                tooltip=["시험분야", "건수"],
            )
            .properties(height=320)
        )
        st.altair_chart(bar + bar.mark_text(dy=-8).encode(text="건수:Q"), width="stretch")
    with right:
        st.subheader("진행상태 비율")
        data = df["진행상태"].value_counts().rename_axis("진행상태").reset_index(name="건수")
        data["비율"] = data["건수"] / data["건수"].sum()
        donut = (
            alt.Chart(data)
            .mark_arc(innerRadius=70, stroke="white")
            .encode(
                theta=alt.Theta("건수:Q"),
                color=alt.Color("진행상태:N"),
                tooltip=["진행상태", "건수", alt.Tooltip("비율:Q", format=".0%")],
            )
            .properties(height=320)
        )
        st.altair_chart(donut, width="stretch")


def highlight_delay(row: pd.Series) -> list[str]:
    """지연 후보 행을 빨간색 계열로 칠한다."""
    style = "background-color: #fde2e2; color: #9b1c1c" if row["일정 상태"] == STATUS_DELAY else ""
    return [style] * len(row)


def render_table(df: pd.DataFrame) -> None:
    """프로젝트 목록 표."""
    st.subheader("프로젝트 목록")
    table = df[TABLE_COLUMNS].reset_index(drop=True)
    st.dataframe(table.style.apply(highlight_delay, axis=1), width="stretch", hide_index=True)
    st.caption("빨간색 행: 지연 후보 (완료가 아니면서 완료예정일이 기준일보다 지난 프로젝트)")


def render_downloads(full: pd.DataFrame, filtered: pd.DataFrame, base_date: date,
                     selections: dict[str, list[str]]) -> None:
    """전체 목록 / 현재 필터 결과 Excel 다운로드 버튼."""
    st.subheader("Excel 다운로드")
    now = datetime.now()
    stamp = f"{now:%Y%m%d}"
    no_filter = {col: "적용 안 함(전체)" for col in FILTER_COLUMNS}
    applied = {col: describe_selection(col, sel) for col, sel in selections.items()}
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    left, right = st.columns(2)
    left.download_button(
        f"전체 목록 다운로드 ({len(full)}건)",
        data=build_excel(full, "전체 목록", base_date, no_filter, now),
        file_name=f"standardbank_project_status_{stamp}.xlsx",
        mime=mime, width="stretch",
    )
    right.download_button(
        f"현재 필터 결과 다운로드 ({len(filtered)}건)",
        data=build_excel(filtered, "현재 필터 결과", base_date, applied, now) if len(filtered) else b"",
        file_name=f"standardbank_project_status_filtered_{stamp}.xlsx",
        mime=mime, width="stretch", disabled=filtered.empty,
    )


def render_detail(df: pd.DataFrame) -> None:
    """의뢰번호를 골라 상세 정보를 보고 메모를 수정한다."""
    st.subheader("의뢰 상세 조회")
    request_id = st.selectbox("의뢰번호", df["의뢰번호"].tolist(), key="detail_id")
    row = df[df["의뢰번호"] == request_id].iloc[0]

    fields = [("고객사", row["고객사"]), ("제품 · 모델", f"{row['제품명']} / {row['모델명']}"),
              ("시험분야", row["시험분야"]), ("적용규격", row["적용규격"]),
              ("담당엔지니어", row["담당엔지니어"]), ("시료상태", row["시료상태"]),
              ("완료예정일", row["완료예정일"] if pd.notna(row["완료예정일"]) else MISSING),
              ("일정 상태", row["일정 상태"])]
    cols = st.columns(4)
    for i, (label, value) in enumerate(fields):
        with cols[i % 4]:
            card(label, str(value))

    memo_key = f"memo_input_{request_id}"
    if memo_key not in ss:
        ss[memo_key] = row["메모"]
    st.text_area("담당자 메모", key=memo_key, height=90)
    st.button("메모 저장", on_click=on_save_memo, args=(request_id,), type="primary")
    if msg := ss.pop("memo_flash", None):
        st.success(msg)


def render_category_manager() -> None:
    """사이드바 '카테고리 관리' 영역: 시험분야 등록/이름 변경/삭제."""
    st.header("🗂️ 카테고리 관리")
    st.caption("시험분야")
    if note := ss.pop("flash", None):
        getattr(st, note[0])(note[1])

    usage = ", ".join(f"{f}({field_usage(f)})" for f in ss.fields)
    st.caption(f"현재 분야(사용 건수): {usage}")

    tab_add, tab_rename, tab_delete = st.tabs(["등록", "이름 변경", "삭제"])
    with tab_add:
        st.text_input("새 시험분야", key="cat_add_name", placeholder="예: 안전")
        st.button("등록", on_click=on_add_field, width="stretch")
    with tab_rename:
        st.selectbox("변경할 분야", ss.fields, key="cat_rename_old")
        st.text_input("새 이름", key="cat_rename_new")
        st.button("이름 변경", on_click=on_rename_field, width="stretch")
        st.caption("관련 의뢰 데이터의 시험분야도 함께 바뀝니다.")
    with tab_delete:
        target = st.selectbox("삭제할 분야", ss.fields, key="cat_delete")
        in_use = field_usage(target)
        if in_use:
            st.warning(f"'{target}'은(는) 의뢰 {in_use}건에서 사용 중이라 바로 삭제할 수 없습니다. "
                       "아래에서 다른 분야로 재분류한 뒤 삭제하세요.")
            others = [f for f in ss.fields if f != target]
            st.selectbox("재분류할 분야", others, key="cat_reclass_to")
            st.button("재분류", on_click=on_reclassify_field, width="stretch",
                      disabled=not others)
        st.button("삭제", on_click=on_delete_field, width="stretch",
                  disabled=bool(in_use) or len(ss.fields) <= 1)

    st.info(SESSION_NOTICE, icon="💾")


# ---------------------------------------------------------------- 메인
st.set_page_config(page_title="인증 프로젝트 대시보드", layout="wide")
init_state()

with st.sidebar:
    st.header("🔍 필터")
    base_date = st.date_input("기준일", value=date.today())
    selections = {}
    for col in FILTER_COLUMNS:
        selections[col] = st.multiselect(col, filter_options(col), key=filter_key(col))
    st.caption("선택을 모두 지우면 해당 조건에 맞는 프로젝트가 없는 것으로 처리됩니다.")
    st.divider()
    render_category_manager()

st.title("인증 프로젝트 현황 대시보드")
st.info("교육용 가상 데이터 기준입니다. 실제 고객·시험 정보가 아닙니다.", icon="📘")

projects = add_schedule_columns(ss.projects, base_date)
filtered = apply_filters(projects, selections)

render_resource_cards(load_resources())
st.divider()

st.caption(f"기준일 {base_date:%Y-%m-%d} · 전체 {len(projects)}건 중 {len(filtered)}건 표시 · {SESSION_NOTICE}")
render_summary_cards(filtered)

if filtered.empty:
    # 0건이면 차트·표 계산을 하지 않음
    st.info(EMPTY_MESSAGE)
else:
    render_charts(filtered)
    render_table(filtered)

render_downloads(projects, filtered, base_date, selections)

if not filtered.empty:
    st.divider()
    render_detail(filtered)
