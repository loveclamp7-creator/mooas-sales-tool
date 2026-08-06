import hashlib
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from core import OUTPUT_COLUMNS, normalize_output_dataframe, parse_sales_file, read_vendor_mapping


APP_VERSION = "2.0.0"
BASE_DIR = Path(__file__).resolve().parent
VENDOR_MAPPING_PATH = BASE_DIR / "vendor_mapping.csv"


st.set_page_config(
    page_title="인플루언서 공동구매 매출 정리",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .block-container {padding-top: 2rem; padding-bottom: 4rem;}
        [data-testid="stSidebar"] {min-width: 270px; max-width: 270px;}
        .app-title {font-size: 2rem; font-weight: 800; margin-bottom: .3rem;}
        .app-subtitle {color: #667085; margin-bottom: 1.4rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def dataframe_to_excel(dataframe: pd.DataFrame) -> bytes:
    result = normalize_output_dataframe(dataframe)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result.to_excel(writer, index=False, sheet_name="Sheet1")

        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        header_format = workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#44546A",
                "align": "center",
                "valign": "vcenter",
                "border": 1,
            }
        )
        text_format = workbook.add_format(
            {
                "align": "left",
                "valign": "vcenter",
                "border": 1,
            }
        )
        money_format = workbook.add_format(
            {
                "num_format": "#,##0",
                "align": "right",
                "valign": "vcenter",
                "border": 1,
            }
        )

        for column_index, column_name in enumerate(OUTPUT_COLUMNS):
            worksheet.write(0, column_index, column_name, header_format)

        worksheet.set_column("A:A", 11, text_format)
        worksheet.set_column("B:B", 9, text_format)
        worksheet.set_column("C:C", 27, text_format)
        worksheet.set_column("D:D", 18, text_format)
        worksheet.set_column("E:E", 18, text_format)
        worksheet.set_column("F:F", 48, text_format)
        worksheet.set_column("G:G", 16, money_format)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(result), 1), len(OUTPUT_COLUMNS) - 1)
        worksheet.set_row(0, 25)

    output.seek(0)
    return output.getvalue()


with st.sidebar:
    st.title("🛠️ 업무 자동화 도구")
    st.radio("메뉴 선택", ["📊 공동구매 매출 정리"], index=0)
    st.divider()
    st.caption(f"v{APP_VERSION} · 정상금액 기준")

st.markdown('<div class="app-title">📊 인플루언서 공동구매 매출 정리</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-subtitle">상품별 매출 파일을 올리면 필요한 7개 항목만 정리해줍니다.</div>',
    unsafe_allow_html=True,
)

st.info(
    "자동 입력: 셀러 · 품목 · 판매금액(정상금액) · 등록된 밴더사  |  "
    "직접 입력: 년 · 월 · 진행일 · 미등록 밴더사"
)

include_zero_sales = st.checkbox("판매금액 0원인 공구도 포함", value=False)

uploaded_file = st.file_uploader(
    "상품별 매출 파일 업로드",
    type=["xlsx", "xls", "csv"],
    help="첨부해주신 '다운로드파일.xlsx' 같은 파일을 그대로 올리면 됩니다.",
)

if uploaded_file is None:
    st.caption("파일을 올리면 아래에 바로 정리 표가 나타납니다.")
    st.stop()

file_bytes = uploaded_file.getvalue()
file_signature = hashlib.sha256(file_bytes + str(include_zero_sales).encode()).hexdigest()

try:
    vendor_mapping = read_vendor_mapping(VENDOR_MAPPING_PATH)
    parsed = parse_sales_file(
        file_bytes=file_bytes,
        file_name=uploaded_file.name,
        vendor_mapping=vendor_mapping,
        include_zero_sales=include_zero_sales,
    )
except Exception as error:
    st.error(f"파일을 처리하지 못했습니다.\n\n{error}")
    st.stop()

if parsed.empty:
    st.warning("'셀러 x 무아스 상품명' 형태의 공구 상품을 찾지 못했습니다.")
    st.stop()

session_key = "editable_sales_data"
signature_key = "editable_sales_signature"

if st.session_state.get(signature_key) != file_signature:
    st.session_state[session_key] = parsed
    st.session_state[signature_key] = file_signature

st.subheader("정리 결과")
st.caption(
    "표 안의 빈칸은 직접 입력하면 됩니다. 엑셀처럼 셀을 복사·붙여넣기 할 수 있고, "
    "수정한 내용 그대로 다운로드됩니다."
)

edited = st.data_editor(
    st.session_state[session_key],
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic",
    column_order=OUTPUT_COLUMNS,
    column_config={
        "년": st.column_config.TextColumn("년", width="small", help="예: 2026년"),
        "월": st.column_config.TextColumn("월", width="small", help="예: 6월"),
        "진행일": st.column_config.TextColumn(
            "진행일", width="medium", help="예: 2026-06-01 ~ 2026-06-07"
        ),
        "밴더사": st.column_config.TextColumn("밴더사", width="medium"),
        "셀러": st.column_config.TextColumn("셀러", width="medium"),
        "품목": st.column_config.TextColumn("품목", width="large"),
        "판매금액": st.column_config.NumberColumn(
            "판매금액", width="medium", format="%,d"
        ),
    },
    key=f"sales_editor_{file_signature}",
)

edited = normalize_output_dataframe(edited)
st.session_state[session_key] = edited

unmapped = sorted(
    seller
    for seller in edited.loc[edited["밴더사"] == "", "셀러"].dropna().unique().tolist()
    if seller
)

metric1, metric2, metric3 = st.columns(3)
metric1.metric("정리 행 수", f"{len(edited):,}건")
metric2.metric("판매금액 합계", f"{edited['판매금액'].sum():,.0f}원")
metric3.metric("밴더사 빈칸", f"{len(unmapped):,}명")

if unmapped:
    st.warning("밴더사를 직접 입력해야 하는 셀러: " + ", ".join(unmapped))

st.download_button(
    "📥 인플루언서 공동구매 매출현황 엑셀 다운로드",
    data=dataframe_to_excel(edited),
    file_name="인플루언서_공동구매_품목및매출현황_정리본.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

with st.expander("어떤 행이 제외되나요?"):
    st.write(
        "상품명이 '셀러명 x 무아스 상품명' 형식이 아닌 자사몰 일반상품, "
        "벤더사 구입건, 비밀특가 상품 등은 자동으로 제외합니다. "
        "또한 상품코드가 없는 '(부분취소)' 보조 행도 제외합니다."
    )
