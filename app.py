import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st


APP_VERSION = "1.0.0"
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_MAPPING_PATH = BASE_DIR / "vendor_mapping.csv"


st.set_page_config(
    page_title="공구 매출 정리 도구",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# -----------------------------
# 화면 스타일
# -----------------------------
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 2.2rem;
            padding-bottom: 4rem;
        }
        [data-testid="stSidebar"] {
            min-width: 280px;
            max-width: 280px;
        }
        .main-title {
            font-size: 2.1rem;
            font-weight: 800;
            margin-bottom: 0.2rem;
        }
        .sub-text {
            color: #667085;
            margin-bottom: 1.5rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# 공통 함수
# -----------------------------
def normalize_text(value) -> str:
    """공백, 줄바꿈, NBSP 등 엑셀에서 생기는 이상한 문자를 정리한다."""
    if pd.isna(value):
        return ""

    text = str(value)
    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_column_name(value) -> str:
    return re.sub(r"\s+", "", normalize_text(value))


def make_unique_columns(columns):
    """중복된 열 이름이 있어도 안전하게 읽도록 열 이름을 고유하게 만든다."""
    seen = {}
    result = []

    for column in columns:
        base_name = normalize_text(column) or "빈열"
        count = seen.get(base_name, 0)
        seen[base_name] = count + 1

        if count == 0:
            result.append(base_name)
        else:
            result.append(f"{base_name}_{count + 1}")

    return result


def find_column(dataframe: pd.DataFrame, target: str) -> str:
    normalized_target = normalize_column_name(target)

    for column in dataframe.columns:
        if normalize_column_name(column) == normalized_target:
            return column

    raise ValueError(f"'{target}' 열을 찾지 못했습니다.")


def find_header_row(raw: pd.DataFrame) -> int:
    """파일 상단에 빈 행이나 부가 설명 행이 있어도 실제 헤더 행을 찾는다."""
    max_rows = min(len(raw), 60)

    for row_index in range(max_rows):
        values = [normalize_column_name(value) for value in raw.iloc[row_index].tolist()]

        has_code = "상품코드" in values
        has_name = "상품명" in values
        has_sales = "정상금액" in values

        if has_code and has_name and has_sales:
            return row_index

    raise ValueError(
        "상품코드·상품명·정상금액이 있는 헤더 행을 찾지 못했습니다. "
        "스마트스토어 상품별 매출 파일인지 확인해주세요."
    )


def decode_csv(file_bytes: bytes) -> str:
    for encoding in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue

    raise ValueError("CSV 파일의 문자 인코딩을 확인하지 못했습니다.")


def get_sheet_names(file_bytes: bytes, file_name: str):
    suffix = Path(file_name).suffix.lower()

    if suffix in {".xlsx", ".xls"}:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        return excel_file.sheet_names

    return ["CSV"]


def load_raw_table(file_bytes: bytes, file_name: str, sheet_name=None) -> pd.DataFrame:
    suffix = Path(file_name).suffix.lower()

    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet_name or 0,
            header=None,
            dtype=object,
        )

    if suffix == ".csv":
        csv_text = decode_csv(file_bytes)
        return pd.read_csv(
            io.StringIO(csv_text),
            header=None,
            dtype=object,
            sep=None,
            engine="python",
        )

    raise ValueError("XLSX, XLS, CSV 파일만 업로드할 수 있습니다.")


def parse_sales_file(file_bytes: bytes, file_name: str, sheet_name=None) -> pd.DataFrame:
    raw = load_raw_table(file_bytes, file_name, sheet_name)
    header_row = find_header_row(raw)

    header = make_unique_columns(raw.iloc[header_row].tolist())
    dataframe = raw.iloc[header_row + 1 :].copy()
    dataframe.columns = header
    dataframe = dataframe.dropna(how="all").reset_index(drop=True)

    code_column = find_column(dataframe, "상품코드")
    name_column = find_column(dataframe, "상품명")
    sales_column = find_column(dataframe, "정상금액")

    dataframe[code_column] = dataframe[code_column].map(normalize_text)
    dataframe[name_column] = dataframe[name_column].map(normalize_text)

    # '(부분취소)'처럼 상품코드가 없는 보조 행을 제외한다.
    dataframe = dataframe[
        dataframe[code_column].str.match(r"^SMO", case=False, na=False)
    ].copy()

    sales_text = (
        dataframe[sales_column]
        .map(normalize_text)
        .str.replace(",", "", regex=False)
        .str.replace("원", "", regex=False)
        .str.replace("₩", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    dataframe["판매금액"] = pd.to_numeric(sales_text, errors="coerce").fillna(0).astype("int64")
    dataframe["원본상품명"] = dataframe[name_column]
    dataframe["상품코드"] = dataframe[code_column]

    parsed = dataframe.apply(
        lambda row: parse_product_name(row["원본상품명"]),
        axis=1,
        result_type="expand",
    )
    parsed.columns = ["셀러", "품목", "공구상품여부"]

    dataframe = pd.concat([dataframe, parsed], axis=1)
    return dataframe[
        ["상품코드", "원본상품명", "셀러", "품목", "판매금액", "공구상품여부"]
    ].reset_index(drop=True)


def parse_product_name(product_name: str):
    """
    예:
    김희경 x 무아스 2 in 1 스윙 핸디 스팀다리미옵션별 보기
    -> 김희경 / 2 in 1 스윙 핸디 스팀다리미
    """
    cleaned = normalize_text(product_name)
    cleaned = re.sub(r"\s*옵션별\s*보기\s*$", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.strip()

    pattern = re.compile(
        r"^(?P<seller>.+?)\s*[xX×]\s*무아스\s*(?P<item>.+)$",
        flags=re.IGNORECASE,
    )
    match = pattern.match(cleaned)

    if match:
        seller = normalize_text(match.group("seller"))
        item = normalize_text(match.group("item"))
        return seller, item, True

    # '무아스 ○○' 형태는 자사상품으로 분류한다.
    if re.match(r"^\s*(\[.*?\]\s*)*무아스\s+", cleaned):
        item = re.sub(
            r"^\s*(\[.*?\]\s*)*무아스\s+",
            "",
            cleaned,
            flags=re.IGNORECASE,
        )
        return "자사몰", normalize_text(item), False

    return "확인 필요", cleaned, False


def read_mapping_dataframe(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    suffix = Path(file_name).suffix.lower()

    if suffix == ".csv":
        text = decode_csv(file_bytes)
        mapping = pd.read_csv(io.StringIO(text), dtype=str)
    elif suffix in {".xlsx", ".xls"}:
        mapping = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    else:
        raise ValueError("벤더 매핑 파일은 XLSX, XLS, CSV만 가능합니다.")

    mapping.columns = [normalize_text(column) for column in mapping.columns]

    seller_column = find_column(mapping, "셀러")
    vendor_column = find_column(mapping, "벤더사")

    mapping = mapping[[seller_column, vendor_column]].copy()
    mapping.columns = ["셀러", "벤더사"]
    mapping["셀러"] = mapping["셀러"].map(normalize_text)
    mapping["벤더사"] = mapping["벤더사"].map(normalize_text)
    mapping = mapping[mapping["셀러"] != ""]
    mapping = mapping.drop_duplicates("셀러", keep="last")

    return mapping.reset_index(drop=True)


def load_default_mapping() -> pd.DataFrame:
    if not DEFAULT_MAPPING_PATH.exists():
        return pd.DataFrame(columns=["셀러", "벤더사"])

    try:
        return pd.read_csv(DEFAULT_MAPPING_PATH, dtype=str).fillna("")
    except Exception:
        return pd.DataFrame(columns=["셀러", "벤더사"])


def combine_mappings(default_mapping, uploaded_mapping):
    mapping = pd.concat(
        [default_mapping, uploaded_mapping],
        ignore_index=True,
    )
    mapping["셀러"] = mapping["셀러"].map(normalize_text)
    mapping["벤더사"] = mapping["벤더사"].map(normalize_text)
    mapping = mapping[mapping["셀러"] != ""]
    return mapping.drop_duplicates("셀러", keep="last").reset_index(drop=True)


def create_excel_download(result: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result.to_excel(
            writer,
            index=False,
            sheet_name="공구 매출 정리",
        )

        workbook = writer.book
        worksheet = writer.sheets["공구 매출 정리"]

        header_format = workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#344054",
                "align": "center",
                "valign": "vcenter",
                "border": 1,
            }
        )
        text_format = workbook.add_format(
            {
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

        for column_index, column_name in enumerate(result.columns):
            worksheet.write(0, column_index, column_name, header_format)

        if len(result) > 0:
            worksheet.set_column("A:A", 18, text_format)
            worksheet.set_column("B:B", 18, text_format)
            worksheet.set_column("C:C", 48, text_format)
            worksheet.set_column("D:D", 16, money_format)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(result), 1), len(result.columns) - 1)
        worksheet.set_row(0, 24)

    output.seek(0)
    return output.getvalue()


def create_mapping_csv(mapping: pd.DataFrame) -> bytes:
    return mapping.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# -----------------------------
# 사이드바
# -----------------------------
with st.sidebar:
    st.title("🛠️ 업무 자동화 도구")

    st.radio(
        "메뉴 선택",
        ["📊 공구 매출 정리"],
        index=0,
    )

    st.divider()
    st.caption(f"v{APP_VERSION} · 정상금액 기준")


# -----------------------------
# 메인 화면
# -----------------------------
st.markdown(
    '<div class="main-title">📊 공구 매출 정리 도구</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-text">'
    "상품별 매출 파일을 올리면 벤더사 · 셀러 · 품목 · 판매금액만 자동으로 정리합니다."
    "</div>",
    unsafe_allow_html=True,
)

uploaded_sales_file = st.file_uploader(
    "상품별 매출 파일 업로드",
    type=["xlsx", "xls", "csv"],
    help="스마트스토어 상품별 매출 파일을 그대로 올려주세요.",
)

with st.expander("⚙️ 벤더사 매핑 설정", expanded=False):
    st.write(
        "원본 매출 파일에는 벤더사가 없어서, 셀러별 벤더사를 등록해야 합니다. "
        "기본 매핑표에 없는 셀러만 이 화면에서 입력하면 됩니다."
    )

    uploaded_mapping_file = st.file_uploader(
        "추가 벤더 매핑 파일 업로드",
        type=["xlsx", "xls", "csv"],
        key="mapping_upload",
        help="열 이름은 반드시 '셀러', '벤더사'여야 합니다.",
    )

    default_mapping = load_default_mapping()
    uploaded_mapping = pd.DataFrame(columns=["셀러", "벤더사"])

    if uploaded_mapping_file is not None:
        try:
            uploaded_mapping = read_mapping_dataframe(
                uploaded_mapping_file.getvalue(),
                uploaded_mapping_file.name,
            )
            st.success(f"추가 매핑 {len(uploaded_mapping):,}건을 불러왔습니다.")
        except Exception as error:
            st.error(str(error))

    base_mapping = combine_mappings(default_mapping, uploaded_mapping)

    st.download_button(
        "📥 기본 벤더 매핑표 다운로드",
        data=create_mapping_csv(base_mapping),
        file_name="vendor_mapping.csv",
        mime="text/csv",
    )


if uploaded_sales_file is None:
    st.info(
        "위의 업로드 영역에 XLSX, XLS 또는 CSV 파일을 넣어주세요. "
        "파일 상단에 빈 행이 있어도 자동으로 헤더를 찾아냅니다."
    )
    st.stop()


try:
    file_bytes = uploaded_sales_file.getvalue()
    sheet_names = get_sheet_names(file_bytes, uploaded_sales_file.name)

    selected_sheet = sheet_names[0]

    if len(sheet_names) > 1:
        selected_sheet = st.selectbox(
            "처리할 시트 선택",
            options=sheet_names,
        )

    parsed_data = parse_sales_file(
        file_bytes,
        uploaded_sales_file.name,
        selected_sheet if selected_sheet != "CSV" else None,
    )

except Exception as error:
    st.error(f"파일을 처리하지 못했습니다.\n\n{error}")
    st.stop()


col_filter1, col_filter2 = st.columns(2)

with col_filter1:
    only_groupbuy = st.checkbox(
        "공구 상품만 표시",
        value=True,
        help="'셀러 x 무아스 상품명' 형식의 상품만 남깁니다.",
    )

with col_filter2:
    exclude_zero = st.checkbox(
        "판매금액 0원 제외",
        value=True,
    )


working = parsed_data.copy()

if only_groupbuy:
    working = working[working["공구상품여부"]].copy()

if exclude_zero:
    working = working[working["판매금액"] != 0].copy()


# 현재 파일에 등장한 셀러를 기준으로 매핑 편집 표 구성
seller_list = sorted(
    seller
    for seller in working["셀러"].dropna().unique().tolist()
    if seller not in {"", "자사몰", "확인 필요"}
)

mapping_dict = dict(zip(base_mapping["셀러"], base_mapping["벤더사"]))

mapping_editor_source = pd.DataFrame(
    {
        "셀러": seller_list,
        "벤더사": [mapping_dict.get(seller, "") for seller in seller_list],
    }
)

st.subheader("1. 셀러별 벤더사 확인")

edited_mapping = st.data_editor(
    mapping_editor_source,
    use_container_width=True,
    hide_index=True,
    disabled=["셀러"],
    column_config={
        "셀러": st.column_config.TextColumn("셀러"),
        "벤더사": st.column_config.TextColumn(
            "벤더사",
            help="비어 있는 셀러만 벤더사를 입력해주세요.",
        ),
    },
    key=f"vendor_editor_{uploaded_sales_file.name}_{selected_sheet}",
)

current_mapping = dict(
    zip(
        edited_mapping["셀러"].map(normalize_text),
        edited_mapping["벤더사"].map(normalize_text),
    )
)

working["벤더사"] = working["셀러"].map(current_mapping).fillna("")
working.loc[working["벤더사"] == "", "벤더사"] = "미등록"

unmapped_sellers = sorted(
    working.loc[working["벤더사"] == "미등록", "셀러"].unique().tolist()
)

if unmapped_sellers:
    st.warning(
        "벤더사가 등록되지 않은 셀러가 있습니다: "
        + ", ".join(unmapped_sellers)
    )


st.subheader("2. 정리 결과")

aggregate_rows = st.checkbox(
    "같은 벤더사·셀러·품목은 한 줄로 합산",
    value=True,
)

result = working[["벤더사", "셀러", "품목", "판매금액"]].copy()

if aggregate_rows:
    result = (
        result.groupby(
            ["벤더사", "셀러", "품목"],
            as_index=False,
            dropna=False,
        )["판매금액"]
        .sum()
    )

result = result.sort_values(
    ["벤더사", "셀러", "판매금액"],
    ascending=[True, True, False],
).reset_index(drop=True)

metric1, metric2, metric3 = st.columns(3)
metric1.metric("정리 행 수", f"{len(result):,}건")
metric2.metric("판매금액 합계", f"{result['판매금액'].sum():,.0f}원")
metric3.metric("미등록 셀러", f"{len(unmapped_sellers):,}명")

st.dataframe(
    result,
    use_container_width=True,
    hide_index=True,
    column_config={
        "벤더사": st.column_config.TextColumn("벤더사"),
        "셀러": st.column_config.TextColumn("셀러"),
        "품목": st.column_config.TextColumn("품목", width="large"),
        "판매금액": st.column_config.NumberColumn(
            "판매금액",
            format="%,d원",
        ),
    },
)

download_col1, download_col2 = st.columns(2)

with download_col1:
    st.download_button(
        "📥 정리된 엑셀 다운로드",
        data=create_excel_download(result),
        file_name="공구_매출_정리.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with download_col2:
    saved_mapping = edited_mapping.copy()
    st.download_button(
        "📥 현재 벤더 매핑표 다운로드",
        data=create_mapping_csv(saved_mapping),
        file_name="vendor_mapping_수정본.csv",
        mime="text/csv",
        use_container_width=True,
    )


with st.expander("파싱되지 않은 상품 확인"):
    failed_rows = parsed_data[~parsed_data["공구상품여부"]][
        ["상품코드", "원본상품명", "셀러", "품목", "판매금액"]
    ].copy()

    if failed_rows.empty:
        st.success("모든 상품명이 정상적으로 파싱됐습니다.")
    else:
        st.caption(
            "'셀러 x 무아스 상품명' 형식이 아닌 자사상품·개인결제창 등이 표시됩니다."
        )
        st.dataframe(
            failed_rows,
            use_container_width=True,
            hide_index=True,
        )
