import io
import re
from pathlib import Path
from typing import Iterable

import pandas as pd


OUTPUT_COLUMNS = ["년", "월", "진행일", "밴더사", "셀러", "품목", "판매금액"]


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("\xa0", " ").replace("\u200b", " ")
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def clean_column_name(value) -> str:
    return re.sub(r"\s+", "", clean_text(value))


def make_unique_columns(columns: Iterable[object]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []

    for raw_column in columns:
        base = clean_text(raw_column) or "빈열"
        count = seen.get(base, 0)
        seen[base] = count + 1
        result.append(base if count == 0 else f"{base}_{count + 1}")

    return result


def find_column(dataframe: pd.DataFrame, target: str) -> str:
    normalized_target = clean_column_name(target)
    for column in dataframe.columns:
        if clean_column_name(column) == normalized_target:
            return column
    raise ValueError(f"'{target}' 열을 찾지 못했습니다.")


def find_header_row(raw: pd.DataFrame) -> int:
    max_rows = min(len(raw), 60)
    for row_index in range(max_rows):
        values = [clean_column_name(value) for value in raw.iloc[row_index].tolist()]
        if all(required in values for required in ("상품코드", "상품명", "정상금액")):
            return row_index

    raise ValueError(
        "상품코드·상품명·정상금액이 있는 제목 행을 찾지 못했습니다. "
        "스마트스토어 상품별 매출 다운로드 파일인지 확인해주세요."
    )


def decode_csv(file_bytes: bytes) -> str:
    for encoding in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV 파일의 문자 인코딩을 확인하지 못했습니다.")


def load_raw_table(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    suffix = Path(file_name).suffix.lower()

    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=0,
            header=None,
            dtype=object,
        )

    if suffix == ".csv":
        text = decode_csv(file_bytes)
        return pd.read_csv(
            io.StringIO(text),
            header=None,
            dtype=object,
            sep=None,
            engine="python",
        )

    raise ValueError("XLSX, XLS, CSV 파일만 업로드할 수 있습니다.")


def parse_groupbuy_product_name(product_name: object) -> tuple[str, str] | None:
    cleaned = clean_text(product_name)
    cleaned = re.sub(r"\s*옵션별\s*보기\s*$", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.strip()

    match = re.match(
        r"^(?P<seller>.+?)\s*[xX×]\s*무아스\s*(?P<item>.+)$",
        cleaned,
        flags=re.IGNORECASE,
    )
    if not match:
        return None

    seller = clean_text(match.group("seller"))
    item = clean_text(match.group("item"))
    if not seller or not item:
        return None
    return seller, item


def read_vendor_mapping(mapping_path: str | Path | None) -> dict[str, str]:
    if mapping_path is None:
        return {}

    path = Path(mapping_path)
    if not path.exists():
        return {}

    mapping_df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
    seller_col = find_column(mapping_df, "셀러")
    vendor_col = find_column(mapping_df, "밴더사")

    mapping: dict[str, str] = {}
    for _, row in mapping_df.iterrows():
        seller = clean_text(row[seller_col])
        vendor = clean_text(row[vendor_col])
        if seller:
            mapping[seller] = vendor
    return mapping


def parse_sales_file(
    file_bytes: bytes,
    file_name: str,
    vendor_mapping: dict[str, str] | None = None,
    include_zero_sales: bool = False,
) -> pd.DataFrame:
    vendor_mapping = vendor_mapping or {}
    raw = load_raw_table(file_bytes, file_name)
    header_row = find_header_row(raw)

    header = make_unique_columns(raw.iloc[header_row].tolist())
    dataframe = raw.iloc[header_row + 1 :].copy()
    dataframe.columns = header
    dataframe = dataframe.dropna(how="all").reset_index(drop=True)

    code_column = find_column(dataframe, "상품코드")
    name_column = find_column(dataframe, "상품명")
    sales_column = find_column(dataframe, "정상금액")

    dataframe[code_column] = dataframe[code_column].map(clean_text)
    dataframe[name_column] = dataframe[name_column].map(clean_text)

    # 상품코드가 비어 있는 '(부분취소)' 보조 행 등을 제거한다.
    dataframe = dataframe[
        dataframe[code_column].str.match(r"^SMO", case=False, na=False)
    ].copy()

    sales_text = (
        dataframe[sales_column]
        .map(clean_text)
        .str.replace(",", "", regex=False)
        .str.replace("원", "", regex=False)
        .str.replace("₩", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.replace(" ", "", regex=False)
    )
    dataframe["판매금액"] = (
        pd.to_numeric(sales_text, errors="coerce").fillna(0).round().astype("int64")
    )

    parsed_rows: list[dict[str, object]] = []
    for _, row in dataframe.iterrows():
        parsed = parse_groupbuy_product_name(row[name_column])
        if parsed is None:
            continue

        seller, item = parsed
        sales = int(row["판매금액"])
        if not include_zero_sales and sales == 0:
            continue

        parsed_rows.append(
            {
                "년": "",
                "월": "",
                "진행일": "",
                "밴더사": vendor_mapping.get(seller, ""),
                "셀러": seller,
                "품목": item,
                "판매금액": sales,
            }
        )

    return pd.DataFrame(parsed_rows, columns=OUTPUT_COLUMNS)


def normalize_output_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    result = dataframe.copy()
    for column in OUTPUT_COLUMNS:
        if column not in result.columns:
            result[column] = "" if column != "판매금액" else 0

    result = result[OUTPUT_COLUMNS]
    for column in ["년", "월", "진행일", "밴더사", "셀러", "품목"]:
        result[column] = result[column].map(clean_text)

    result["판매금액"] = (
        pd.to_numeric(result["판매금액"], errors="coerce").fillna(0).round().astype("int64")
    )
    return result
