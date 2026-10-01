from __future__ import annotations

import difflib
import io
import re
import unicodedata
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from openpyxl import load_workbook


EXCLUDED_VENDORS = {"비즈브릭스", "레몬트리", "지엠홀딩스"}


@dataclass
class DealSales:
    seller: str
    item: str
    amount: int
    order_count: int
    source_title: str


def clean_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"\s+", " ", text.replace("\xa0", " ").replace("\u200b", "")).strip()


def clean_key(value: Any) -> str:
    return re.sub(r"[^가-힣a-z0-9]", "", clean_text(value).lower())


def read_excel_sheets(file_bytes: bytes, file_name: str) -> dict[str, pd.DataFrame]:
    if Path(file_name).suffix.lower() not in {".xlsx", ".xls"}:
        raise ValueError(f"{file_name}: 엑셀 파일만 사용할 수 있습니다.")
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None, dtype=object)


def find_header(raw: pd.DataFrame, required: Iterable[str], max_rows: int = 30) -> int | None:
    required_keys = [clean_key(value) for value in required]
    for index in range(min(max_rows, len(raw))):
        values = {clean_key(value) for value in raw.iloc[index].tolist() if clean_text(value)}
        if all(key in values for key in required_keys):
            return index
    return None


def dataframe_from_raw(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    columns, seen = [], Counter()
    for number, value in enumerate(raw.iloc[header_row].tolist()):
        base = clean_text(value) or f"빈열_{number}"
        seen[base] += 1
        columns.append(base if seen[base] == 1 else f"{base}.{seen[base] - 1}")
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = columns
    return df.dropna(how="all").reset_index(drop=True)


def find_column(df: pd.DataFrame, names: Iterable[str]) -> str:
    wanted = {clean_key(name) for name in names}
    for column in df.columns:
        if clean_key(column) in wanted:
            return column
    raise ValueError(" / ".join(names) + " 열을 찾지 못했습니다.")


def parse_deal_title(value: Any) -> tuple[str, str] | None:
    text = clean_text(value)
    match = re.match(r"^(.*?)\s*[xX×]\s*(.+)$", text)
    if not match:
        return None
    seller = clean_text(match.group(1))
    item = re.sub(r"(^|\s)무아스(?=\s|$)", " ", clean_text(match.group(2)), flags=re.IGNORECASE)
    item = clean_text(item)
    return (seller, item) if seller and item else None


def parse_detail_file(file_bytes: bytes, file_name: str) -> dict[str, list[tuple[str, str, str]]]:
    sheets = read_excel_sheets(file_bytes, file_name)
    target = next((raw for name, raw in sheets.items() if clean_key(name) == clean_key("상세내역_통합")), None)
    if target is None:
        raise ValueError(f"{file_name}: '상세내역_통합' 탭이 없습니다.")
    header = find_header(target, ["주문번호", "상품명"])
    if header is None:
        raise ValueError(f"{file_name}: 상세내역_통합 탭에서 주문번호·상품명 제목을 찾지 못했습니다.")
    df = dataframe_from_raw(target, header)
    order_col, title_col = find_column(df, ["주문번호"]), find_column(df, ["상품명"])
    result: dict[str, list[tuple[str, str, str]]] = defaultdict(list)
    for _, row in df.iterrows():
        order, parsed = clean_text(row[order_col]), parse_deal_title(row[title_col])
        if not order.upper().startswith("NMO") or parsed is None:
            continue
        triple = (parsed[0], parsed[1], clean_text(row[title_col]))
        if triple not in result[order]:
            result[order].append(triple)
    return dict(result)


def parse_order_file(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    for raw in read_excel_sheets(file_bytes, file_name).values():
        header = find_header(raw, ["★판매처주문번호(수정가능)", "품명", "★결제액(수정가능)"])
        if header is None:
            continue
        df = dataframe_from_raw(raw, header)
        return pd.DataFrame({
            "주문번호": df[find_column(df, ["★판매처주문번호(수정가능)"])].map(clean_text),
            "품명": df[find_column(df, ["품명"])].map(clean_text),
            "결제액": pd.to_numeric(df[find_column(df, ["★결제액(수정가능)"])], errors="coerce").fillna(0),
        })
    raise ValueError(f"{file_name}: 스룩 주문 파일의 주문번호·품명·결제액 제목을 찾지 못했습니다.")


def parse_calendar_file(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    required = ["시작일", "종료일", "진행일", "밴더사", "셀러", "품목", "공구명"]
    for raw in read_excel_sheets(file_bytes, file_name).values():
        header = find_header(raw, required)
        if header is None:
            continue
        df = dataframe_from_raw(raw, header)
        result = pd.DataFrame()
        for name in ["시작일", "종료일", "진행일", "밴더사", "셀러", "셀러 링크", "팔로워", "품번", "품목", "공구명", "차수", "비고"]:
            try:
                result[name] = df[find_column(df, [name])]
            except ValueError:
                result[name] = ""
        result["시작일"] = pd.to_datetime(result["시작일"], errors="coerce")
        result["종료일"] = pd.to_datetime(result["종료일"], errors="coerce")
        for index, row in result.iterrows():
            seller = clean_text(row["셀러"])
            deal_name = clean_text(row["공구명"])
            if not seller:
                bracket = re.match(r"^\[([^\]]+)\]\s*(.*)$", deal_name)
                if bracket:
                    result.at[index, "셀러"] = clean_text(bracket.group(1))
                    result.at[index, "공구명"] = clean_text(bracket.group(2))
        return result.dropna(subset=["시작일"]).reset_index(drop=True)
    raise ValueError(f"{file_name}: 캘린더 최신본 형식의 제목 행을 찾지 못했습니다.")


def tokenize(value: Any) -> set[str]:
    text = clean_text(value).lower()
    text = re.sub(r"무아스|\d+\s*종\s*택\s*1|\([^)]*\)|[&+/,_\-]", " ", text)
    aliases = {"헤어드라이기": "드라이기", "스팀다리미": "다리미", "자동디스펜서": "디스펜서"}
    return {aliases.get(token, token) for token in re.findall(r"[가-힣a-z0-9]+", text)
            if len(token) > 1 and token not in {"프리미엄", "스마트센서", "옵션별", "보기"}}


def item_similarity(left: Any, right: Any) -> float:
    left_text, right_text = clean_key(left), clean_key(right)
    left_tokens, right_tokens = tokenize(left), tokenize(right)
    union = left_tokens | right_tokens
    jaccard = len(left_tokens & right_tokens) / len(union) if union else 0
    return 0.62 * jaccard + 0.38 * difflib.SequenceMatcher(None, left_text, right_text).ratio()


def seller_similarity(left: Any, right: Any) -> float:
    a, b = clean_key(left), clean_key(right)
    if not a or not b:
        return 0
    if a == b:
        return 1
    if (a in b or b in a) and min(len(a), len(b)) >= 2:
        return 0.94
    return difflib.SequenceMatcher(None, a, b).ratio()


def build_deal_sales(order_df: pd.DataFrame, order_deals: dict[str, list[tuple[str, str, str]]]) -> tuple[list[DealSales], pd.DataFrame]:
    amounts: dict[tuple[str, str, str], float] = defaultdict(float)
    orders: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    audit_rows = []
    for order, rows in order_df.groupby("주문번호", sort=False):
        candidates = order_deals.get(clean_text(order), [])
        if not candidates:
            audit_rows.append({"주문번호": order, "스룩 품명": " / ".join(rows["품명"].astype(str)), "결제액": rows["결제액"].sum(), "셀러": "", "공구명": "", "상태": "상세내역_통합에 주문번호 없음"})
            continue
        for _, sales_row in rows.iterrows():
            chosen = max(candidates, key=lambda deal: item_similarity(sales_row["품명"], deal[1]))
            amounts[chosen] += float(sales_row["결제액"])
            orders[chosen].add(clean_text(order))
            audit_rows.append({"주문번호": order, "스룩 품명": sales_row["품명"], "결제액": sales_row["결제액"], "셀러": chosen[0], "공구명": chosen[1], "상태": "주문번호 매칭" if len(candidates) == 1 else "주문번호+품명 매칭"})
    deals = [DealSales(key[0], key[1], int(round(amount)), len(orders[key]), key[2]) for key, amount in amounts.items()]
    return deals, pd.DataFrame(audit_rows)


def detect_sources(files: list[tuple[bytes, str]]) -> tuple[pd.DataFrame, dict[str, list[tuple[str, str, str]]], pd.DataFrame, list[str]]:
    order_frames, recognized = [], []
    order_deals: dict[str, list[tuple[str, str, str]]] = {}
    calendar_frames = []
    for file_bytes, file_name in files:
        try:
            parsed = parse_detail_file(file_bytes, file_name)
            for order, deals in parsed.items():
                order_deals.setdefault(order, [])
                order_deals[order].extend(deal for deal in deals if deal not in order_deals[order])
            recognized.append(f"{file_name}: 상세내역_통합")
            continue
        except Exception:
            pass
        try:
            order_frames.append(parse_order_file(file_bytes, file_name))
            recognized.append(f"{file_name}: 스룩 주문내역")
            continue
        except Exception:
            pass
        try:
            calendar_frames.append(parse_calendar_file(file_bytes, file_name))
            recognized.append(f"{file_name}: 캘린더 최신본")
        except Exception:
            pass
    if not order_frames:
        raise ValueError("왼쪽 파일에서 '(주)스룩_오클릭' 형식의 주문내역을 찾지 못했습니다.")
    if not order_deals:
        raise ValueError("왼쪽 파일에서 '상세내역_통합' 탭이 있는 차액 파일을 찾지 못했습니다.")
    if not calendar_frames:
        raise ValueError("왼쪽 파일에서 시작일·종료일·밴더사·셀러·공구명이 있는 캘린더 최신본을 찾지 못했습니다.")
    return pd.concat(order_frames, ignore_index=True), order_deals, pd.concat(calendar_frames, ignore_index=True), recognized


def locate_template_sheet(workbook):
    required = ["년", "월", "진행일", "밴더사", "셀러", "판매금액"]
    for sheet in workbook.worksheets:
        for row_number in range(1, min(sheet.max_row, 30) + 1):
            values = [clean_text(sheet.cell(row_number, col).value) for col in range(1, sheet.max_column + 1)]
            if all(name in values for name in required) and ("딜명" in values or "품목" in values):
                columns = {name: values.index(name) + 1 for name in required}
                columns["딜명"] = values.index("딜명") + 1 if "딜명" in values else None
                columns["품목"] = values.index("품목") + 1 if "품목" in values else None
                columns["매칭품목"] = columns["딜명"] or columns["품목"]
                for optional in ["셀러 링크", "팔로워", "품번"]:
                    columns[optional] = values.index(optional) + 1 if optional in values else None
                columns["비고"] = values.index("비고") + 1 if "비고" in values else None
                return sheet, row_number, columns
    raise ValueError("오른쪽 파일에서 년·월·밴더사·셀러·딜명(또는 품목)·판매금액 제목을 찾지 못했습니다.")


def row_period(sheet, row_number: int, columns: dict[str, int]) -> tuple[int | None, int | None]:
    year_match = re.search(r"20\d{2}", clean_text(sheet.cell(row_number, columns["년"]).value))
    month_match = re.search(r"1[0-2]|0?[1-9]", clean_text(sheet.cell(row_number, columns["월"]).value))
    return (int(year_match.group()) if year_match else None, int(month_match.group()) if month_match else None)


def copy_row_style(sheet, source_row: int, target_row: int) -> None:
    for col in range(1, sheet.max_column + 1):
        source, target = sheet.cell(source_row, col), sheet.cell(target_row, col)
        if source.has_style:
            target._style = copy(source._style)
        target.number_format, target.alignment = source.number_format, copy(source.alignment)
    sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height


def add_dataframe_sheet(workbook, name: str, dataframe: pd.DataFrame) -> None:
    if name in workbook.sheetnames:
        del workbook[name]
    ws = workbook.create_sheet(name)
    ws.append(list(dataframe.columns))
    for row in dataframe.itertuples(index=False, name=None):
        ws.append(list(row))
    ws.freeze_panes, ws.auto_filter.ref = "A2", ws.dimensions
    for cell in ws[1]:
        cell.font = copy(cell.font); cell.font = cell.font.copy(bold=True, color="FFFFFF")
        cell.fill = copy(cell.fill); cell.fill = cell.fill.copy(fill_type="solid", fgColor="44546A")
    for column in ws.columns:
        ws.column_dimensions[column[0].column_letter].width = min(45, max(12, max(len(clean_text(cell.value)) for cell in column) + 2))


def fill_template(template_bytes: bytes, target_month: int, deals: list[DealSales], audit_df: pd.DataFrame, calendar_df: pd.DataFrame):
    workbook = load_workbook(io.BytesIO(template_bytes))
    sheet, header_row, columns = locate_template_sheet(workbook)
    periods = [row_period(sheet, row, columns) for row in range(header_row + 1, sheet.max_row + 1)]
    years = [year for year, month in periods if month == target_month and year]
    target_year = max(years) if years else 2026
    target_rows = [row for row in range(header_row + 1, sheet.max_row + 1) if row_period(sheet, row, columns) == (target_year, target_month)]
    excluded_keys = {clean_key(vendor) for vendor in EXCLUDED_VENDORS}

    # 최신 캘린더에만 있는 일정을 결과 양식에 먼저 추가한다.
    calendar_month = calendar_df[(calendar_df["시작일"].dt.year == target_year) & (calendar_df["시작일"].dt.month == target_month)].copy()
    style_row = max(target_rows) if target_rows else max(header_row + 1, sheet.max_row)
    insert_row = max(target_rows) + 1 if target_rows else sheet.max_row + 1
    calendar_added_rows: set[int] = set()
    for _, calendar in calendar_month.iterrows():
        cal_seller = clean_text(calendar["셀러"])
        cal_item = clean_text(calendar["공구명"]) or clean_text(calendar["품목"])
        if not cal_seller or not cal_item:
            continue
        exists = False
        for row_number in target_rows:
            row_seller = clean_text(sheet.cell(row_number, columns["셀러"]).value)
            row_item = clean_text(sheet.cell(row_number, columns["매칭품목"]).value)
            if seller_similarity(cal_seller, row_seller) >= 0.9 and item_similarity(cal_item, row_item) >= 0.34:
                exists = True
                break
        if exists:
            continue
        copy_row_style(sheet, style_row, insert_row)
        for col in range(1, sheet.max_column + 1):
            sheet.cell(insert_row, col).value = None
        sheet.cell(insert_row, columns["년"]).value = f"{target_year}년"
        sheet.cell(insert_row, columns["월"]).value = f"{target_month}월"
        sheet.cell(insert_row, columns["진행일"]).value = clean_text(calendar["진행일"])
        sheet.cell(insert_row, columns["밴더사"]).value = clean_text(calendar["밴더사"])
        sheet.cell(insert_row, columns["셀러"]).value = cal_seller
        if columns["셀러 링크"]:
            sheet.cell(insert_row, columns["셀러 링크"]).value = clean_text(calendar["셀러 링크"])
        if columns["팔로워"]:
            sheet.cell(insert_row, columns["팔로워"]).value = clean_text(calendar["팔로워"])
        if columns["품번"]:
            sheet.cell(insert_row, columns["품번"]).value = clean_text(calendar["품번"])
        if columns["품목"]:
            sheet.cell(insert_row, columns["품목"]).value = clean_text(calendar["품목"])
        if columns["딜명"]:
            sheet.cell(insert_row, columns["딜명"]).value = cal_item
        if columns["비고"]:
            sheet.cell(insert_row, columns["비고"]).value = clean_text(calendar["비고"]) or "캘린더 최신본에서 자동 추가"
        target_rows.append(insert_row)
        calendar_added_rows.add(insert_row)
        style_row = insert_row
        insert_row += 1
    historical_vendors: dict[str, str] = {}
    for row in range(header_row + 1, sheet.max_row + 1):
        seller_value = clean_text(sheet.cell(row, columns["셀러"]).value)
        vendor_value = clean_text(sheet.cell(row, columns["밴더사"]).value)
        if seller_value and vendor_value:
            for seller_name in re.split(r"[\r\n,/]+", seller_value):
                if clean_key(seller_name):
                    historical_vendors[clean_key(seller_name)] = vendor_value
    seller_row_counts: Counter[str] = Counter()
    for row_number in target_rows:
        for seller_name in re.split(r"[\r\n,/]+", clean_text(sheet.cell(row_number, columns["셀러"]).value)):
            if clean_key(seller_name):
                seller_row_counts[clean_key(seller_name)] += 1
    used, result_rows = set(), []
    for row_number in target_rows:
        vendor = clean_text(sheet.cell(row_number, columns["밴더사"]).value)
        seller = clean_text(sheet.cell(row_number, columns["셀러"]).value)
        deal_item = clean_text(sheet.cell(row_number, columns["매칭품목"]).value)
        product_item = clean_text(sheet.cell(row_number, columns["품목"]).value) if columns["품목"] else ""
        item = clean_text(f"{deal_item} {product_item}")
        amount_cell = sheet.cell(row_number, columns["판매금액"])
        if clean_key(vendor) in excluded_keys:
            amount_cell.value = None
            status = "지정 밴더 공란 유지"
        elif amount_cell.value not in (None, ""):
            status = "기존 금액 유지"
        else:
            sellers = [part.strip() for part in re.split(r"[\r\n,/]+", seller) if part.strip()]
            candidates = []
            for deal_index, deal in enumerate(deals):
                if deal_index in used:
                    continue
                seller_score = max((seller_similarity(value, deal.seller) for value in sellers), default=0)
                item_score = item_similarity(item, deal.item)
                single_seller_schedule = any(seller_row_counts[clean_key(value)] == 1 for value in sellers)
                if seller_score >= 0.82 and (item_score >= 0.28 or (single_seller_schedule and deal.amount > 0)):
                    candidates.append((seller_score * 0.58 + item_score * 0.42, deal_index, deal))
            if candidates:
                single_seller_schedule = any(seller_row_counts[clean_key(value)] == 1 for value in sellers)
                combined = single_seller_schedule or any(mark in item for mark in ("&", "\n", "2종", "3종")) or len(product_item.splitlines()) > 1
                selected = sorted(candidates, reverse=True)
                if not combined:
                    selected = selected[:1]
                elif not single_seller_schedule:
                    selected = [candidate for candidate in selected if item_similarity(item, candidate[2].item) >= 0.2]
                for _, deal_index, _ in selected:
                    used.add(deal_index)
                amount_cell.value = sum(deal.amount for _, _, deal in selected)
                status = "자동 입력" if len(selected) == 1 else f"자동 입력({len(selected)}건 합산)"
            else:
                status = "매출 미확인"
        if row_number in calendar_added_rows:
            status = f"캘린더 추가 · {status}"
        result_rows.append({"행": row_number, "밴더사": vendor, "셀러": seller, "품목": deal_item, "판매금액": amount_cell.value, "상태": status})

    result_df = pd.DataFrame(result_rows)
    unmatched_deals = pd.DataFrame([
        {"셀러": deal.seller, "공구명": deal.item, "판매금액": deal.amount, "주문수": deal.order_count, "확인사항": "당월 캘린더 일정 없음"}
        for deal_index, deal in enumerate(deals) if deal_index not in used
    ])
    add_dataframe_sheet(workbook, "매칭결과", result_df)
    add_dataframe_sheet(workbook, "주문번호매칭", audit_df)
    add_dataframe_sheet(workbook, "캘린더에없는매출", unmatched_deals)
    output = io.BytesIO(); workbook.save(output)
    return output.getvalue(), result_df, target_year


def process_files(source_files: list[tuple[bytes, str]], template_bytes: bytes, template_name: str, target_month: int):
    if Path(template_name).suffix.lower() != ".xlsx":
        raise ValueError("오른쪽 결과 양식은 서식을 유지하기 위해 XLSX 파일만 사용할 수 있습니다.")
    order_df, order_deals, calendar_df, recognized = detect_sources(source_files)
    deals, audit_df = build_deal_sales(order_df, order_deals)
    if not deals:
        raise ValueError("주문번호로 연결되는 셀러×무아스 매출이 없습니다.")
    excel_bytes, result_df, target_year = fill_template(template_bytes, target_month, deals, audit_df, calendar_df)
    return result_df, audit_df, excel_bytes, deals, recognized, target_year
