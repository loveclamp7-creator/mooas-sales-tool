from __future__ import annotations

import difflib
import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd


ENTRY_REQUIRED = ["년", "월", "진행일", "밴더사", "셀러", "셀러 링크", "팔로워", "품목", "판매금액", "비고"]
SALES_REQUIRED = ["상품코드", "상품명", "상품등록일", "최근주문일", "정상금액"]


@dataclass
class SalesRow:
    source_index: int
    code: str
    seller: str
    item: str
    registered: datetime | None
    recent: datetime | None
    amount: int
    original_name: str


def clean_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).replace("\xa0", " ").replace("\u200b", "")
    text = re.sub(r"[\t ]+", " ", text)
    return text.strip()


def clean_column(value: Any) -> str:
    return re.sub(r"\s+", "", clean_text(value))


def find_header_row(raw: pd.DataFrame, required: list[str], max_rows: int = 60) -> int:
    normalized_required = {clean_column(value) for value in required}
    for row_index in range(min(len(raw), max_rows)):
        row_values = {clean_column(value) for value in raw.iloc[row_index].tolist()}
        if normalized_required.issubset(row_values):
            return row_index
    raise ValueError("필수 제목 행을 찾지 못했습니다: " + ", ".join(required))


def read_first_sheet(file_bytes: bytes, file_name: str, header: int | None = 0) -> pd.DataFrame:
    suffix = Path(file_name).suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=header, dtype=object)
    if suffix == ".csv":
        for encoding in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
            try:
                text = file_bytes.decode(encoding)
                return pd.read_csv(io.StringIO(text), header=header, dtype=object, sep=None, engine="python")
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV 문자 인코딩을 확인하지 못했습니다.")
    raise ValueError("XLSX, XLS, CSV 파일만 사용할 수 있습니다.")


def find_column(df: pd.DataFrame, target: str) -> str:
    normalized = clean_column(target)
    for column in df.columns:
        if clean_column(column) == normalized:
            return column
    raise ValueError(f"'{target}' 열을 찾지 못했습니다.")


def parse_product_name(product_name: Any) -> tuple[str, str]:
    text = clean_text(product_name)
    text = re.sub(r"\s*옵션별\s*보기\s*$", "", text, flags=re.IGNORECASE)
    match = re.match(r"^(.*?)\s*[xX×]\s*(.+)$", text)
    if match:
        seller = clean_text(match.group(1))
        item = clean_text(match.group(2))
        item = re.sub(r"^무아스\s*", "", item, flags=re.IGNORECASE)
        return seller, item
    item = re.sub(r"^(?:\[[^\]]+\]\s*)*무아스\s*", "", text, flags=re.IGNORECASE)
    return "", clean_text(item)


def normalize_seller(value: Any) -> str:
    text = unicodedata.normalize("NFKC", clean_text(value)).lower()
    return re.sub(r"[^가-힣a-z0-9]", "", text)


def split_sellers(value: Any) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    original = str(value).replace("\xa0", " ")
    parts = [part.strip() for part in re.split(r"[\r\n]+", original) if part.strip()]
    return parts or [text]


def parse_schedule(value: Any) -> tuple[datetime | None, datetime | None]:
    dates = []
    for date_text in re.findall(r"20\d{2}-\d{2}-\d{2}", clean_text(value)):
        try:
            dates.append(datetime.strptime(date_text, "%Y-%m-%d"))
        except ValueError:
            pass
    if not dates:
        return None, None
    return min(dates), max(dates)


def product_family(value: Any) -> set[str]:
    text = unicodedata.normalize("NFKC", clean_text(value)).lower()
    families: set[str] = set()
    if "핸디팬" in text or "휴대용 선풍기" in text:
        families.add("휴대용선풍기")
    if "탁상용" in text and "선풍기" in text:
        families.add("탁상용선풍기")
    if ("wifi" in text or "wi-fi" in text) and "선풍기" in text:
        families.add("wifi선풍기")
    if "선풍기" in text and not families:
        families.add("선풍기")
    if "스팀" in text and "다리미" in text:
        families.add("스팀다리미")
    elif "다리미" in text:
        families.add("다리미")
    if "에어롤" in text:
        families.add("에어롤")
    if "라이트젯" in text:
        families.add("라이트젯")
    if "오브제" in text and "디스펜서" in text:
        families.add("오브제디스펜서")
    elif "메탈" in text and "디스펜서" in text:
        families.add("메탈디스펜서")
    elif "디스펜서" in text:
        families.add("디스펜서")
    if "육각" in text and "타이머" in text:
        families.add("육각타이머")
    elif "타이머" in text:
        families.add("타이머")
    if "슈퍼와이드" in text and "스탠드" in text:
        families.add("슈퍼와이드스탠드")
    elif "스탠드" in text:
        families.add("스탠드")
    for word in ("거울", "클리너", "스피커", "에어프라이어", "스텝퍼", "제습제", "멀티탭"):
        if word in text:
            families.add(word)
    return families


def tokenize_item(value: Any) -> tuple[list[str], str]:
    text = unicodedata.normalize("NFKC", clean_text(value)).lower()
    replacements = {
        "wi-fi": "wifi",
        "wi fi": "wifi",
        "스팀 다리미": "스팀다리미",
        "헤어 드라이기": "드라이기",
        "헤어드라이기": "드라이기",
        "자동 디스펜서": "디스펜서",
        "자동디스펜서": "디스펜서",
    }
    for before, after in replacements.items():
        text = text.replace(before, after)
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"(?:옵션별\s*보기|무아스|공동구매\s*\d*차)", " ", text)
    text = re.sub(r"\b(?:2\s*in\s*1|4\s*in\s*1|5\s*in\s*1|2\s*way)\b", " ", text)
    text = re.sub(r"\d+\s*종\s*택\s*1", " ", text)
    text = re.sub(r"\d+\s*종", " ", text)
    text = re.sub(r"[&+/,_\-]+", " ", text)
    text = re.sub(r"[^가-힣a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    ignored = {
        "스팀", "핸디", "냉각", "bldc", "프리미엄", "스마트센서", "아이케어링",
        "브러시", "길이조절", "스퀘어", "젤", "폼", "퓨어", "라인", "워시",
        "멀티", "3d", "회전", "접이식", "파우치", "포함", "프로",
    }
    aliases = {"스팀다리미": "다리미"}
    tokens = []
    for token in re.findall(r"[가-힣a-z0-9]+", text):
        if token in ignored:
            continue
        token = aliases.get(token, token)
        if len(token) > 1:
            tokens.append(token)
    return tokens, text


def item_similarity(left: Any, right: Any) -> float:
    left_tokens, left_text = tokenize_item(left)
    right_tokens, right_text = tokenize_item(right)
    left_set, right_set = set(left_tokens), set(right_tokens)
    union = left_set | right_set
    jac = len(left_set & right_set) / len(union) if union else 0.0
    seq = difflib.SequenceMatcher(None, left_text, right_text).ratio()
    left_family, right_family = product_family(left), product_family(right)
    family_union = left_family | right_family
    family_score = len(left_family & right_family) / max(1, len(family_union)) if family_union else 0.0
    if left_family & right_family:
        score = 0.5 * family_score + 0.25 * jac + 0.25 * seq
    else:
        score = 0.45 * jac + 0.55 * seq
    return min(score, 1.0)


def date_eligible(start: datetime | None, end: datetime | None, sale: SalesRow) -> bool:
    if not start or not end:
        return True
    if sale.registered and sale.registered.date() > (end + timedelta(days=3)).date():
        return False
    if sale.recent and sale.recent.date() < (start - timedelta(days=2)).date():
        return False
    return True


def date_score(start: datetime | None, sale: SalesRow) -> float:
    if not start or not sale.registered:
        return 0.0
    days = (sale.registered.date() - start.date()).days
    if days <= 0:
        return max(0.0, 20.0 - abs(days) * 0.5)
    return max(-20.0, 10.0 - days * 3.0)


def parse_sales_file(file_bytes: bytes, file_name: str) -> list[SalesRow]:
    raw = read_first_sheet(file_bytes, file_name, header=None)
    header_row = find_header_row(raw, ["상품코드", "상품명", "정상금액"])
    headers = [clean_text(value) or f"빈열_{index}" for index, value in enumerate(raw.iloc[header_row].tolist())]
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all")

    code_col = find_column(df, "상품코드")
    name_col = find_column(df, "상품명")
    registered_col = find_column(df, "상품등록일")
    recent_col = find_column(df, "최근주문일")
    amount_col = find_column(df, "정상금액")

    rows: list[SalesRow] = []
    for source_index, (_, row) in enumerate(df.iterrows()):
        code = clean_text(row[code_col])
        if not re.match(r"^SMO", code, flags=re.IGNORECASE):
            continue
        seller, item = parse_product_name(row[name_col])
        amount_text = clean_text(row[amount_col]).replace(",", "").replace("원", "").replace("₩", "")
        amount = int(round(pd.to_numeric(amount_text, errors="coerce") or 0))
        registered = pd.to_datetime(row[registered_col], errors="coerce")
        recent = pd.to_datetime(row[recent_col], errors="coerce")
        rows.append(
            SalesRow(
                source_index=source_index,
                code=code,
                seller=seller,
                item=item,
                registered=None if pd.isna(registered) else registered.to_pydatetime(),
                recent=None if pd.isna(recent) else recent.to_pydatetime(),
                amount=amount,
                original_name=clean_text(row[name_col]),
            )
        )
    return rows


def parse_entry_file(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    raw = read_first_sheet(file_bytes, file_name, header=0)
    renamed = {}
    for target in ENTRY_REQUIRED:
        renamed[find_column(raw, target)] = target
    df = raw.rename(columns=renamed).copy()
    for column in ENTRY_REQUIRED:
        if column not in df.columns:
            df[column] = ""
    return df[ENTRY_REQUIRED]


def merge_sales(entry_df: pd.DataFrame, sales_rows: list[SalesRow]):
    used: set[int] = set()
    result_rows: list[dict[str, Any]] = []
    match_log: list[dict[str, Any]] = []

    entry_seller_norms = {
        normalize_seller(seller)
        for value in entry_df["셀러"].tolist()
        for seller in split_sellers(value)
        if seller
    }

    for original_number, (_, entry) in enumerate(entry_df.iterrows(), start=2):
        base = {column: entry.get(column, "") for column in ENTRY_REQUIRED}
        sellers = split_sellers(base["셀러"])
        start, end = parse_schedule(base["진행일"])
        selected: list[tuple[int, SalesRow, float]] = []

        if len(sellers) == 1 and sellers[0] and clean_text(base["품목"]):
            seller_norm = normalize_seller(sellers[0])
            candidates = []
            for index, sale in enumerate(sales_rows):
                if index in used or normalize_seller(sale.seller) != seller_norm:
                    continue
                if not date_eligible(start, end, sale):
                    continue
                similarity = item_similarity(base["품목"], sale.item)
                if similarity < 0.45:
                    continue
                candidates.append((index, sale, similarity, date_score(start, sale)))

            combined = any(symbol in clean_text(base["품목"]) for symbol in ("&", "+", "\n")) or (
                "핸디팬" in clean_text(base["품목"]) and "탁상용" in clean_text(base["품목"])
            )

            if combined:
                families: list[set[str]] = []
                for candidate in sorted(candidates, key=lambda x: (x[2], x[1].amount != 0, x[3]), reverse=True):
                    family = product_family(candidate[1].item)
                    if candidate[2] < 0.45 or any(family == existing for existing in families):
                        continue
                    selected.append((candidate[0], candidate[1], candidate[2]))
                    families.append(family)
            elif [candidate for candidate in candidates if candidate[2] >= 0.5]:
                candidates = [candidate for candidate in candidates if candidate[2] >= 0.5]
                best = max(candidates, key=lambda x: (x[2], x[1].amount != 0, x[3], x[1].registered or datetime(1900, 1, 1)))
                selected = [(best[0], best[1], best[2])]

        if selected:
            for index, _, _ in selected:
                used.add(index)
            base["판매금액"] = sum(sale.amount for _, sale, _ in selected)
            if len(selected) > 1:
                status = f"자동매칭({len(selected)}건 합산)"
            elif selected[0][2] < 0.65:
                status = "자동매칭(품목명 차이)"
            else:
                status = "자동매칭"
            base["매칭상태"] = status
            result_rows.append(base)
            for _, sale, _ in selected:
                match_log.append(
                    {
                        "기재용 원본행": original_number,
                        "밴더사": clean_text(base["밴더사"]),
                        "기재용 셀러": clean_text(base["셀러"]),
                        "기재용 품목": clean_text(base["품목"]),
                        "스룩 상품코드": sale.code,
                        "상품등록일": sale.registered,
                        "최근주문일": sale.recent,
                        "스룩 셀러": sale.seller,
                        "스룩 품목": sale.item,
                        "정상금액": sale.amount,
                        "매칭상태": status,
                    }
                )
            continue

        base["판매금액"] = None
        base["매칭상태"] = "매출 미확인"
        result_rows.append(base)

        # 셀러명이 비슷하지만 완전히 동일하지 않으면 원본 행 아래 후보 행을 추가한다.
        if len(sellers) == 1 and sellers[0] and clean_text(base["품목"]):
            entry_seller = normalize_seller(sellers[0])
            fuzzy = []
            for index, sale in enumerate(sales_rows):
                if index in used or not sale.seller or not date_eligible(start, end, sale):
                    continue
                sale_seller = normalize_seller(sale.seller)
                seller_ratio = difflib.SequenceMatcher(None, entry_seller, sale_seller).ratio()
                contains = (entry_seller in sale_seller or sale_seller in entry_seller) and min(len(entry_seller), len(sale_seller)) >= 2
                product_ratio = item_similarity(base["품목"], sale.item)
                if (seller_ratio >= 0.78 or contains) and product_ratio >= 0.5:
                    fuzzy.append((index, sale, seller_ratio, product_ratio))

            for index, sale, seller_ratio, _ in sorted(fuzzy, key=lambda x: (x[2], x[3], x[1].amount != 0), reverse=True)[:3]:
                used.add(index)
                candidate = {column: base.get(column, "") for column in ENTRY_REQUIRED}
                candidate["셀러"] = sale.seller
                candidate["셀러 링크"] = ""
                candidate["팔로워"] = ""
                candidate["품목"] = sale.item
                candidate["판매금액"] = sale.amount
                candidate["비고"] = "스룩 유사 셀러 후보"
                candidate["매칭상태"] = f"유사 후보: {clean_text(base['셀러'])} ↔ {sale.seller}"
                result_rows.append(candidate)
                match_log.append(
                    {
                        "기재용 원본행": original_number,
                        "밴더사": clean_text(base["밴더사"]),
                        "기재용 셀러": clean_text(base["셀러"]),
                        "기재용 품목": clean_text(base["품목"]),
                        "스룩 상품코드": sale.code,
                        "상품등록일": sale.registered,
                        "최근주문일": sale.recent,
                        "스룩 셀러": sale.seller,
                        "스룩 품목": sale.item,
                        "정상금액": sale.amount,
                        "매칭상태": f"유사 후보(셀러 유사도 {seller_ratio:.0%})",
                    }
                )

    unmatched = []
    for index, sale in enumerate(sales_rows):
        if index in used:
            continue
        if sale.seller:
            if normalize_seller(sale.seller) in entry_seller_norms:
                reason = "동일 셀러는 있으나 품목·일정 또는 별도 상품코드 확인 필요"
            else:
                reason = "기재용 파일에 동일 셀러 없음"
            kind = "셀러 상품"
        else:
            reason = "셀러명이 없어 자동매칭 제외"
            kind = "셀러 미표기/자사상품"
        if sale.amount == 0:
            reason += " / 정상금액 0원"
        unmatched.append(
            {
                "구분": kind,
                "상품코드": sale.code,
                "상품등록일": sale.registered,
                "최근주문일": sale.recent,
                "셀러": sale.seller,
                "품목": sale.item,
                "정상금액": sale.amount,
                "확인사항": reason,
            }
        )

    main_df = pd.DataFrame(result_rows, columns=ENTRY_REQUIRED + ["매칭상태"])
    unmatched_df = pd.DataFrame(unmatched)
    if not unmatched_df.empty:
        unmatched_df["_sort"] = unmatched_df["구분"].map({"셀러 상품": 0, "셀러 미표기/자사상품": 1}).fillna(2)
        unmatched_df = unmatched_df.sort_values(["_sort", "상품등록일", "셀러", "품목"]).drop(columns=["_sort"]).reset_index(drop=True)
    log_df = pd.DataFrame(match_log)
    return main_df, unmatched_df, log_df


def export_result(main_df: pd.DataFrame, unmatched_df: pd.DataFrame, log_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm") as writer:
        main_df.to_excel(writer, index=False, sheet_name="6월매출기재용")
        unmatched_df.to_excel(writer, index=False, sheet_name="스룩만 있음")
        log_df.to_excel(writer, index=False, sheet_name="매칭내역", startrow=7)

        workbook = writer.book
        header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#44546A", "align": "center", "valign": "vcenter", "border": 1})
        body = workbook.add_format({"valign": "vcenter", "border": 1})
        center = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1})
        money = workbook.add_format({"num_format": "#,##0", "align": "right", "valign": "vcenter", "border": 1})
        link = workbook.add_format({"font_color": "#0563C1", "underline": True, "border": 1, "valign": "vcenter"})
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd hh:mm", "align": "center", "border": 1})
        matched = workbook.add_format({"bg_color": "#E2F0D9", "font_color": "#375623", "align": "center", "border": 1})
        warning = workbook.add_format({"bg_color": "#FFF2CC", "font_color": "#7F6000", "align": "center", "border": 1})
        missing = workbook.add_format({"bg_color": "#FCE4D6", "font_color": "#C00000", "align": "center", "border": 1})

        main = writer.sheets["6월매출기재용"]
        widths = [10, 8, 28, 18, 18, 42, 10, 40, 15, 22, 24]
        for idx, width in enumerate(widths):
            main.set_column(idx, idx, width, body)
        main.set_column(0, 4, None, center)
        main.set_column(5, 5, 42, link)
        main.set_column(6, 6, 10, center)
        main.set_column(8, 8, 15, money)
        main.freeze_panes(1, 0)
        main.autofilter(0, 0, len(main_df), len(main_df.columns) - 1)
        main.set_row(0, 26, header)
        for col, name in enumerate(main_df.columns):
            main.write(0, col, name, header)
        for row_index, status in enumerate(main_df["매칭상태"].tolist(), start=1):
            if "품목명 차이" in clean_text(status) or clean_text(status).startswith("유사 후보"):
                fmt = warning
            elif clean_text(status) == "매출 미확인":
                fmt = missing
            else:
                fmt = matched
            main.write(row_index, 10, status, fmt)

        unmatched = writer.sheets["스룩만 있음"]
        unmatched_widths = [20, 16, 19, 19, 18, 46, 15, 42]
        for idx, width in enumerate(unmatched_widths):
            unmatched.set_column(idx, idx, width, body)
        unmatched.set_column(2, 3, 19, date_format)
        unmatched.set_column(6, 6, 15, money)
        unmatched.freeze_panes(1, 0)
        unmatched.autofilter(0, 0, len(unmatched_df), max(0, len(unmatched_df.columns) - 1))
        unmatched.set_row(0, 26, header)
        for col, name in enumerate(unmatched_df.columns):
            unmatched.write(0, col, name, header)
        if len(unmatched_df):
            unmatched.conditional_format(1, 0, len(unmatched_df), len(unmatched_df.columns) - 1, {"type": "formula", "criteria": '=$A2="셀러 상품"', "format": workbook.add_format({"bg_color": "#FFF2CC"})})
            unmatched.conditional_format(1, 0, len(unmatched_df), len(unmatched_df.columns) - 1, {"type": "formula", "criteria": '=$A2="셀러 미표기/자사상품"', "format": workbook.add_format({"bg_color": "#E7E6E6"})})

        log = writer.sheets["매칭내역"]
        log.merge_range("A1:K1", "6월 매출 자동 매칭 결과", workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#203864", "font_size": 16, "align": "center", "valign": "vcenter"}))
        labels = ["기재용 행 수", "매출 입력 행", "매출 미확인", "스룩만 있음", "입력 매출 합계"]
        formulas = [
            f"=COUNTA('6월매출기재용'!A2:A{len(main_df)+1})",
            f"=COUNT('6월매출기재용'!I2:I{len(main_df)+1})",
            f'=COUNTIF(\'6월매출기재용\'!K2:K{len(main_df)+1},"매출 미확인")',
            f"=COUNTA('스룩만 있음'!B2:B{len(unmatched_df)+1})",
            f"=SUM('6월매출기재용'!I2:I{len(main_df)+1})",
        ]
        label_fmt = workbook.add_format({"bold": True, "bg_color": "#D9E2F3", "border": 1})
        for row_index, (label, formula) in enumerate(zip(labels, formulas), start=1):
            log.write(row_index, 0, label, label_fmt)
            log.write_formula(row_index, 1, formula, money if row_index == 5 else body)
        log_widths = [13, 18, 18, 38, 16, 19, 19, 18, 46, 15, 24]
        for idx, width in enumerate(log_widths):
            log.set_column(idx, idx, width, body)
        log.set_column(5, 6, 19, date_format)
        log.set_column(9, 9, 15, money)
        log.freeze_panes(8, 0)
        for col, name in enumerate(log_df.columns):
            log.write(7, col, name, header)

    output.seek(0)
    return output.getvalue()


def process_files(sales_bytes: bytes, sales_name: str, entry_bytes: bytes, entry_name: str):
    sales_rows = parse_sales_file(sales_bytes, sales_name)
    entry_df = parse_entry_file(entry_bytes, entry_name)
    main_df, unmatched_df, log_df = merge_sales(entry_df, sales_rows)
    excel_bytes = export_result(main_df, unmatched_df, log_df)
    return main_df, unmatched_df, log_df, excel_bytes
