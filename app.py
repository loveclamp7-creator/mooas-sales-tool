from __future__ import annotations

import hashlib

import streamlit as st

from example_files import FINAL_EXAMPLE_BYTES, INTERMEDIATE_EXAMPLE_BYTES
from matcher import process_files


APP_VERSION = "4.0.0"

st.set_page_config(
    page_title="공동구매 매출 자동 매칭",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 2rem;
            padding-bottom: 4rem;
        }

        [data-testid="stSidebar"] {
            min-width: 280px;
            max-width: 280px;
        }

        .title {
            font-size: 2rem;
            font-weight: 800;
            margin-bottom: 0.25rem;
        }

        .subtitle {
            color: #667085;
            margin-bottom: 1.4rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# 왼쪽 메뉴
# -----------------------------
with st.sidebar:
    st.title("🛠️ 업무 자동화 도구")

    st.radio(
        "메뉴 선택",
        ["📊 매출 자동 매칭"],
        index=0,
    )

    st.divider()

    st.caption(f"v{APP_VERSION} · 결제금액/정상금액 선택")

    st.markdown("---")

    st.markdown(
        """
        <div style="
            text-align: center;
            color: #98A2B3;
            font-size: 12px;
            line-height: 1.7;
            padding: 6px 0 12px 0;
        ">
            © 2026 Developed by MINJEEWON<br>
            MOOAS Sales Automation
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------
# 메인 화면
# -----------------------------
st.title("📊 스룩 매출 자동 매칭")

st.markdown(
    """
    <div class="subtitle">
        스룩 상품별 매출 파일과 매출 기재용 파일을 함께 올리면,
        동일 셀러·품목을 찾아 판매금액을 자동 입력합니다.
    </div>
    """,
    unsafe_allow_html=True,
)

mode = st.radio(
    "매출 관리 기준",
    [
        "중간 매출 관리 (결제금액 기준)",
        "최종 매출 관리 (정상금액 기준)",
    ],
    horizontal=True,
)

is_intermediate = mode.startswith("중간")
amount_column = "결제금액" if is_intermediate else "정상금액"
example_bytes = (
    INTERMEDIATE_EXAMPLE_BYTES
    if is_intermediate
    else FINAL_EXAMPLE_BYTES
)
example_name = (
    "중간 매출 관리 (결제금액기준) 예시파일.xlsx"
    if is_intermediate
    else "최종 매출 관리 (정상금액기준) 예시파일.xlsx"
)

st.info(
    f"현재 선택: {mode}\n\n"
    "완전 일치·상품명 축약은 자동 입력하고, "
    "셀러명이 비슷하지만 다른 경우에는 원본 행을 위에 둔 채 "
    "후보 행을 아래에 따로 추가합니다. "
    "스룩에만 있는 상품은 별도 시트로 정리됩니다."
)


# -----------------------------
# 파일 업로드
# -----------------------------
left, right = st.columns(2)

with left:
    sales_file = st.file_uploader(
        f"① {mode} 원본 파일",
        type=["xlsx", "xls", "csv"],
        help=(
            "상품코드·상품명·상품등록일·최근주문일·"
            f"{amount_column} 열이 있는 파일"
        ),
        key=f"sales_file_{amount_column}",
    )

    if is_intermediate:
        st.caption(
            "스룩 > 매출/정산 > 상품별매출관리 > "
            "기간설정 검색 > 엑셀 내려받기 후 파일을 첨부하세요."
        )
    else:
        st.caption(
            "스룩 > 매출/정산 > 상품별매출관리 > "
            "기간설정 검색 > 아래 내용을 긁어서 엑셀 새 파일에 "
            "붙여넣은 후 첨부하세요. 상단 내용도 보이게 붙여넣어 주세요!"
        )

    st.download_button(
        f"📎 {mode} 예시파일 다운로드",
        data=example_bytes,
        file_name=example_name,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
        key=f"example_{amount_column}",
    )

with right:
    entry_file = st.file_uploader(
        "② 매출 기재용 파일",
        type=["xlsx", "xls", "csv"],
        help=(
            "년·월·진행일·밴더사·셀러·품목·"
            "판매금액 열이 있는 파일"
        ),
        key="entry_file",
    )


# 파일이 없을 때 여기에서 정지
# 제작자 표시는 위쪽 사이드바에 있어서 항상 보임
if sales_file is None or entry_file is None:
    st.caption("두 파일을 모두 올리면 자동 매칭 결과가 나타납니다.")
    st.stop()


signature = hashlib.sha256(
    sales_file.getvalue() + entry_file.getvalue()
).hexdigest()


# -----------------------------
# 파일 처리
# -----------------------------
try:
    main_df, unmatched_df, log_df, excel_bytes = process_files(
        sales_file.getvalue(),
        sales_file.name,
        entry_file.getvalue(),
        entry_file.name,
        amount_column,
    )

except Exception as error:
    st.error(
        "파일을 처리하지 못했습니다.\n\n"
        f"{error}"
    )
    st.stop()


# -----------------------------
# 결과 요약
# -----------------------------
matched_count = int(
    main_df["판매금액"].notna().sum()
)

missing_count = int(
    (main_df["매칭상태"] == "매출 미확인").sum()
)

matched_amount = int(
    main_df["판매금액"].fillna(0).sum()
)

metric1, metric2, metric3, metric4 = st.columns(4)

metric1.metric(
    "기재용 행",
    f"{len(main_df):,}건",
)

metric2.metric(
    "매출 입력",
    f"{matched_count:,}건",
)

metric3.metric(
    "매출 미확인",
    f"{missing_count:,}건",
)

metric4.metric(
    f"입력 매출 합계 ({amount_column})",
    f"{matched_amount:,.0f}원",
)


# -----------------------------
# 결과 표
# -----------------------------
main_tab, unmatched_tab, log_tab = st.tabs(
    [
        "최종 매출기재용",
        "스룩만 있음",
        "매칭내역",
    ]
)

with main_tab:
    st.dataframe(
        main_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "판매금액": st.column_config.NumberColumn(
                "판매금액",
                format="%,d원",
            ),
            "매칭상태": st.column_config.TextColumn(
                "매칭상태",
                width="medium",
            ),
        },
    )

with unmatched_tab:
    if unmatched_df.empty:
        st.success(
            "스룩에만 있는 항목이 없습니다."
        )

    else:
        st.warning(
            "기재용 파일에서 찾지 못한 스룩 항목이 "
            f"{len(unmatched_df):,}건 있습니다."
        )

        st.dataframe(
            unmatched_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "정상금액": st.column_config.NumberColumn(
                    "정상금액",
                    format="%,d원",
                )
            },
        )

with log_tab:
    st.dataframe(
        log_df,
        use_container_width=True,
        hide_index=True,
    )


# -----------------------------
# 엑셀 다운로드
# -----------------------------
st.download_button(
    "📥 최종 매출 기재용 엑셀 다운로드",
    data=excel_bytes,
    file_name=f"{mode}_자동매칭.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True,
    key=f"download_{signature}",
)
