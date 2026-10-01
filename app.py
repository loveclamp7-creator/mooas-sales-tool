from __future__ import annotations

import hashlib
from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st

from matcher import process_files


APP_VERSION = "5.0.0"
st.set_page_config(page_title="무아스 공동구매 매출 자동 정리", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
  .block-container {padding-top:2rem;padding-bottom:4rem;}
  [data-testid="stSidebar"] {min-width:280px;max-width:280px;}
  .subtitle {color:#667085;margin-bottom:1.2rem;}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.title("🛠️ 업무 자동화 도구")
    st.radio("메뉴 선택", ["📊 공동구매 매출 정리"], index=0)
    st.divider()
    st.caption(f"v{APP_VERSION} · 주문번호 기준 매출 매칭")
    st.markdown("---")
    st.markdown('<div style="text-align:center;color:#98A2B3;font-size:12px;line-height:1.7">© 2026 Developed by MINJEEWON<br>MOOAS Sales Automation</div>', unsafe_allow_html=True)

st.title("📊 무아스 공동구매 매출 자동 정리")
st.markdown('<div class="subtitle">스룩 주문내역과 차액 파일을 주문번호로 연결해, 무아스 공동구매 누적 리스트의 판매금액 공란을 채웁니다.</div>', unsafe_allow_html=True)

month_options = [f"{month}월" for month in range(1, 13)]
current_month = datetime.now(ZoneInfo("Asia/Seoul")).month
selected_month_label = st.selectbox("확인할 매출월", month_options, index=current_month - 1)
target_month = int(selected_month_label.replace("월", ""))

st.info(
    "처리 기준\n\n"
    "• 차액 파일은 `상세내역_통합` 탭의 주문번호와 ‘셀러명 × 무아스 상품명’을 확인합니다.\n"
    "• 실제 판매금액은 스룩 주문내역의 같은 주문번호 `★결제액(수정가능)`을 합산합니다.\n"
    "• 오른쪽 누적 리스트의 기존 금액은 유지하고, 선택 월의 판매금액 공란만 채웁니다.\n"
    "• 비즈브릭스·레몬트리·지엠홀딩스는 판매금액을 공란으로 유지합니다."
)

left, right = st.columns(2)
with left:
    source_files = st.file_uploader(
        "① 매출 원본 파일들", type=["xlsx", "xls"], accept_multiple_files=True,
        help="(주)스룩_오클릭 파일과 상세내역_통합 탭이 있는 차액 파일을 함께 올려주세요. 여러 파일도 가능합니다."
    )
    st.caption("필수: `(주)스룩_오클릭.xls` + `차액_YYMM.xlsx` · 왼쪽은 여러 파일 첨부 가능")
with right:
    template_file = st.file_uploader(
        "② 결과 양식 파일", type=["xlsx"],
        help="무아스 공동구매 리스트(26년도).xlsx처럼 년·월·밴더사·셀러·딜명·판매금액 열이 있는 누적 파일입니다."
    )
    st.caption("`무아스 공동구매 리스트(26년도).xlsx`를 올리면 기존 서식과 다른 월의 내용은 그대로 유지됩니다.")

with st.expander("구글 캘린더 일정 보완 안내"):
    st.write(
        "이 사이트는 개인 구글 계정 권한을 직접 보관하지 않기 때문에 구글 캘린더를 자동 조회할 수는 없습니다. "
        "대신 매출 원본에는 있지만 누적 리스트에 없는 셀러·공구를 결과 파일에 자동으로 추가하고, 비고에 `일정/밴더 확인 필요`라고 표시합니다. "
        "이 행만 구글 캘린더와 확인하면 누락 일정도 빠르게 보완할 수 있습니다."
    )

if not source_files or template_file is None:
    st.caption("왼쪽에 두 종류의 매출 원본을, 오른쪽에 누적 리스트를 올리면 자동 정리됩니다.")
    st.stop()

signature = hashlib.sha256(b"".join(file.getvalue() for file in source_files) + template_file.getvalue()).hexdigest()
try:
    result_df, audit_df, excel_bytes, deals, recognized, target_year = process_files(
        [(file.getvalue(), file.name) for file in source_files], template_file.getvalue(), template_file.name, target_month
    )
except Exception as error:
    st.error(f"파일을 처리하지 못했습니다.\n\n{error}")
    st.stop()

auto_count = int((result_df["상태"] == "자동 입력").sum())
added_count = int((result_df["상태"] == "리스트에 없어 행 추가").sum())
excluded_count = int(result_df["상태"].isin(["지정 밴더 공란 유지", "누락 행 추가 · 지정 밴더 공란"]).sum())
input_total = int(result_df.loc[result_df["상태"].isin(["자동 입력", "리스트에 없어 행 추가"]), "판매금액"].fillna(0).sum())

st.success(f"{target_year}년 {target_month}월 매출 정리가 완료됐습니다.")
metric1, metric2, metric3, metric4 = st.columns(4)
metric1.metric("판매금액 입력", f"{auto_count:,}건")
metric2.metric("누락 행 자동 추가", f"{added_count:,}건")
metric3.metric("지정 밴더 공란", f"{excluded_count:,}건")
metric4.metric("입력 매출 합계", f"{input_total:,.0f}")

result_tab, order_tab, file_tab = st.tabs(["결과 확인", "주문번호 매칭", "인식한 파일"])
with result_tab:
    st.dataframe(result_df, use_container_width=True, hide_index=True, column_config={"판매금액": st.column_config.NumberColumn("판매금액", format="%,d")})
with order_tab:
    st.dataframe(audit_df, use_container_width=True, hide_index=True, column_config={"결제액": st.column_config.NumberColumn("결제액", format="%,d")})
with file_tab:
    for label in recognized:
        st.write(f"✓ {label}")

st.download_button(
    "📥 매출 입력 완료 파일 다운로드", data=excel_bytes,
    file_name=f"무아스_공동구매_리스트_{target_year}년_{target_month}월_매출입력완료.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True, key=f"download_{signature}"
)
