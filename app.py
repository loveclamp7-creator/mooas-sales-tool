from __future__ import annotations

import hashlib
from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st

from matcher import process_files


APP_VERSION = "5.2.0"
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
    st.caption(f"v{APP_VERSION} · 마감 정리 · 중간 매출 확인")
    st.markdown("---")
    st.markdown('<div style="text-align:center;color:#98A2B3;font-size:12px;line-height:1.7">© 2026 Developed by MINJEEWON<br>MOOAS Sales Automation</div>', unsafe_allow_html=True)

st.title("📊 무아스 공동구매 매출 자동 정리")
st.markdown('<div class="subtitle">마감용 원본 3개 또는 중간 확인용 요약표 1개로 공동구매 누적 리스트에 매출을 입력합니다.</div>', unsafe_allow_html=True)

month_options = [f"{month}월" for month in range(1, 13)]
current_month = datetime.now(ZoneInfo("Asia/Seoul")).month
selected_month_label = st.selectbox("확인할 매출월", month_options, index=current_month - 1)
target_month = int(selected_month_label.replace("월", ""))

mode = st.radio("매출 원본 업로드 방식", ["마감 정리 · 원본 3개", "중간 확인 · 요약표 1개"], horizontal=True)
intermediate = mode.startswith("중간")
refresh_existing = False
if intermediate:
    st.info("상품명에서 ‘셀러 × 상품’을 구분해 **정상금액**을 매출로 입력합니다. "
            "선택 월의 기존 일정에 매칭하며, 캘린더 없이 새 일정은 추가하지 않습니다. "
            "셀러명이 없거나 같은 상품의 반복 일정으로 구분이 어려운 매출은 확인용으로 남깁니다.")
    refresh_existing = st.checkbox("매칭된 기존 판매금액도 최신 정상금액으로 갱신", value=False,
                                   help="반복해서 중간 확인할 때 선택하세요. 선택하지 않으면 판매금액 공란만 입력합니다.")
else:
    st.info(
        "처리 기준\n\n"
        "• 차액 파일의 상세내역_통합과 스룩 주문내역을 주문번호로 연결합니다.\n"
        "• 스룩 주문내역의 ★결제액(수정가능)을 합산합니다.\n"
        "• 캘린더 최신본에서 누락 일정을 추가하고 선택 월 판매금액 공란을 채웁니다.\n"
        "• 기존 금액과 다른 월의 내용은 유지합니다."
    )
st.caption("공통: 비즈브릭스·레몬트리·지엠홀딩스는 판매금액을 공란으로 유지합니다.")
with st.expander("매출 원본 업로드 예시 · 2가지 상황", expanded=True):
    st.markdown(
        "**① 마감 정리 — 원본 필수 3개**\n\n"
        "공구 종료 후 주문번호를 대조하고 누락 일정까지 추가할 때 사용합니다.\n\n"
        "1. `(주)스룩_오클릭.xls` — 주문번호·품명·결제액\n"
        "2. `차액_YYMM.xlsx` — 상세내역_통합 탭\n"
        "3. `당월 캘린더 최신본.xlsx` — 시작일·종료일·밴더사·셀러·공구명\n\n"
        "**② 중간 확인 — 원본 1개**\n\n"
        "진행 중인 공구의 현재 매출을 확인할 때 사용합니다.\n\n"
        "- `통합 문서2.xlsx` 같은 상품별 매출 요약표 1개\n"
        "- 필수 열: 상품코드 · 상품명 · **정상금액**\n"
        "- 예: 상품명 `예시셀러 × 무아스 가열식 가습기`, 정상금액 `399,600` → 판매금액 `399,600`\n"
        "- 결제금액·수수료는 매출로 사용하지 않습니다. 조회 기간은 확인할 매출월에 맞춰 내려받으세요.\n\n"
        "**두 방식 모두 오른쪽 결과 양식(누적 리스트) 1개는 별도로 필요합니다.**"
    )

left, right = st.columns(2)
with left:
    source_files = st.file_uploader(
        "① 매출 원본 파일들", type=["xlsx", "xls"], accept_multiple_files=True,
        help="중간 확인은 상품별 요약표 1개, 마감 정리는 스룩 주문내역·차액·캘린더 3개를 올려주세요."
    )
    st.caption("원본 1개: `통합 문서2.xlsx` 같은 상품별 매출 요약표" if intermediate else "필수 3개: `(주)스룩_오클릭.xls` + `차액_YYMM.xlsx` + `당월 캘린더 최신본.xlsx`")
with right:
    template_file = st.file_uploader(
        "② 결과 양식 파일", type=["xlsx"],
        help="무아스 공동구매 리스트(26년도).xlsx처럼 년·월·밴더사·셀러·딜명·판매금액 열이 있는 누적 파일입니다."
    )
    st.caption("`무아스 공동구매 리스트(26년도).xlsx`를 올리면 기존 서식과 다른 월의 내용은 그대로 유지됩니다.")

if not source_files or template_file is None:
    st.caption("선택한 방식의 매출 원본을 왼쪽에, 누적 리스트 1개를 오른쪽에 올려주세요.")
    st.stop()

expected_count = 1 if intermediate else 3
if len(source_files) != expected_count:
    st.warning(f"선택한 방식은 왼쪽 매출 원본 {expected_count}개가 필요합니다. 업로드 파일 수를 확인해주세요.")
    st.stop()

signature = hashlib.sha256(b"".join(file.getvalue() for file in source_files) + template_file.getvalue()).hexdigest()
try:
    result_df, audit_df, excel_bytes, deals, recognized, target_year = process_files(
        [(file.getvalue(), file.name) for file in source_files], template_file.getvalue(), template_file.name, target_month, refresh_existing=refresh_existing
    )
except Exception as error:
    st.error(f"파일을 처리하지 못했습니다.\n\n{error}")
    st.stop()

status_text = result_df["상태"].astype(str)
auto_count = int(status_text.str.contains("자동 입력").sum())
added_count = int(status_text.str.contains("캘린더 추가").sum())
excluded_count = int(status_text.str.contains("지정 밴더 공란").sum())
input_total = int(result_df.loc[status_text.str.contains("자동 입력"), "판매금액"].fillna(0).sum()) if not result_df.empty else 0

if result_df.empty:
    st.warning(f"{target_year}년 {target_month}월에 처리할 일정이 없습니다. 위의 매출월이 업로드한 자료와 맞는지 확인해주세요.")
else:
    st.success(f"{target_year}년 {target_month}월 매출 정리가 완료됐습니다.")
metric1, metric2, metric3, metric4 = st.columns(4)
metric1.metric("판매금액 입력", f"{auto_count:,}건")
metric2.metric("캘린더 일정 추가", f"{added_count:,}건")
metric3.metric("지정 밴더 공란", f"{excluded_count:,}건")
metric4.metric("입력 매출 합계", f"{input_total:,.0f}")

result_tab, order_tab, file_tab = st.tabs(["결과 확인", "상품별 정상금액" if intermediate else "주문번호 매칭", "인식한 파일"])
with result_tab:
    st.dataframe(result_df, use_container_width=True, hide_index=True, column_config={"판매금액": st.column_config.NumberColumn("판매금액", format="%,d")})
with order_tab:
    st.dataframe(audit_df, use_container_width=True, hide_index=True, column_config={"결제액": st.column_config.NumberColumn("결제액", format="%,d"), "정상금액": st.column_config.NumberColumn("정상금액", format="%,d")})
with file_tab:
    for label in recognized:
        st.write(f"✓ {label}")

st.download_button(
    "📥 매출 입력 완료 파일 다운로드", data=excel_bytes,
    file_name=f"무아스_공동구매_리스트_{target_year}년_{target_month}월_매출입력완료.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True, key=f"download_{signature}"
)
