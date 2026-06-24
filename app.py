import streamlit as st
import pandas as pd
import datetime
import os

# 파일 이름 설정 (GitHub에 올린 파일명과 정확히 일치해야 합니다)
ORIGINAL_EXCEL_PATH = "VINI_COFFEE_통합_식자재_및_매출관리_시스템_v3_간략시트연동.xlsx"
STOCK_LOG_FILE = "vini_daily_stock_log.csv"

st.set_page_config(page_title="VINI COFFEE 재고관리", layout="wide")

@st.cache_data
def load_item_list():
    """엑셀의 간략 시트들에서 품목명을 자동으로 추출하는 함수"""
    sheets = ["원재료(간략)", "부자재(간략)", "디저트&완제품(간략)"]
    all_items = {}
    
    if os.path.exists(ORIGINAL_EXCEL_PATH):
        try:
            for sheet in sheets:
                df = pd.read_excel(ORIGINAL_EXCEL_PATH, sheet_name=sheet, skiprows=2)
                df = df.dropna(subset=['품목이름'])
                df = df[df['품목이름'] != 0]
                # 문자열 정제 및 중복 제거
                items = [str(x).strip() for x in df['품목이름'].tolist() if str(x).strip() != '']
                all_items[sheet.replace("(간략)", "")] = sorted(list(set(items)))
            return all_items
        except Exception as e:
            st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
    
    # 파일이 없거나 오류 시 사용할 기본 백업 데이터
    return {
        "원재료": ["커피(퍼플-마스터피스)", "디카페인커피(퍼를-감성고릴라)", "서울우유"],
        "부자재": ["아이스 컵(메가)  (줄)", "핫 컵  (줄)", "빨대 (소) 커피용"],
        "디저트&완제품": ["고르곤졸라피자(시카고)", "페페로니 피자(시카고)", "콤비네이션  씬피자"]
    }

# 일별 로그 저장용 CSV 파일 초기화
if not os.path.exists(STOCK_LOG_FILE):
    df_empty = pd.DataFrame(columns=["날짜", "대분류", "품목명", "구분", "수량"])
    df_empty.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')

item_dict = load_item_list()

# --- 사이드바 및 UI ---
st.title("☕ VINI COFFEE 안락동점 재고관리 시스템")
menu = st.sidebar.radio("메뉴 이동", ["📝 일별 입출고 기록", "📊 월별 수불 대장 조회"])

# 1) 매일 입출고 기록 화면
if menu == "📝 일별 입출고 기록":
    st.subheader("일별 입출고 등록")
    
    with st.form("input_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            selected_date = st.date_input("날짜 선택", datetime.date.today())
            category = st.selectbox("품목 분류", list(item_dict.keys()))
        with col2:
            item_list = item_dict[category]
            selected_item = st.selectbox("품목 선택", item_list)
            type_io = st.selectbox("입/출고 구분", ["금월 입고", "월 소모(출고)"])
            
        quantity = st.number_input("수량 입력", min_value=0, step=1, value=0)
        submit_btn = st.form_submit_button("기록 저장하기")
        
        if submit_btn:
            if quantity > 0:
                log_df = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
                new_data = pd.DataFrame([{
                    "날짜": selected_date.strftime("%Y-%m-%d"),
                    "대분류": category,
                    "품목명": selected_item,
                    "구분": type_io,
                    "수량": quantity
                }])
                log_df = pd.concat([log_df, new_data], ignore_index=True)
                log_df.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')
                st.success(f"✅ 저장 완료: {selected_date} | {selected_item} | {type_io} {quantity}개")
            else:
                st.warning("⚠️ 수량을 1개 이상 입력해주세요.")

    # 오늘의 입력 내역 현황
    st.markdown("---")
    st.subheader("🔍 오늘 입력된 내역")
    log_df = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    today_df = log_df[log_df["날짜"] == today_str]
    if not today_df.empty:
        st.dataframe(today_df, use_container_width=True)
    else:
        st.caption("오늘 입력된 내역이 없습니다.")

# 2) 월별 조회 화면
elif menu == "📊 월별 수불 대장 조회":
    st.subheader("월별 수불 대장 조회")
    log_df = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
    
    if log_df.empty:
        st.info("아직 누적된 데이터가 없습니다. 입출고를 먼저 기록해주세요.")
    else:
        log_df['날짜'] = pd.to_datetime(log_df['날짜'])
        log_df['년월'] = log_df['날짜'].dt.to_period('M').astype(str)
        
        available_months = sorted(log_df['년월'].unique(), reverse=True)
        selected_month = st.selectbox("조회할 월 선택", available_months)
        selected_cat = st.selectbox("분류 필터", ["전체"] + list(item_dict.keys()))
        
        filtered_df = log_df[log_df['년월'] == selected_month]
        if selected_cat != "전체":
            filtered_df = filtered_df[filtered_df['대분류'] == selected_cat]
            
        if filtered_df.empty:
            st.warning("선택한 조건에 맞는 데이터가 없습니다.")
        else:
            summary = filtered_df.pivot_table(
                index=['대분류', '품목명'], 
                columns='구분', 
                values='수량', 
                aggfunc='sum'
            ).fillna(0).reset_index()
            
            if "금월 입고" not in summary.columns: summary["금월 입고"] = 0
            if "월 소모(출고)" not in summary.columns: summary["월 소모(출고)"] = 0
                
            summary = summary.rename(columns={"금월 입고": "총 입고량", "월 소모(출고)": "총 소모량"})
            
            st.markdown(f"### 📅 {selected_month} 품목별 입출고 요약")
            st.dataframe(summary, use_container_width=True)
            
            with st.expander("정렬된 월간 상세 입력 로그 보기"):
                st.dataframe(filtered_df.sort_values(by="날짜"), use_container_width=True)
