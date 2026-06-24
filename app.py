import streamlit as st
import pandas as pd
import datetime
import os

# 파일 이름 설정
ORIGINAL_EXCEL_PATH = "VINI_COFFEE_통합_식자재_및_매출관리_시스템_v3_간략시트연동.xlsx"
STOCK_LOG_FILE = "vini_daily_stock_log.csv"

st.set_page_config(page_title="VINI COFFEE 재고관리 시스템", layout="wide")

@st.cache_data
def load_master_data():
    """기존 엑셀 시트에서 품목 정보를 파싱하되, 엑셀의 원본 행 순서를 그대로 유지하는 함수"""
    sheets_to_try = {
        "원재료": ["원재료(간략)", "원재료"],
        "부자재": ["부자재(간략)", "부자재"],
        "디저트&완제품": ["디저트&완제품(간략)", "디저트&완제품"]
    }
    
    master_df_list = []
    
    if os.path.exists(ORIGINAL_EXCEL_PATH):
        try:
            xl = pd.ExcelFile(ORIGINAL_EXCEL_PATH)
            for cat, standard_names in sheets_to_try.items():
                target_sheet = None
                for name in xl.sheet_names:
                    if name.strip() in standard_names or cat in name:
                        target_sheet = name
                        break
                
                if target_sheet:
                    df = pd.read_excel(xl, sheet_name=target_sheet, skiprows=2)
                    df.columns = [str(c).strip().replace(" ", "") for c in df.columns]
                    
                    # 품목명 컬럼 찾기
                    name_col = None
                    for col in df.columns:
                        if '품목이름' in col or '구분' in col:
                            name_col = col
                            break
                    
                    if name_col:
                        df = df.dropna(subset=[name_col])
                        df = df[df[name_col].astype(str).str.strip() != '0']
                        df = df[df[name_col].astype(str).str.strip() != '']
                        
                        expiry_col = [c for c in df.columns if '유통' in c]
                        stock_col = [c for c in df.columns if '재고' in c]
                        
                        temp_df = pd.DataFrame()
                        temp_df['품목명'] = df[name_col].astype(str).str.strip()
                        temp_df['대분류'] = cat
                        
                        if expiry_col:
                            temp_df['유통기한'] = df[expiry_col[0]]
                        else:
                            temp_df['유통기한'] = None
                            
                        if stock_col:
                            temp_df['엑셀기본재고'] = pd.to_numeric(df[stock_col[0]], errors='coerce').fillna(0).astype(int)
                        else:
                            temp_df['엑셀기본재고'] = 0
                            
                        master_df_list.append(temp_df)
                        
            if master_df_list:
                return pd.concat(master_df_list, ignore_index=True)
        except Exception as e:
            st.error(f"엑셀 품목 로드 실패 (기본값 전환): {e}")
            
    return pd.DataFrame([
        {"품목명": "커피(퍼플-마스터피스)", "대분류": "원재료", "유통기한": "2027-04-13", "엑셀기본재고": 54},
        {"품목명": "서울우유", "대분류": "원재료", "유통기한": "2026-06-25", "엑셀기본재고": 42}
    ])

if not os.path.exists(STOCK_LOG_FILE):
    df_empty = pd.DataFrame(columns=["날짜", "대분류", "품목명", "구분", "수량"])
    df_empty.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')

master_data = load_master_data()

# --- 유통기한 임박 알림 ---
st.title("☕ VINI COFFEE 안락동점 통합 재고관리 시스템")

today = datetime.date.today()
imminent_items = []

for idx, row in master_data.iterrows():
    if pd.notna(row['유통기한']) and row['유통기한'] != 0 and str(row['유통기한']) != '1899-12-31':
        try:
            expiry_date = pd.to_datetime(row['유통기한']).date()
            days_left = (expiry_date - today).days
            if 0 <= days_left <= 30:
                imminent_items.append(f"• **{row['품목명']}** ({days_left}일 남음)")
        except:
            pass

if imminent_items:
    with st.sidebar.expander("🚨 유통기한 임박 품목 알림", expanded=True):
        for item in imminent_items:
            st.warning(item)

menu = st.sidebar.radio("메뉴 이동", ["📝 일별 입출고 기록", "📊 월별 수불 대장 및 백업 다운로드"])

# 1) 매일 입출고 기록 화면
if menu == "📝 일별 입출고 기록":
    st.subheader("일별 입출고 등록")
    st.info("💡 숫자를 입력한 뒤 바로 **엔터(Enter) 키**를 누르면 즉시 저장됩니다!")
    
    categories = list(master_data['대분류'].unique())
    selected_cat = st.selectbox("1. 품목 분류 선택", categories)
    
    filtered_items = master_data[master_data['대분류'] == selected_cat]
    item_list = filtered_items['품목명'].drop_duplicates().tolist()
    
    # ★핵심 수정: 타이밍 오류를 유발하는 clear_on_submit 옵션을 제거했습니다.
    with st.form("input_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            selected_date = st.date_input("날짜 선택", datetime.date.today())
        with col2:
            selected_item = st.selectbox("2. 품목 선택", item_list)
        with col3:
            type_io = st.selectbox("3. 입/출고 구분", ["금월 입고", "월 소모(출고)"])
            
        item_info = filtered_items[filtered_items['품목명'] == selected_item].iloc[0]
        expiry_text = f" / 유통기한: {str(item_info['유통기한'])[:10]}" if pd.notna(item_info['유통기한']) else ""
        st.caption(f"📊 선택 품목 정보 ➡️ [엑셀 기본재고: {item_info['엑셀기본재고']}개{expiry_text}]")
        
        # 수량 입력창
        quantity = st.number_input("4. 수량 입력 후 엔터(Enter)", min_value=0, step=1, value=0)
        submit_btn = st.form_submit_button("💾 기록 저장하기 (또는 엔터)")
        
        if submit_btn:
            if quantity > 0:
                log_df = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
                new_data = pd.DataFrame([{
                    "날짜": selected_date.strftime("%Y-%m-%d"),
                    "대분류": selected_cat,
                    "품목명": selected_item,
                    "구분": type_io,
                    "수량": quantity
                }])
                log_df = pd.concat([log_df, new_data], ignore_index=True)
                log_df.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')
                st.success(f"✅ 즉시 기록 완료: {selected_date} | {selected_item} | {type_io} {quantity}개")
            else:
                st.warning("⚠️ 수량을 1개 이상 입력하셔야 기록됩니다.")

    # 오늘의 입력 내역 현황
    st.markdown("---")
    st.subheader("🔍 오늘 실시간 입력된 내역")
    log_df = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    today_df = log_df[log_df["날짜"] == today_str]
    if not today_df.empty:
        st.dataframe(today_df, use_container_width=True)
    else:
        st.caption("오늘 아직 입력된 내역이 없습니다.")

# 2) 월별 조회 및 백업 화면
elif menu == "📊 월별 수불 대장 및 백업 다운로드":
    st.subheader("월별 수불 대장 및 데이터 다운로드")
    log_df = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
    
    if log_df.empty:
        st.info("아직 누적된 데이터가 없습니다.")
    else:
        log_df['날짜'] = pd.to_datetime(log_df['날짜'])
        log_df['년월'] = log_df['날짜'].dt.to_period('M').astype(str)
        
        col1, col2 = st.columns(2)
        with col1:
            available_months = sorted(log_df['년월'].unique(), reverse=True)
            selected_month = st.selectbox("조회할 월 선택", available_months)
        with col2:
            selected_cat = st.selectbox("분류 필터", ["전체"] + list(master_data['대분류'].unique()))
            
        filtered_df = log_df[log_df['년월'] == selected_month]
        if selected_cat != "전체":
            filtered_df = filtered_df[filtered_df['대분류'] == selected_cat]
            
        if filtered_df.empty:
            st.warning("선택한 조건에 맞는 기록이 없습니다.")
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
            summary = pd.merge(master_data[['대분류', '품목명', '엑셀기본재고']], summary, on=['대분류', '품목명'], how='inner')
            summary
