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

# --- 세션 상태 초기화 ---
if "qty_value" not in st.session_state:
    st.session_state.qty_value = 0
if "success_msg" not in st.session_state:
    st.session_state.success_msg = ""

# --- 유통기한 임박 계산 로직 ---
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

# 사이드바 유통기한 알림
st.title("☕ VINI COFFEE 안락동점 통합 재고관리 시스템")
if imminent_items:
    with st.sidebar.expander("🚨 유통기한 임박 품목 알림", expanded=True):
        for item in imminent_items:
            st.warning(item)

# ★ 메뉴 구조에 '📋 실시간 현재고 현황판' 추가
menu = st.sidebar.radio("메뉴 이동", ["📝 일별 입출고 기록", "📋 실시간 현재고 현황판", "📊 월별 수불 대장 및 백업 다운로드"])

# 1) 매일 입출고 기록 화면
if menu == "📝 일별 입출고 기록":
    st.subheader("일별 입출고 등록")
    st.info("💡 숫자를 입력한 뒤 바로 **엔터(Enter) 키**를 누르면 즉시 저장되고 입력창이 비워집니다!")
    
    if st.session_state.success_msg:
        st.success(st.session_state.success_msg)
        st.session_state.success_msg = "" 
    
    categories = list(master_data['대분류'].unique())
    selected_cat = st.selectbox("1. 품목 분류 선택", categories)
    
    filtered_items = master_data[master_data['대분류'] == selected_cat]
    item_list = filtered_items['품목명'].drop_duplicates().tolist()
    
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
        
        quantity = st.number_input("4. 수량 입력 후 엔터(Enter)", min_value=0, step=1, key="qty_value")
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
                
                st.session_state.success_msg = f"✅ 즉시 기록 완료: {selected_date} | {selected_item} | {type_io} {quantity}개"
                st.session_state.qty_value = 0
                st.rerun()
            else:
                st.warning("⚠️ 수량을 1개 이상 입력하셔야 기록됩니다.")

    # 오늘의 입력 내역 현황 및 수정/삭제
    st.markdown("---")
    st.subheader("🔍 오늘 실시간 입력된 내역")
    log_df = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    today_df = log_df[log_df["날짜"] == today_str]
    
    if not today_df.empty:
        display_df = today_df.copy()
        display_df.insert(0, '기록번호', today_df.index)
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        with st.expander("🛠️ 방금 입력한 기록 수정 / 삭제하기", expanded=False):
            options = today_df.index.tolist()
            def make_label(idx):
                row = today_df.loc[idx]
                return f"번호[{idx}] - {row['품목명']} ({row['구분']} {row['수량']}개)"
                
            col_m1, col_m2, col_m3 = st.columns([2, 1, 2])
            with col_m1:
                selected_idx = st.selectbox("수정/삭제할 기록 줄 선택", options, format_func=make_label)
            with col_m2:
                manage_action = st.radio("작업 선택", ["수량 수정", "기록 삭제"])
            with col_m3:
                current_target = today_df.loc[selected_idx]
                if manage_action == "수량 수정":
                    new_qty = st.number_input("변경할 새 수량 입력", min_value=1, step=1, value=int(current_target['수량']))
                    execute_btn = st.button("🔧 수량 수정 완료")
                else:
                    st.write(f"⚠️ 정말로 이 기록을 삭제하시겠습니까?")
                    execute_btn = st.button("❌ 선택 기록 삭제 확정")
                    
            if execute_btn:
                if manage_action == "기록 삭제":
                    log_df = log_df.drop(selected_idx).reset_index(drop=True)
                    st.session_state.success_msg = "❌ 선택하신 내역이 정상적으로 삭제되었습니다."
                elif manage_action == "수량 수정":
                    log_df.loc[selected_idx, '수량'] = new_qty
                    st.session_state.success_msg = f"🔧 수량이 {new_qty}개로 정상 수정되었습니다."
                log_df.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')
                st.rerun()
    else:
        st.caption("오늘 아직 입력된 내역이 없습니다.")

# ★ 2) [신규 메뉴] 실시간 현재고 현황판 화면
elif menu == "📋 실시간 현재고 현황판":
    st.subheader("📊 매장 실시간 현재고 현황판")
    
    log_df = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
    
    # 누적된 데이터를 바탕으로 실시간 총 입고/총 출고 계산
    if not log_df.empty:
        pivot_all = log_df.pivot_table(
            index=['대분류', '품목명'], 
            columns='구분', 
            values='수량', 
            aggfunc='sum'
        ).fillna(0).reset_index()
        if "금월 입고" not in pivot_all.columns: pivot_all["금월 입고"] = 0
        if "월 소모(출고)" not in pivot_all.columns: pivot_all["월 소모(출고)"] = 0
    else:
        pivot_all = pd.DataFrame(columns=['대분류', '품목명', '금월 입고', '월 소모(출고)'])
        
    # 마스터 데이터와 병합하여 전체 현재고 대시보드 생성
    dashboard_df = pd.merge(master_data[['대분류', '품목명', '엑셀기본재고', '유통기한']], pivot_all, on=['대분류', '품목명'], how='left').fillna(0)
    dashboard_df['현재고'] = dashboard_df['엑셀기본재고'] + dashboard_df['금월 입고'] - dashboard_df['월 소모(출고)']
    dashboard_df['현재고'] = dashboard_df['현재고'].astype(int)
    
    # 컬럼명 예쁘게 정리
    dashboard_df = dashboard_df.rename(columns={"엑셀기본재고": "기본재고(이월)", "금월 입고": "누적 입고량", "월 소모(출고)": "누적 소모량"})
    
    # 대시보드 상단 요약 카드 (KPI Metrics)
    total_count = len(dashboard_df)
    shortage_count = len(dashboard_df[dashboard_df['현재고'] <= 0])
    expiry_count = len(imminent_items)
    
    col_card1, col_card2, col_card3 = st.columns(3)
    with col_card1:
        st.metric(label="📦 관리 중인 전체 품목 수", value=f"{total_count}개")
    with col_card2:
        st.metric(label="⚠️ 재고 부족 및 품절 품목", value=f"{shortage_count}개", delta=f"-{shortage_count}" if shortage_count > 0 else "0", delta_color="inverse")
    with col_card3:
        st.metric(label="⏳ 유통기한 임박 품목 (30일 이내)", value=f"{expiry_count}개")
        
    st.markdown("---")
    
    # 검색 및 대분류 필터 UI
    col_f1, col_f2 = st.columns([1, 2])
    with col_f1:
        filter_cat = st.radio("분류별 보기", ["전체"] + list(master_data['대분류'].unique()), horizontal=True)
    with col_f2:
        search_query = st.text_input("🔍 품목 이름 검색 (예: 우유, 컵, 커피)", "")
        
    # 필터링 적용
    display_dash = dashboard_df.copy()
    if filter_cat != "전체":
        display_dash = display_dash[display_dash['대분류'] == filter_cat]
    if search_query:
        display_dash = display_dash[display_dash['품목명'].str.contains(search_query, case=False)]
        
    # 사용자가 보기 편하게 재고가 적은 순(품절 위험 순)으로 기본 정렬하여 노출
    display_dash = display_dash.sort_values(by="현재고", ascending=True)
    
    st.dataframe(display_dash, use_container_width=True, hide_index=True)

# 3) 월별 조회 및 백업 화면
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
            summary['실시간 예상 현재고'] = summary['엑셀기본재고'] + summary['총 입고량'] - summary['총 소모량']
            
            st.markdown(f"### 📅 {selected_month} 품목별 종합 수불 집계")
            st.dataframe(summary, use_container_width=True)
            
            st.markdown("---")
            st.subheader("💾 데이터 안전 백업 및 내보내기")
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                st.download_button(
                    label=f"📊 {selected_month} 월간 집계표 다운로드",
                    data=summary.to_csv(index=False, encoding='utf-8-sig'),
                    file_name=f"vini_coffee_summary_{selected_month}.csv",
                    mime="text/csv"
                )
            with col_b2:
                st.download_button(
                    label="📝 전체 일별 로그 백업 (CSV)",
                    data=pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig').to_csv(index=False, encoding='utf-8-sig'),
                    file_name="vini_daily_stock_log_backup.csv",
                    mime="text/csv"
                )
                )
