import streamlit as st
import pandas as pd
import datetime
import os

# 파일 이름 설정
ORIGINAL_EXCEL_PATH = "VINI_COFFEE_통합_식자재_및_매출관리_시스템_v3_간략시트연동.xlsx"
STOCK_LOG_FILE = "vini_daily_stock_log.csv"

st.set_page_config(page_title="VINI COFFEE 재고관리 시스템", layout="wide")

@st.cache_data
def load_excel_master():
    """기존 엑셀 시트에서 최초 마스터 품목 정보를 파싱하는 함수"""
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
                            temp_df['유통기한'] = df[expiry_col[0]].astype(str).str.strip()
                        else:
                            temp_df['유통기한'] = ""
                            
                        if stock_col:
                            temp_df['엑셀기본재고'] = pd.to_numeric(df[stock_col[0]], errors='coerce').fillna(0).astype(int)
                        else:
                            temp_df['엑셀기본재고'] = 0
                            
                        master_df_list.append(temp_df)
                        
            if master_df_list:
                return pd.concat(master_df_list, ignore_index=True)
        except Exception as e:
            st.error(f"엑셀 품목 로드 실패: {e}")
            
    return pd.DataFrame([
        {"품목명": "커피(퍼플-마스터피스)", "대분류": "원재료", "유통기한": "2027-04-13", "엑셀기본재고": 54},
        {"품목명": "서울우유", "대분류": "원재료", "유통기한": "2026-06-25", "엑셀기본재고": 42}
    ])

# 일별 로그 파일 초기화
if not os.path.exists(STOCK_LOG_FILE):
    df_empty = pd.DataFrame(columns=["날짜", "대분류", "품목명", "구분", "수량", "유통기한"])
    df_empty.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')

# 세션 상태(State) 안전 초기화
if "success_msg" not in st.session_state: st.session_state.success_msg = ""
if "warning_msg" not in st.session_state: st.session_state.warning_msg = ""
if "in_qty" not in st.session_state: st.session_state.in_qty = 0
if "out_qty" not in st.session_state: st.session_state.out_qty = 0

# 데이터 로드 및 유통기한 실시간 매핑
base_master = load_excel_master()
master_data = base_master.copy()
log_df = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
if not log_df.empty and "유통기한" in log_df.columns:
    expiry_logs = log_df[log_df["유통기한"].notna() & (log_df["유통기한"].astype(str) != "") & (log_df["유통기한"].astype(str) != "nan")]
    if not expiry_logs.empty:
        latest_expiries = expiry_logs.sort_values(by="날짜").groupby("품목명").last()["유통기한"].to_dict()
        master_data["유통기한"] = master_data["품목명"].map(latest_expiries).fillna(master_data["유통기한"])

# 유통기한 임박 계산
today = datetime.date.today()
imminent_items = []
for idx, row in master_data.iterrows():
    utg = str(row['유통기한']).strip()
    if utg and utg != '0' and utg != '1899-12-31' and utg != 'nan':
        try:
            expiry_date = pd.to_datetime(utg).date()
            days_left = (expiry_date - today).days
            if 0 <= days_left <= 30:
                imminent_items.append(f"• **{row['품목명']}** ({days_left}일 남음)")
        except:
            pass

# --- 콜백 함수 구역 (에러 해결의 핵심) ---
def save_inbound_callback():
    """입고 데이터 저장 후 수량창을 안전하게 0으로 비우는 콜백 함수"""
    qty = st.session_state.in_qty
    if qty > 0:
        current_logs = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
        new_data = pd.DataFrame([{
            "날짜": st.session_state.in_date.strftime("%Y-%m-%d"),
            "대분류": st.session_state.in_cat,
            "품목명": st.session_state.in_item,
            "구분": "금월 입고",
            "수량": qty,
            "유통기한": st.session_state.in_utg.strftime("%Y-%m-%d")
        }])
        current_logs = pd.concat([current_logs, new_data], ignore_index=True)
        current_logs.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')
        st.session_state.success_msg = f"📥 입고 완료: {st.session_state.in_item} | {qty}개 | 유통기한: {st.session_state.in_utg}"
        st.session_state.in_qty = 0  # 위젯이 그려지기 전에 세션을 안전하게 초기화
    else:
        st.session_state.warning_msg = "⚠️ 입고 수량을 1개 이상 입력하셔야 합니다."

def save_outbound_callback():
    """출고 데이터 저장 후 수량창을 안전하게 0으로 비우는 콜백 함수"""
    qty = st.session_state.out_qty
    if qty > 0:
        current_logs = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
        new_data = pd.DataFrame([{
            "날짜": st.session_state.out_date.strftime("%Y-%m-%d"),
            "대분류": st.session_state.out_cat,
            "품목명": st.session_state.out_item,
            "구분": "월 소모(출고)",
            "수량": qty,
            "유통기한": ""
        }])
        current_logs = pd.concat([current_logs, new_data], ignore_index=True)
        current_logs.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')
        st.session_state.success_msg = f"📤 출고 완료: {st.session_state.out_item} | {qty}개 소모"
        st.session_state.out_qty = 0  # 위젯이 그려지기 전에 세션을 안전하게 초기화
    else:
        st.session_state.warning_msg = "⚠️ 출고 수량을 1개 이상 입력하셔야 합니다."


# --- UI 레이아웃 ---
st.title("☕ VINI COFFEE 안락동점 통합 재고관리 시스템")

if imminent_items:
    with st.sidebar.expander("🚨 유통기한 임박 품목 알림", expanded=True):
        for item in imminent_items:
            st.warning(item)

menu = st.sidebar.radio(
    "메뉴 이동", 
    [
        "📥 물품 입고 등록 (+유통기한)", 
        "📤 물품 출고 등록 (소모)", 
        "📋 실시간 현재고 현황판", 
        "📊 월별 수불 대장 및 백업 다운로드"
    ]
)

# 알림 메시지 출력 구역
if st.session_state.success_msg:
    st.success(st.session_state.success_msg)
    st.session_state.success_msg = ""
if st.session_state.warning_msg:
    st.warning(st.session_state.warning_msg)
    st.session_state.warning_msg = ""

# 공통 기능: 오늘 입력 내역 및 수정/삭제
def show_today_logs_and_management():
    st.markdown("---")
    st.subheader("🔍 오늘 실시간 입력된 내역")
    current_logs = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    today_df = current_logs[current_logs["날짜"] == today_str]
    
    if not today_df.empty:
        display_df = today_df.copy()
        display_df.insert(0, '기록번호', today_df.index)
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        with st.expander("🛠️ 방금 입력한 기록 수정 / 삭제하기", expanded=False):
            options = today_df.index.tolist()
            def make_label(idx):
                r = today_df.loc[idx]
                utg_str = f" | 유통기한: {r['유통기한']}" if pd.notna(r['유통기한']) and str(r['유통기한']) != 'nan' and r['유통기한'] != '' else ""
                return f"번호[{idx}] - {r['품목명']} ({r['구분']} {r['수량']}개{utg_str})"
                
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
                    updated_logs = current_logs.drop(selected_idx).reset_index(drop=True)
                    st.session_state.success_msg = "❌ 선택하신 내역이 정상적으로 삭제되었습니다."
                elif manage_action == "수량 수정":
                    current_logs.loc[selected_idx, '수량'] = new_qty
                    updated_logs = current_logs
                    st.session_state.success_msg = f"🔧 수량이 {new_qty}개로 정상 수정되었습니다."
                current_logs.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')
                st.rerun()
    else:
        st.caption("오늘 아직 입력된 내역이 없습니다.")


# 1) 입고 등록 메뉴
if menu == "📥 물품 입고 등록 (+유통기한)":
    st.subheader("📥 매장 물품 입고 등록")
    st.info("💡 입고 수량과 유통기한을 맞춘 후 **엔터(Enter) 키**를 누르면 즉시 저장 및 비워집니다.")
    
    categories = list(master_data['대분류'].unique())
    selected_cat = st.selectbox("1. 입고 품목 분류 선택", categories, key="in_cat")
    
    filtered_items = master_data[master_data['대분류'] == selected_cat]
    item_list = filtered_items['품목명'].drop_duplicates().tolist()
    
    with st.form("inbound_form"):
        col1, col2 = st.columns(2)
        with col1:
            st.date_input("날짜 선택", datetime.date.today(), key="in_date")
            selected_item = st.selectbox("2. 입고 품목 선택", item_list, key="in_item")
        with col2:
            item_info = filtered_items[filtered_items['품목명'] == selected_item].iloc[0]
            existing_utg = str(item_info['유통기한']).strip()
            
            default_utg_date = datetime.date.today()
            if existing_utg and existing_utg != '0' and existing_utg != '1899-12-31' and existing_utg != 'nan':
                try: default_utg_date = pd.to_datetime(existing_utg).date()
                except: pass
                
            st.caption(f"📊 기존 정보 ➡️ [현재 엑셀상 기본재고: {item_info['엑셀기본재고']}개 / 유통기한: {existing_utg}]")
            st.date_input("3. 유통기한 확인 및 변경 설정", default_utg_date, key="in_utg")
            
        st.number_input("4. 입고 수량 입력 후 엔터(Enter)", min_value=0, step=1, key="in_qty")
        # ★ 핵심 바인딩: 버튼 클릭 및 엔터 시 콜백 함수 실행
        st.form_submit_button("📥 입고 데이터 저장하기", on_click=save_inbound_callback)
                
    show_today_logs_and_management()

# 2) 출고 등록 메뉴
elif menu == "📤 물품 출고 등록 (소모)":
    st.subheader("📤 매장 소모(출고) 등록")
    st.info("💡 출고(소모) 수량을 입력한 뒤 **엔터(Enter) 키**를 누르면 즉시 저장 및 비워집니다.")
    
    categories = list(master_data['대분류'].unique())
    selected_cat = st.selectbox("1. 출고 품목 분류 선택", categories, key="out_cat")
    
    filtered_items = master_data[master_data['대분류'] == selected_cat]
    item_list = filtered_items['품목명'].drop_duplicates().tolist()
    
    with st.form("outbound_form"):
        col1, col2 = st.columns(2)
        with col1:
            st.date_input("날짜 선택", datetime.date.today(), key="out_date")
            selected_item = st.selectbox("2. 출고 품목 선택", item_list, key="out_item")
        with col2:
            item_info = filtered_items[filtered_items['품목명'] == selected_item].iloc[0]
            st.caption(f"📊 현재 재고 참고 ➡️ [엑셀상 기본재고: {item_info['엑셀기본재고']}개 / 유통기한: {item_info['유통기한']}]")
            
        st.number_input("3. 소모(출고) 수량 입력 후 엔터(Enter)", min_value=0, step=1, key="out_qty")
        # ★ 핵심 바인딩: 버튼 클릭 및 엔터 시 콜백 함수 실행
        st.form_submit_button("📤 출고 데이터 저장하기", on_click=save_outbound_callback)
                
    show_today_logs_and_management()

# 3) 실시간 현재고 현황판 (가시성 최고)
elif menu == "📋 실시간 현재고 현황판":
    st.subheader("📋 매장 실시간 현재고 현황판")
    
    current_logs = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
    
    if not current_logs.empty:
        pivot_all = current_logs.pivot_table(
            index=['대분류', '품목명'], 
            columns='구분', 
            values='수량', 
            aggfunc='sum'
        ).fillna(0).reset_index()
        if "금월 입고" not in pivot_all.columns: pivot_all["금월 입고"] = 0
        if "월 소모(출고)" not in pivot_all.columns: pivot_all["월 소모(출고)"] = 0
    else:
        pivot_all = pd.DataFrame(columns=['대분류', '품목명', '금월 입고', '월 소모(출고)'])
        
    dashboard_df = pd.merge(master_data[['대분류', '품목명', '엑셀기본재고', '유통기한']], pivot_all, on=['대분류', '품목명'], how='left').fillna(0)
    dashboard_df['현재고'] = dashboard_df['엑셀기본재고'] + dashboard_df['금월 입고'] - dashboard_df['월 소모(출고)']
    dashboard_df['현재고'] = dashboard_df['현재고'].astype(int)
    
    dashboard_df = dashboard_df.rename(columns={"엑셀기본재고": "기본재고(이월)", "금월 입고": "누적 입고량", "월 소모(출고)": "누적 소모량"})
    
    total_count = len(dashboard_df)
    shortage_count = len(dashboard_df[dashboard_df['현재고'] <= 0])
    expiry_count = len(imminent_items)
    
    col_card1, col_card2, col_card3 = st.columns(3)
    with col_card1: st.metric(label="📦 관리 품목 수", value=f"{total_count}개")
    with col_card2: st.metric(label="⚠️ 품절 및 재고부족 품목", value=f"{shortage_count}개")
    with col_card3: st.metric(label="⏳ 유통기한 임박 품목", value=f"{expiry_count}개")
        
    st.markdown("---")
    
    col_f1, col_f2 = st.columns([1, 2])
    with col_f1: filter_cat = st.radio("분류별 필터", ["전체"] + list(master_data['대분류'].unique()), horizontal=True)
    with col_f2: search_query = st.text_input("🔍 품목 실시간 키워드 검색", "")
        
    display_dash = dashboard_df.copy()
    if filter_cat != "전체": display_dash = display_dash[display_dash['대분류'] == filter_cat]
    if search_query: display_dash = display_dash[display_dash['품목명'].str.contains(search_query, case=False)]
        
    display_dash = display_dash.sort_values(by="현재고", ascending=True)
    
    def highlight_shortage(row):
        styles = [''] * len(row)
        if row['현재고'] <= 0:
            idx = row.index.get_loc('현재고')
            styles[idx] = 'background-color: #FADBD8; color: #78281F; font-weight: bold;'
        return styles

    if not display_dash.empty:
        styled_dash = display_dash.style.apply(highlight_shortage, axis=1)
        st.dataframe(styled_dash, use_container_width=True, hide_index=True)
    else:
        st.caption("검색 조건에 맞는 품목이 없습니다.")

# 4) 월별 조회 및 백업 화면
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
            summary = pd.merge(master_data[['대분류', '품목명', '엑셀기본재고', '유통기한']], summary, on=['대분류', '품목명'], how='inner')
            summary['실시간 예상 현재고'] = summary['엑셀기본재고'] + summary['총 입고량'] - summary['총 소모량']
            
            st.markdown(f"### 📅 {selected_month} 품목별 종합 수불 집계")
            st.dataframe(summary, use_container_width=True, hide_index=True)
            
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
