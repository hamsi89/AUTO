import streamlit as st
import pandas as pd
import datetime
import io
import json
import gspread

# [필수] 페이지 설정은 앱 전체 최상단에 딱 1번만 실행되어야 오류가 발생하지 않습니다.
st.set_page_config(page_title="VINI COFFEE 재고관리 시스템", page_icon="☕", layout="wide")

# =========================================================================
# ⚙️ 구글 스프레드시트 주소 설정 (본인의 시트 URL 주소로 꼭 변경하세요!)
# =========================================================================
VINI_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_VINI_COFFEE_SHEET_ID/edit"
WINE_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_WINE_SHEET_ID/edit"

# 원본 엑셀의 진짜 시트 이름 지정
SHEET_MAP = {
    "원재료": "원재료",
    "부자재": "부자재",
    "디저트&완제품": "디저트&완제품"
}

# =========================================================================
# 🔒 구글 클라우드 OAuth 인증 및 세션 로그인 관리
# =========================================================================
raw_creds = st.secrets.get("gcp_credentials_json", "")
raw_auth = st.secrets.get("gcp_authorized_user_json", "")

if not raw_creds or not raw_auth:
    st.error("Streamlit Secrets에 구글 인증 정보가 등록되지 않았거나 비어있습니다.")
    st.stop()

try:
    creds_dict = json.loads(raw_creds.strip())
    auth_user_dict = json.loads(raw_auth.strip())
except json.decoder.JSONDecodeError as e:
    st.error("⚠️ 구글 인증서 JSON 형식이 올바르지 않습니다. Secrets 입력을 확인해 주세요.")
    st.stop()

# gspread 인증 및 시트 전역 연결
@st.cache_resource
def get_gspread_client():
    return gspread.oauth_from_dict(creds_dict, auth_user_dict)

gc = get_gspread_client()

try:
    vini_sh = gc.open_by_url(VINI_SPREADSHEET_URL)
    wine_sh = gc.open_by_url(WINE_SPREADSHEET_URL)
except Exception as e:
    st.error(f"❌ 구글 스프레드시트 로드 실패: URL 주소 또는 공유 권한(편집자)을 확인하세요. 에러: {e}")
    st.stop()

# 로그인 유무 세션 초기화
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("🔒 매장 재고 관리 시스템")
    st.subheader("관리자 인증이 필요합니다.")
    
    password_input = st.text_input("관리자 비밀번호를 입력하세요", type="password")
    login_button = st.button("로그인")

    if login_button or password_input:
        if password_input == "122000":
            st.session_state.logged_in = True
            st.success("인증에 성공했습니다! 잠시만 기다려주세요...")
            st.rerun()  
        else:
            st.error("비밀번호가 올바르지 않습니다. 다시 입력해주세요.")
    st.stop() 

# =========================================================================
# 🛠️ 구글 시트 전용 데이터 헬퍼 함수 (로컬 파일 I/O 대체)
# =========================================================================
def get_or_create_tab(sh, tab_title, headers):
    """시트에 특정 탭이 없으면 자동으로 헤더를 넣어서 생성해 주는 함수"""
    try:
        return sh.worksheet(tab_title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_title, rows="2000", cols="20")
        ws.append_row(headers)
        return ws

def load_gsheets_master_raw(sh):
    """기존 대장 시트에서 최초 품목 정보를 실시간 파싱"""
    sheets_to_try = {
        "원재료": ["원재료", "원재료(간략)"],
        "부자재": ["부자재", "부자재(간략)"],
        "디저트&완제품": ["디저트&완제품", "디저트&완제품(간략)"]
    }
    master_df_list = []
    all_sheet_names = [sheet.title for sheet in sh.worksheets()]
    
    for cat, standard_names in sheets_to_try.items():
        target_sheet = None
        for name in all_sheet_names:
            if name.strip() in standard_names or cat in name:
                target_sheet = name
                break
        
        if target_sheet:
            ws = sh.worksheet(target_sheet)
            all_vals = ws.get_all_values()
            if len(all_vals) <= 2: continue
            
            df = pd.DataFrame(all_vals[3:], columns=all_vals[2])
            df.columns = [str(c).strip().replace(" ", "") for c in df.columns]
            
            name_col = None
            for col in df.columns:
                if '품목이름' in col or '구분' in col or '품목명' in col:
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
                temp_df['유통기한'] = df[expiry_col[0]].astype(str).str.strip() if expiry_col else ""
                temp_df['엑셀기본재고'] = pd.to_numeric(df[stock_col[0]], errors='coerce').fillna(0).astype(int) if stock_col else 0
                master_df_list.append(temp_df)
                
    if master_df_list:
        return pd.concat(master_df_list, ignore_index=True)
    return pd.DataFrame(columns=["품목명", "대분류", "유통기한", "엑셀기본재고"])

def get_latest_master_and_logs_from_gsheets(sh, is_wine=False):
    """로컬 CSV 대신 구글 시트 내부의 전용 탭에서 마스터와 로그를 가져옴"""
    master_tab_name = "wine_custom_master" if is_wine else "vini_custom_master"
    log_tab_name = "wine_daily_stock_log" if is_wine else "vini_daily_stock_log"
    
    # 마스터 로드 및 자동 생성
    try:
        ws_master = sh.worksheet(master_tab_name)
        master_vals = ws_master.get_all_values()
        master = pd.DataFrame(master_vals[1:], columns=master_vals[0])
    except gspread.exceptions.WorksheetNotFound:
        if is_wine:
            # 와인 마스터 파싱 기본 로직
            master = pd.DataFrame(columns=["품목명", "대분류", "유통기한", "엑셀기본재고"])
            try:
                ws_w = sh.get_worksheet(0)
                w_vals = ws_w.get_all_values()
                if len(w_vals) > 2:
                    w_df = pd.DataFrame(w_vals[3:], columns=w_vals[2])
                    w_df.columns = [str(c).strip().replace(" ", "") for c in w_df.columns]
                    n_col = [c for c in w_df.columns if '품목' in c or '와인명' in c or '이름' in c][0]
                    s_col = [c for c in w_df.columns if '재고' in c or '이월' in c]
                    e_col = [c for c in w_df.columns if '유통' in c or '빈티지' in c]
                    master['품목명'] = w_df[n_col].astype(str).str.strip()
                    master['대분류'] = "와인"
                    master['유통기한'] = w_df[e_col[0]].astype(str).str.strip() if e_col else ""
                    master['엑셀기본재고'] = pd.to_numeric(w_df[s_col[0]], errors='coerce').fillna(0).astype(int) if s_col else 0
            except:
                master = pd.DataFrame([{"품목명": "레드와인_A", "대분류": "와인", "유통기한": "", "엑셀기본재고": 10}])
        else:
            master = load_gsheets_master_raw(sh)
            
        ws_master = sh.add_worksheet(title=master_tab_name, rows="1000", cols="10")
        ws_master.append_row(master.columns.tolist())
        if not master.empty:
            ws_master.append_rows(master.values.tolist())
            
    # 로그 로드 및 자동 생성
    ws_log = get_or_create_tab(sh, log_tab_name, ["날짜", "대분류", "품목명", "구분", "수량", "유통기한"])
    log_vals = ws_log.get_all_values()
    logs = pd.DataFrame(log_vals[1:], columns=log_vals[0])
    
    master['엑셀기본재고'] = pd.to_numeric(master['엑셀기본재고'], errors='coerce').fillna(0).astype(int)
    master['유통기한'] = master['유통기한'].fillna("").astype(str)
    
    if not logs.empty and "유통기한" in logs.columns:
        logs['수량'] = pd.to_numeric(logs['수량'], errors='coerce').fillna(0).astype(int)
        expiry_logs = logs[logs["유통기한"].notna() & (logs["유통기한"].astype(str) != "") & (logs["유통기한"].astype(str) != "nan")]
        if not expiry_logs.empty:
            latest_expiries = expiry_logs.sort_values(by="날짜").groupby("품목명").last()["유통기한"].to_dict()
            master["유통기한"] = master["품목명"].map(latest_expiries).fillna(master["유통기한"])
            
    return master, logs

def update_gsheets_vini_rule(sh, cat, item_name, qty, target_date, new_expiry=None, action="inbound", is_wine=False):
    """오픈픽슬(openpyxl) 대장 수정 기능을 구글 시트 실시간 셀 업데이트로 변경"""
    try:
        if is_wine:
            ws = sh.get_worksheet(0)
        else:
            target_sheet = SHEET_MAP.get(cat, cat)
            ws = sh.worksheet(target_sheet)
            
        all_vals = ws.get_all_values()
        if len(all_vals) < 3: return False
        
        headers = [str(c).strip().replace(" ", "") for c in all_vals[2]] # 3번째 행 복사
        name_col_idx, expiry_col_idx, target_col_idx = None, None, None
        target_day_num = str(target_date.day)
        target_day_label = f"{target_date.day}일"
        
        for idx, val in enumerate(headers, start=1):
            if '품목이름' in val or '품목명' in val or '와인' in val or '이름' in val or '구분' in val: name_col_idx = idx
            elif '유통' in val or '빈티지' in val: expiry_col_idx = idx
            
            if action == "inbound" and ('금월입고' in val or '입고' in val): target_col_idx = idx
            elif action == "outbound" and (val == target_day_label or val == target_day_num or val == f"0{target_day_num}"): target_col_idx = idx

        if not name_col_idx or not target_col_idx: return False

        for r_idx, row in enumerate(all_vals[3:], start=4):
            if len(row) < name_col_idx: continue
            if str(row[name_col_idx - 1]).strip() == item_name:
                cur_val_str = row[target_col_idx - 1] if len(row) >= target_col_idx else "0"
                try: current_val = int(cur_val_str) if cur_val_str else 0
                except: current_val = 0
                
                ws.update_cell(r_idx, target_col_idx, current_val + qty)
                if new_expiry and expiry_col_idx:
                    ws.update_cell(r_idx, expiry_col_idx, new_expiry)
                return True
        return False
    except Exception as e:
        st.error(f"구글 대장 시트 동기화 실패: {e}")
        return False

# 데이터 바인딩 (실시간)
master_data, log_df = get_latest_master_and_logs_from_gsheets(vini_sh, is_wine=False)
wine_master_data, wine_log_df = get_latest_master_and_logs_from_gsheets(wine_sh, is_wine=True)

# 세션 상태 변수 세팅
for key in ["success_msg", "warning_msg"]:
    if key not in st.session_state: st.session_state[key] = ""
for key in ["in_qty", "out_qty"]:
    if key not in st.session_state: st.session_state[key] = 0

# 유통기한 계산 로직
today = datetime.date.today()
imminent_items = []
for idx, row in master_data.iterrows():
    utg = str(row['유통기한']).strip()
    if utg and utg not in ['0', '1899-12-31', 'nan', '']:
        try:
            expiry_date = pd.to_datetime(utg).date()
            days_left = (expiry_date - today).days
            if 0 <= days_left <= 30:
                imminent_items.append(f"• **{row['품목명']}** ({days_left}일 남음)")
        except: pass

# --- 콜백 함수 구역 ---
def save_inbound_callback():
    qty = st.session_state.in_qty
    if qty > 0:
        cat = st.session_state.in_cat
        item = st.session_state.in_item
        in_date_val = st.session_state.in_date
        expiry_str = st.session_state.in_utg.strftime("%Y-%m-%d")
        
        success = update_gsheets_vini_rule(vini_sh, cat, item, qty, in_date_val, new_expiry=expiry_str, action="inbound")
        if success:
            ws_log = vini_sh.worksheet("vini_daily_stock_log")
            ws_log.append_row([in_date_val.strftime("%Y-%m-%d"), cat, item, "금월 입고", qty, expiry_str])
            st.session_state.success_msg = f"📥 [입고 완료] {item} | {qty}개 ➡️ 구글 대장 시트 입고 열에 실시간 누적되었습니다."
            st.session_state.in_qty = 0
    else: st.warning("입고 수량을 입력하세요.")

def save_outbound_callback():
    qty = st.session_state.out_qty
    if qty > 0:
        cat = st.session_state.out_cat
        item = st.session_state.out_item
        out_date_val = st.session_state.out_date
        
        success = update_gsheets_vini_rule(vini_sh, cat, item, qty, out_date_val, action="outbound")
        if success:
            ws_log = vini_sh.worksheet("vini_daily_stock_log")
            ws_log.append_row([out_date_val.strftime("%Y-%m-%d"), cat, item, "월 소모(출고)", qty, ""])
            st.session_state.success_msg = f"📤 [소모 완료] {item} | {qty}개 ➡️ 구글 대장 시트 [{out_date_val.day}일] 열에 정확히 차감/합산되었습니다."
            st.session_state.out_qty = 0
    else: st.warning("출고 수량을 입력하세요.")

# --- 사이드바 메뉴 레이아웃 ---
st.sidebar.title("관리자 메뉴")
if st.sidebar.button("로그아웃"):
    st.session_state.logged_in = False
    st.rerun()

if imminent_items:
    with st.sidebar.expander("🚨 유통기한 임박 품목 알림", expanded=True):
        for item in imminent_items: st.warning(item)

menu = st.sidebar.radio(
    "메뉴 이동", 
    [
        "📥 물품 입고 등록 (개별)", 
        "📤 물품 출고 등록 (소모-개별)", 
        "📝 전품목 일괄 입력 (엑셀 스타일)",
        "🍷 와인 재고 관리 (일괄 입력)",
        "📋 실시간 현재고 현황판", 
        "📊 월별 수불 대장 및 백업 다운로드",
        "⚙️ 품목 추가/삭제 관리",
        "🛠️ 데이터 관리 및 엑셀 동기화"
    ]
)

if st.session_state.success_msg:
    st.success(st.session_state.success_msg)
    st.session_state.success_msg = ""

def show_today_logs_and_management(is_wine=False):
    st.markdown("---")
    st.subheader("🔍 오늘 실시간 입력된 내역")
    sh_target = wine_sh if is_wine else vini_sh
    tab_name = "wine_daily_stock_log" if is_wine else "vini_daily_stock_log"
    
    ws_log = sh_target.worksheet(tab_name)
    log_vals = ws_log.get_all_values()
    if len(log_vals) <= 1:
        st.caption("오늘 아직 입력된 내역이 없습니다.")
        return
        
    current_logs = pd.DataFrame(log_vals[1:], columns=log_vals[0])
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    today_df = current_logs[current_logs["날짜"] == today_str]
    
    if not today_df.empty:
        st.dataframe(today_df, use_container_width=True, hide_index=False)
        st.caption("💡 기록 관리 및 오입력 제어는 구글 스프레드시트 웹페이지에서 직접 행을 행 단위로 관리/삭제하는 것을 추천합니다.")
    else:
        st.caption("오늘 아직 입력된 내역이 없습니다.")

# 1) 입고 등록 메뉴(개별)
if menu == "📥 물품 입고 등록 (개별)":
    st.subheader("📥 매장 물품 입고 등록 (개별)")
    categories = list(master_data['대분류'].unique()) if not master_data.empty else ["원재료", "부자재", "디저트&완제품"]
    selected_cat = st.selectbox("1. 입고 품목 분류 선택", categories, key="in_cat")
    filtered_items = master_data[master_data['대분류'] == selected_cat]
    item_list = filtered_items['품목명'].drop_duplicates().tolist()
    
    if not item_list: st.warning("등록된 품목이 없습니다.")
    else:
        with st.form("inbound_form"):
            col1, col2 = st.columns(2)
            with col1:
                st.date_input("날짜 선택", datetime.date.today(), key="in_date")
                selected_item = st.selectbox("2. 입고 품목 선택", item_list, key="in_item")
            with col2:
                item_info = filtered_items[filtered_items['품목명'] == selected_item].iloc[0]
                existing_utg = str(item_info['유통기한']).strip()
                default_utg_date = datetime.date.today()
                try: default_utg_date = pd.to_datetime(existing_utg).date()
                except: pass
                st.caption(f"📊 기본재고: {item_info['엑셀기본재고']}개 | 기존 유통기한: {existing_utg}")
                st.date_input("3. 유통기한 확인 및 설정", default_utg_date, key="in_utg")
            st.number_input("4. 입고 수량 입력", min_value=0, step=1, key="in_qty")
            st.form_submit_button("📥 입고 데이터 저장하기", on_click=save_inbound_callback)
        show_today_logs_and_management()

# 2) 출고 등록 메뉴(개별)
elif menu == "📤 물품 출고 등록 (소모-개별)":
    st.subheader("📤 매장 소모(출고) 등록 (개별)")
    categories = list(master_data['대분류'].unique()) if not master_data.empty else ["원재료", "부자재", "디저트&완제품"]
    selected_cat = st.selectbox("1. 출고 품목 분류 선택", categories, key="out_cat")
    filtered_items = master_data[master_data['대분류'] == selected_cat]
    item_list = filtered_items['품목명'].drop_duplicates().tolist()
    
    if not item_list: st.warning("등록된 품목이 없습니다.")
    else:
        with st.form("outbound_form"):
            st.date_input("날짜 선택", datetime.date.today(), key="out_date")
            selected_item = st.selectbox("2. 출고 품목 선택", item_list, key="out_item")
            item_info = filtered_items[filtered_items['품목명'] == selected_item].iloc[0]
            st.caption(f"📊 기본재고: {item_info['엑셀기본재고']}개 | 유통기한: {item_info['유통기한']}")
            st.number_input("3. 소모(출고) 수량 입력", min_value=0, step=1, key="out_qty")
            st.form_submit_button("📤 출고 데이터 저장하기", on_click=save_outbound_callback)
        show_today_logs_and_management()

# 3) 전품목 일괄 입력 (엑셀 스타일)
elif menu == "📝 전품목 일괄 입력 (엑셀 스타일)":
    st.subheader("📝 전품목 일괄 입력 (엑셀 스타일)")
    bulk_date = st.date_input("🗓️ 기록할 날짜 선택", datetime.date.today())
    bulk_cat = st.radio("분류 필터링", ["전체"] + list(master_data['대분류'].unique()), horizontal=True)
    
    # 피벗 연산 및 데이터프레임 조립
    if not log_df.empty:
        log_df['수량'] = pd.to_numeric(log_df['수량'], errors='coerce').fillna(0)
        pivot_all = log_df.pivot_table(index=['대분류', '품목명'], columns='구분', values='수량', aggfunc='sum').fillna(0).reset_index()
        if "금월 입고" not in pivot_all.columns: pivot_all["금월 입고"] = 0
        if "월 소모(출고)" not in pivot_all.columns: pivot_all["월 소모(출고)"] = 0
    else:
        pivot_all = pd.DataFrame(columns=['대분류', '품목명', '금월 입고', '월 소모(출고)'])
        
    bulk_df = pd.merge(master_data[['대분류', '품목명', '엑셀기본재고', '유통기한']], pivot_all, on=['대분류', '품목명'], how='left').fillna(0)
    bulk_df['현재고'] = bulk_df['엑셀기본재고'] + bulk_df['금월 입고'] - bulk_df['월 소모(출고)']
    bulk_df['📥 오늘 입고량'] = 0
    bulk_df['📤 오늘 소모량'] = 0
    
    display_bulk = bulk_df.rename(columns={"유통기한": "⏳ 유통기한"})[['대분류', '품목명', '현재고', '📥 오늘 입고량', '📤 오늘 소모량', '⏳ 유통기한']]
    if bulk_cat != "전체": display_bulk = display_bulk[display_bulk['대분류'] == bulk_cat]
    
    edited_bulk = st.data_editor(
        display_bulk,
        column_config={
            "대분류": st.column_config.TextColumn("분류", disabled=True),
            "품목명": st.column_config.TextColumn("품목명", disabled=True),
            "현재고": st.column_config.NumberColumn("현재고", disabled=True),
            "📥 오늘 입고량": st.column_config.NumberColumn("오늘 입고", min_value=0, step=1),
            "📤 오늘 소모량": st.column_config.NumberColumn("오늘 소모", min_value=0, step=1),
        },
        use_container_width=True, hide_index=True
    )
    
    if st.button("💾 위 입력된 모든 내역 구글 시트에 일괄 동기화", use_container_width=True):
        ws_log = vini_sh.worksheet("vini_daily_stock_log")
        with st.spinner("구글 스프레드시트에 일괄 업로드 중..."):
            for idx, row in edited_bulk.iterrows():
                in_v = int(row['📥 오늘 입고량'])
                out_v = int(row['📤 오늘 소모량'])
                p_name = row['품목명']
                p_cat = row['대분류']
                utg_v = str(row['⏳ 유통기한']).strip()
                
                if in_v > 0:
                    update_gsheets_vini_rule(vini_sh, p_cat, p_name, in_v, bulk_date, new_expiry=utg_v, action="inbound")
                    ws_log.append_row([bulk_date.strftime("%Y-%m-%d"), p_cat, p_name, "금월 입고", in_v, utg_v])
                if out_v > 0:
                    update_gsheets_vini_rule(vini_sh, p_cat, p_name, out_v, bulk_date, action="outbound")
                    ws_log.append_row([bulk_date.strftime("%Y-%m-%d"), p_cat, p_name, "월 소모(출고)", out_v, ""])
        st.success("🎉 구글 스프레드시트에 전 품목 수치가 실시간 동기화 완료되었습니다!")
        st.rerun()

# 4) 와인 재고 관리 메뉴 (일괄 입력 엑셀 스타일)
elif menu == "🍷 와인 재고 관리 (일괄 입력)":
    st.subheader("🍷 와인 수불 대장 관리 및 일괄 입력 (엑셀 스타일)")
    w_bulk_date = st.date_input("🗓️ 기록할 날짜 선택", datetime.date.today(), key="w_date")
    
    if not wine_log_df.empty:
        wine_log_df['수량'] = pd.to_numeric(wine_log_df['수량'], errors='coerce').fillna(0)
        w_pivot = wine_log_df.pivot_table(index=['대분류', '품목명'], columns='구분', values='수량', aggfunc='sum').fillna(0).reset_index()
        if "금월 입고" not in w_pivot.columns: w_pivot["금월 입고"] = 0
        if "월 소모(출고)" not in w_pivot.columns: w_pivot["월 소모(출고)"] = 0
    else:
        w_pivot = pd.DataFrame(columns=['대분류', '품목명', '금월 입고', '월 소모(출고)'])
        
    w_bulk_df = pd.merge(wine_master_data[['대분류', '품목명', '엑셀기본재고', '유통기한']], w_pivot, on=['대분류', '품목명'], how='left').fillna(0)
    w_bulk_df['현재고'] = w_bulk_df['엑셀기본재고'] + w_bulk_df['금월 입고'] - w_bulk_df['월 소모(출고)']
    w_bulk_df['📥 오늘 입고량'] = 0
    w_bulk_df['📤 오늘 소모량'] = 0
    
    w_display_bulk = w_bulk_df.rename(columns={"유통기한": "⏳ 빈티지/유통기한"})[['대분류', '품목명', '현재고', '📥 오늘 입고량', '📤 오늘 소모량', '⏳ 빈티지/유통기한']]
    
    edited_wine_bulk = st.data_editor(
        w_display_bulk,
        column_config={
            "대분류": st.column_config.TextColumn("분류", disabled=True),
            "품목명": st.column_config.TextColumn("와인명", disabled=True),
            "현재고": st.column_config.NumberColumn("현재고", disabled=True),
            "📥 오늘 입고량": st.column_config.NumberColumn("오늘 입고", min_value=0, step=1),
            "📤 오늘 소모량": st.column_config.NumberColumn("오늘 소모", min_value=0, step=1),
        },
        use_container_width=True, hide_index=True
    )
    
    if st.button("💾 위 입력된 와인 내역 구글 시트에 일괄 동기화", use_container_width=True):
        ws_w_log = wine_sh.worksheet("wine_daily_stock_log")
        with st.spinner("와인 대장 동기화 중..."):
            for idx, row in edited_wine_bulk.iterrows():
                in_v = int(row['📥 오늘 입고량'])
                out_v = int(row['📤 오늘 소모량'])
                p_name = row['품목명']
                utg_v = str(row['⏳ 빈티지/유통기한']).strip()
                
                if in_v > 0:
                    update_gsheets_vini_rule(wine_sh, "와인", p_name, in_v, w_bulk_date, new_expiry=utg_v, action="inbound", is_wine=True)
                    ws_w_log.append_row([w_bulk_date.strftime("%Y-%m-%d"), "와인", p_name, "금월 입고", in_v, utg_v])
                if out_v > 0:
                    update_gsheets_vini_rule(wine_sh, "와인", p_name, out_v, w_bulk_date, action="outbound", is_wine=True)
                    ws_w_log.append_row([w_bulk_date.strftime("%Y-%m-%d"), "와인", p_name, "월 소모(출고)", out_v, ""])
        st.success("🎉 와인 수불 현황이 실시간 업데이트되었습니다!")
        st.rerun()

# 5) 실시간 현재고 현황판
elif menu == "📋 실시간 현재고 현황판":
    st.subheader("📋 매장 실시간 현재고 현황판")
    
    if not log_df.empty:
        log_df['수량'] = pd.to_numeric(log_df['수량'], errors='coerce').fillna(0)
        pivot_all = log_df.pivot_table(index=['대분류', '품목명'], columns='구분', values='수량', aggfunc='sum').fillna(0).reset_index()
        if "금월 입고" not in pivot_all.columns: pivot_all["금월 입고"] = 0
        if "월 소모(출고)" not in pivot_all.columns: pivot_all["월 소모(출고)"] = 0
    else:
        pivot_all = pd.DataFrame(columns=['대분류', '품목명', '금월 입고', '월 소모(출고)'])
        
    dashboard_df = pd.merge(master_data[['대분류', '품목명', '엑셀기본재고', '유통기한']], pivot_all, on=['대분류', '품목명'], how='left').fillna(0)
    dashboard_df['현재고'] = dashboard_df['엑셀기본재고'] + dashboard_df['금월 입고'] - dashboard_df['월 소모(출고)']
    
    display_dash = dashboard_df.rename(columns={
        "엑셀기본재고": "기본재고(이월)", "금월 입고": "누적 입고량", "월 소모(출고)": "누적 소모량"
    })[["대분류", "품목명", "기본재고(이월)", "누적 입고량", "누적 소모량", "현재고", "유통기한"]]
    
    col_f1, col_f2 = st.columns([1, 2])
    with col_f1: filter_cat = st.radio("분류별 필터", ["전체"] + list(master_data['대분류'].unique()), horizontal=True)
    with col_f2: search_query = st.text_input("🔍 품목 실시간 키워드 검색", "")
    
    if filter_cat != "전체": display_dash = display_dash[display_dash['대분류'] == filter_cat]
    if search_query: display_dash = display_dash[display_dash['품목명'].str.contains(search_query, case=False)]
    
    display_dash = display_dash.sort_values(by="현재고", ascending=True)
    
    def highlight_shortage(row):
        styles = [''] * len(row)
        if int(row['현재고']) <= 0:
            styles[row.index.get_loc('현재고')] = 'background-color: #FADBD8; color: #78281F; font-weight: bold;'
        return styles

    st.dataframe(display_dash.style.apply(highlight_shortage, axis=1), use_container_width=True, hide_index=True)

# 6) 월별 조회 및 백업 화면
elif menu == "📊 월별 수불 대장 및 백업 다운로드":
    st.subheader("월별 수불 대장 분석 및 원본 구글 대장 다운로드")
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        if st.button("📥 VINI COFFEE 전체 대장 Excel 파일 다운로드", use_container_width=True):
            with st.spinner("구글 클라우드에서 최신 엑셀 컴파일 중..."):
                # 구글 드라이브 API 연동 없이 gspread에서 제공하는 xlsx 내보내기 바이너리 활용
                excel_bytes = gc.export(vini_sh.id, format='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                st.download_button("🟢 다운로드 링크 활성화 (클릭)", data=excel_bytes, file_name=f"VINI_COFFEE_대장_{today.strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_dl2:
        if st.button("📥 와인 전체 대장 Excel 파일 다운로드", use_container_width=True):
            with st.spinner("구글 클라우드에서 와인 엑셀 컴파일 중..."):
                wine_bytes = gc.export(wine_sh.id, format='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                st.download_button("🟢 와인 다운로드 링크 활성화 (클릭)", data=wine_bytes, file_name=f"와인_대장_{today.strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 7) 품목 추가 / 삭제 관리 화면
elif menu == "⚙️ 품목 추가/삭제 관리":
    st.subheader("⚙️ 매장 품목 추가 및 삭제 관리 (구글 클라우드 동기화)")
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("### 📥 신규 품목 추가")
        with st.form("add_item_form", clear_on_submit=True):
            add_cat = st.selectbox("품목 분류 선택", ["원재료", "부자재", "디저트&완제품"])
            add_name = st.text_input("새로운 품목명 입력")
            add_stock = st.number_input("초기 기본 이월 재고량", min_value=0, step=1, value=0)
            add_utg = st.date_input("품목 기본 유통기한 지정", datetime.date.today())
            
            if st.form_submit_button("➕ 마스터 품목 추가 완료"):
                if add_name.strip() and add_name.strip() not in master_data['품목명'].values:
                    ws_master = vini_sh.worksheet("vini_custom_master")
                    ws_master.append_row([add_name.strip(), add_cat, add_utg.strftime("%Y-%m-%d"), int(add_stock)])
                    st.success(f"구글 마스터 시트에 [{add_name}] 품목이 안전하게 추가되었습니다.")
                    st.rerun()
                    
    with col_b:
        st.markdown("### ❌ 기존 품목 삭제 (단종)")
        if not master_data.empty:
            all_items = sorted(master_data['품목명'].tolist())
            target_del = st.selectbox("제거할 품목 선택", all_items)
            if st.button("❌ 선택 품목 영구 삭제 확정"):
                ws_master = vini_sh.worksheet("vini_custom_master")
                cells = ws_master.findall(target_del)
                if cells:
                    for cell in reversed(cells): 
                        ws_master.delete_rows(cell.row)
                    st.success(f"구글 마스터 시트에서 [{target_del}] 품목이 영구 제거되었습니다.")
                    st.rerun()

# 8) 데이터 관리 및 엑셀 동기화 통합 메뉴
elif menu == "🛠️ 데이터 관리 및 엑셀 동기화":
    st.subheader("🛠️ 시스템 데이터 관리 및 전체 초기화")
    st.warning("이 섹션은 구글 스프레드시트 클라우드 내부의 가상 데이터베이스(vini_custom_master, vini_daily_stock_log) 탭을 완전히 청소하여 시스템을 리셋하는 기능입니다.")
    
    confirm_destroy = st.checkbox("⚠️ 정말로 모든 가상 로그 탭을 날리고 초기 마스터 상태로 되돌리는 것에 동의합니까?")
    if st.button("🚀 서버 강제 공장 초기화 실행", type="primary", disabled=not confirm_destroy):
        tabs_to_delete = ["vini_custom_master", "vini_daily_stock_log", "wine_custom_master", "wine_daily_stock_log"]
        for tab in tabs_to_delete:
            try:
                ws = vini_sh.worksheet(tab)
                vini_sh.del_worksheet(ws)
            except: pass
            try:
                ws = wine_sh.worksheet(tab)
                wine_sh.del_worksheet(ws)
            except: pass
        st.cache_data.clear()
        st.success("💥 클라우드 캐시 파일 파괴 완수! 구글 대장 시트의 원본 기준으로 원상 복구되었습니다.")
        st.rerun()
