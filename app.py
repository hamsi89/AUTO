# 7) 💡 [데이터 관리 및 엑셀 동기화] 통합 메뉴
elif menu == "🛠️ 데이터 관리 및 엑셀 동기화":
    st.subheader("🛠️ 시스템 데이터 관리 및 엑셀 파일 동기화")
    st.markdown("---")
    
    col_up1, col_up2 = st.columns(2)
    
    # --- [1] VINI COFFEE 일반 식자재/매출 대장 업로드 ---
    with col_up1:
        st.markdown("### 📅 일반 대장 업로드 및 품목 동기화")
        st.info("💡 VINI COFFEE 매출관리 시스템 대장(.xlsx) 파일을 업로드하면 원재료, 부자재, 완제품 현황을 자동으로 파싱하여 동기화합니다.")
        
        uploaded_file = st.file_uploader("일반 대장 파일을 업로드하세요 (.xlsx)", type=["xlsx"], key="vini_uploader")

        if uploaded_file is not None:
            st.success("일반 대장 파일이 웹 브라우저에 준비되었습니다!")
            if st.button("🚀 일반 대장 파싱 및 연동 시작", use_container_width=True):
                try:
                    xl = pd.ExcelFile(uploaded_file)
                    sheets_to_try = {
                        "원재료": ["원재료", "원재료(간략)"],
                        "부자재": ["부자재", "부자재(간략)"],
                        "디저트&완제품": ["디저트&완제품", "디저트&완제품(간략)"]
                    }
                    
                    uploaded_master_list = []
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
                                    
                                uploaded_master_list.append(temp_df)
                    
                    if uploaded_master_list:
                        final_uploaded_master = pd.concat(uploaded_master_list, ignore_index=True)
                        final_uploaded_master.to_csv(CUSTOM_MASTER_FILE, index=False, encoding='utf-8-sig')
                        
                        with open(ORIGINAL_EXCEL_PATH, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                            
                        st.success("🎉 일반 식자재 대장이 마스터에 오차 없이 완벽 연동되었습니다!")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("⚠️ 시트명이나 품목 열을 찾을 수 없습니다.")
                except Exception as e:
                    st.error(f"파싱 도중 에러가 발생했습니다: {e}")

    # --- [2] 🍷 와인 대장 전용 업로드 (추가 항목) ---
    with col_up2:
        st.markdown("### 🍷 와인 대장 업로드 및 품목 동기화")
        st.info("💡 `와인_입출고양식_디자인적용_현재고요약.xlsx` 형태의 와인 전용 대장 파일을 업로드하여 시스템 마스터 품목 정보를 최신으로 갱신합니다.")
        
        uploaded_wine_file = st.file_uploader("와인 대장 파일을 업로드하세요 (.xlsx)", type=["xlsx"], key="wine_uploader")

        if uploaded_wine_file is not None:
            st.success("와인 대장 파일이 웹 브라우저에 준비되었습니다!")
            if st.button("🚀 와인 대장 파싱 및 연동 시작", use_container_width=True):
                try:
                    xl = pd.ExcelFile(uploaded_wine_file)
                    # 첫 번째 시트 혹은 '와인' 관련 시트 자동 매칭
                    target_sheet = xl.sheet_names[0]
                    for name in xl.sheet_names:
                        if "와인" in name or "재고" in name:
                            target_sheet = name
                            break
                    
                    df = pd.read_excel(xl, sheet_name=target_sheet, skiprows=2)
                    df.columns = [str(c).strip().replace(" ", "") for c in df.columns]
                    
                    name_col = None
                    for col in df.columns:
                        if '품목' in col or '와인명' in col or '이름' in col or '구분' in col:
                            name_col = col
                            break
                            
                    if name_col:
                        df = df.dropna(subset=[name_col])
                        df = df[df[name_col].astype(str).str.strip() != '0']
                        
                        stock_col = [c for c in df.columns if '재고' in c or '이월' in c]
                        expiry_col = [c for c in df.columns if '유통' in c or '빈티지' in c]
                        
                        temp_df = pd.DataFrame()
                        temp_df['품목명'] = df[name_col].astype(str).str.strip()
                        temp_df['대분류'] = "와인"
                        temp_df['유통기한'] = df[expiry_col[0]].astype(str).str.strip() if expiry_col else ""
                        temp_df['엑셀기본재고'] = pd.to_numeric(df[stock_col[0]], errors='coerce').fillna(0).astype(int) if stock_col else 0
                        
                        # 와인 마스터 파일(wine_custom_master.csv) 갱신 및 내부 서버용 원본 파일 교체
                        temp_df.to_csv(WINE_MASTER_FILE, index=False, encoding='utf-8-sig')
                        with open(WINE_EXCEL_PATH, "wb") as f:
                            f.write(uploaded_wine_file.getbuffer())
                            
                        st.balloons()
                        st.success("🎉 와인 대장 품목 및 기본 재고가 시스템에 완벽하게 연동되었습니다!")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("⚠️ 와인 파일 내에서 품목명이나 와인 이름을 식별할 수 있는 열을 찾지 못했습니다.")
                except Exception as e:
                    st.error(f"와인 대장 파싱 중 오류 발생: {e}")
                    
    st.markdown("---")
    
    # --- 초기화 기능 레이아웃 (와인 초기화 코드 싱크) ---
    st.markdown("### 🚨 기존 데이터 초기화 섹션")
    col_reset1, col_reset2 = st.columns(2)
    with col_reset1:
        st.markdown("#### 📅 당일 데이터 초기화")
        confirm_day = st.checkbox("정말로 오늘 데이터를 전부 삭제하는 것에 동의합니다.", key="confirm_day_reset")
        if st.button("🗑️ 당일 데이터 초기화 실행", type="primary", disabled=not confirm_day):
            for log_path in [STOCK_LOG_FILE, WINE_LOG_FILE]:
                if os.path.exists(log_path):
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    shutil.copyfile(log_path, f"{log_path.replace('.csv','')}_backup_{timestamp}.csv")
                    logs = pd.read_csv(log_path, encoding='utf-8-sig')
                    filtered_logs = logs[logs["날짜"] != datetime.date.today().strftime("%Y-%m-%d")]
                    filtered_logs.to_csv(log_path, index=False, encoding='utf-8-sig')
            st.session_state.success_msg = "✅ 당일 로그 초기화 완료 (일반 및 와인 일괄 적용)!"
            st.cache_data.clear()
            st.rerun()

    with col_reset2:
        st.markdown("#### 🗓️ 당월 데이터 초기화")
        confirm_month = st.checkbox("정말로 이번 달 데이터를 전부 삭제하는 것에 동의합니다.", key="confirm_month_reset")
        if st.button("💥 당월 데이터 전체 초기화 실행", type="primary", disabled=not confirm_month):
            for log_path in [STOCK_LOG_FILE, WINE_LOG_FILE]:
                if os.path.exists(log_path):
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    shutil.copyfile(log_path, f"{log_path.replace('.csv','')}_backup_{timestamp}.csv")
                    logs = pd.read_csv(log_path, encoding='utf-8-sig')
                    logs['날짜_dt'] = pd.to_datetime(logs['날짜'])
                    filtered_logs = logs[logs['날짜_dt'].dt.strftime("%Y-%m") != datetime.date.today().strftime("%Y-%m")]
                    filtered_logs = filtered_logs.drop(columns=['날짜_dt'])
                    filtered_logs.to_csv(log_path, index=False, encoding='utf-8-sig')
            st.session_state.success_msg = "✅ 당월 로그 전체 초기화 완료 (일반 및 와인 일괄 적용)!"
            st.cache_data.clear()
            st.rerun()
