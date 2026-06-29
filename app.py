import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import os
from datetime import datetime

# ==========================================
# 1. 환경 설정 및 초기화
# ==========================================
ORIGINAL_EXCEL_PATH = "VINI_COFFEE_통합_식자재_및_매출관리_시스템_v3_주간체크리스트추가.xlsx"
STOCK_LOG_FILE = "vini_daily_stock_log.csv"

# 페이지 레이아웃 설정
st.set_page_config(page_title="VINI COFFEE 재고관리 시스템", layout="wide", page_icon="☕")

# 임시 히스토리용 CSV 로그 파일이 없을 경우 초기화
if not os.path.exists(STOCK_LOG_FILE):
    df_empty = pd.DataFrame(columns=["날짜", "대분류", "품목명", "구분", "수량", "유통기한"])
    df_empty.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')

# 세션 상태(Session State) 메시지 및 캐시 제어 초기화
if "success_msg" not in st.session_state:
    st.session_state.success_msg = ""

# 실제 엑셀 파일 내 시트 이름 매핑 규칙 정의
SHEET_MAP = {
    "원재료": "원재료(간략)",
    "부자재": "부자재(간략)",
    "디저트&완제품": "디저트&완제품(간략)"
}


# ==========================================
# 2. 핵심 로직 함수 (서버 원본 엑셀 핸들링)
# ==========================================

def update_single_excel(cat, item_name, qty, new_expiry=None, action="inbound"):
    """[기능 1] 개별 단건 입출고 발생 시 서버의 원본 엑셀을 실시간으로 수정하는 함수"""
    if not os.path.exists(ORIGINAL_EXCEL_PATH):
        st.error(f"❌ 서버에서 원본 엑셀 파일을 찾을 수 없습니다: {ORIGINAL_EXCEL_PATH}")
        return False

    try:
        # openpyxl로 서버 파일 로드
        wb = openpyxl.load_workbook(ORIGINAL_EXCEL_PATH)
        target_sheet = SHEET_MAP.get(cat, cat)
        
        if target_sheet not in wb.sheetnames:
            st.error(f"❌ 엑셀 파일 내에 '{target_sheet}' 시트가 존재하지 않습니다.")
            return False
            
        ws = wb[target_sheet]
        header_row = 3  # skiprows=2 (3번째 행이 헤더 명칭 라인)
        
        # 동적 칼럼 인덱스 매핑 파싱
        name_col_idx, stock_col_idx, expiry_col_idx = None, None, None
        for col in range(1, ws.max_column + 1):
            val = str(ws.cell(row=header_row, column=col).value).strip().replace(" ", "")
            if '품목명' in val or '품목이름' in val or '구분' in val:
                name_col_idx = col
            elif '재고' in val:
                stock_col_idx = col
            elif '유통' in val:
                expiry_col_idx = col

        if not name_col_idx or not stock_col_idx:
            st.error("❌ 엑셀 헤더 열 구조를 파싱할 수 없습니다. ('품목명' 또는 '재고' 열 확인 필요)")
            return False

        # 행단위 품목 탐색 후 데이터 수정
        item_found = False
        for row in range(header_row + 1, ws.max_row + 1):
            cell_item_name = str(ws.cell(row=row, column=name_col_idx).value).strip()
            
            if cell_item_name == item_name:
                current_stock_cell = ws.cell(row=row, column=stock_col_idx)
                
                # 기존 재고 수치 안전하게 정수 변환
                try:
                    current_stock = int(current_stock_cell.value) if current_stock_cell.value is not None else 0
                except:
                    current_stock = 0

                # 입고(+) 및 출고(-) 수량 누적 연산 처리
                if action == "inbound":
                    current_stock_cell.value = current_stock + qty
                elif action == "outbound":
                    current_stock_cell.value = max(0, current_stock - qty) # 음수 방지

                # 유통기한 데이터 실시간 반영
                if new_expiry and expiry_col_idx:
                    ws.cell(row=row, column=expiry_col_idx).value = new_expiry
                
                item_found = True
                break

        if not item_found:
            st.error(f"❌ '{item_name}' 품목을 엑셀 파일 안에서 찾을 수 없습니다.")
            return False

        # 변경 완료 후 파일 잠금 해제 및 저장
        wb.save(ORIGINAL_EXCEL_PATH)
        wb.close()
        return True
    except Exception as e:
        st.error(f"❌ 엑셀 실시간 수정 처리 중 오류가 발생했습니다: {e}")
        return False


# ==========================================
# 3. 콜백 함수 (단건 버튼 이벤트 제어)
# ==========================================

def save_inbound_callback():
    """개별 단건 입고 등록 실행 프로세스"""
    qty = st.session_state.in_qty
    if qty > 0:
        cat = st.session_state.in_cat
        item = st.session_state.in_item
        expiry_str = st.session_state.in_utg.strftime("%Y-%m-%d")
        
        # 1단계: 원본 엑셀 실시간 동기화 업데이트
        if update_single_excel(cat, item, qty, new_expiry=expiry_str, action="inbound"):
            # 2단계: 히스토리 백업용 CSV 로그 누적
            current_logs = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
            new_data = pd.DataFrame([{
                "날짜": st.session_state.in_date.strftime("%Y-%m-%d"),
                "대분류": cat,
                "품목명": item,
                "구분": "금월 입고",
                "수량": qty,
                "유통기한": expiry_str
            }])
            current_logs = pd.concat([current_logs, new_data], ignore_index=True)
            current_logs.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')
            
            st.session_state.success_msg = f"📥 [엑셀 반영 완료] 입고: {item} | {qty}개가 마스터 파일에 가산되었습니다."
            st.session_state.in_qty = 0  # 입력 컴포넌트 초기화
            st.cache_data.clear()        # 화면 동기화를 위해 메모리 캐시 삭제
    else:
        st.warning("⚠️ 입고 수량을 최소 1개 이상 입력하셔야 합니다.")

def save_outbound_callback():
    """개별 단건 출고 등록 실행 프로세스"""
    qty = st.session_state.out_qty
    if qty > 0:
        cat = st.session_state.out_cat
        item = st.session_state.out_item
        
        # 1단계: 원본 엑셀 실시간 동기화 업데이트
        if update_single_excel(cat, item, qty, action="outbound"):
            # 2단계: 히스토리 백업용 CSV 로그 누적
            current_logs = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
            new_data = pd.DataFrame([{
                "날짜": st.session_state.out_date.strftime("%Y-%m-%d"),
                "대분류": cat,
                "품목명": item,
                "구분": "금월 출고",
                "수량": qty,
                "유통기한": "-"
            }])
            current_logs = pd.concat([current_logs, new_data], ignore_index=True)
            current_logs.to_csv(STOCK_LOG_FILE, index=False, encoding='utf-8-sig')
            
            st.session_state.success_msg = f"📤 [엑셀 반영 완료] 출고: {item} | {qty}개가 마스터 파일에서 차감되었습니다."
            st.session_state.out_qty = 0  # 입력 컴포넌트 초기화
            st.cache_data.clear()         # 화면 동기화를 위해 메모리 캐시 삭제
    else:
        st.warning("⚠️ 출고 수량을 최소 1개 이상 입력하셔야 합니다.")


# ==========================================
# 4. 메인 UI 레이아웃 구성
# ==========================================

st.title("☕ VINI COFFEE 통합 식자재 및 마감 관리 시스템")

# 상단 실시간 알림창 출력 구역
if st.session_state.success_msg:
    st.success(st.session_state.success_msg)
    st.session_state.success_msg = "" # 메시지 큐 초기화

# 인터페이스 기능별 탭 분할
tab1, tab2, tab3 = st.tabs([
    "✨ 실시간 현황판 및 개별 등록", 
    "📊 전품목 일괄 조정 & 즉시 다운로드", 
    "📜 단건 입출고 히스토리 대장"
])

# ------------------------------------------
# TAB 1: 실시간 마스터 엑셀 현황 조회 및 개별 건수 조작
# ------------------------------------------
with tab1:
    st.header("🔍 원본 엑셀 실시간 현황판")
    
    if os.path.exists(ORIGINAL_EXCEL_PATH):
        # 상단 라디오 버튼으로 시트 전환
        view_cat = st.radio("카테고리를 전환하며 재고를 확인하세요", ["원재료", "부자재", "디저트&완제품"], horizontal=True)
        target_sheet = SHEET_MAP[view_cat]
        
        # 현재 서버 원본 데이터 실시간 로드 및 미사용 열 전처리
        df_view = pd.read_excel(ORIGINAL_EXCEL_PATH, sheet_name=target_sheet, skiprows=2)
        df_view = df_view.loc[:, ~df_view.columns.str.contains('^Unnamed')]
        st.dataframe(df_view, use_container_width=True, hide_index=True)
        
        # 셀렉트박스용 유효 품목 리스트 바인딩
        name_col = "품목명" if "품목명" in df_view.columns else (df_view.columns[1] if len(df_view.columns)>1 else df_view.columns[0])
        item_list = df_view[name_col].dropna().unique().tolist()
        
        # 단건 등록 레이아웃 (2열 종대 분할)
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("---")
            st.markdown("### 📥 단건 실시간 입고 등록")
            st.date_input("입고 일자", datetime.now(), key="in_date")
            st.selectbox("대분류 구분", ["원재료", "부자재", "디저트&완제품"], key="in_cat")
            st.selectbox("등록할 품목명 선택", item_list, key="in_item")
            st.number_input("입고 수량(EA)", min_value=0, step=1, key="in_qty")
            st.date_input("유통기한 기입", datetime.now(), key="in_utg")
            st.button("📥 입고 데이터 엑셀 전송", on_click=save_inbound_callback)

        with col2:
            st.markdown("---")
            st.markdown("### 📤 단건 실시간 출고 등록")
            st.date_input("출고 일자", datetime.now(), key="out_date")
            st.selectbox("대분류 구분", ["원재료", "부자재", "디저트&완제품"], key="out_cat")
            st.selectbox("등록할 품목명 선택", item_list, key="out_item")
            st.number_input("출고 수량(EA)", min_value=0, step=1, key="out_qty")
            st.button("📤 출고 데이터 엑셀 전송", on_click=save_outbound_callback)
    else:
        st.error(f"서버 내부 경로에 원본 마스터 엑셀 파일이 존재하지 않습니다: {ORIGINAL_EXCEL_PATH}")

# ------------------------------------------
# TAB 2: 전품목 한 눈에 일괄 편집 후 마스터 엑셀 즉시 추출
# ------------------------------------------
with tab2:
    st.header("📊 전품목 일괄 마감 처리 및 원본 파일 다운로드")
    
    bulk_cat = st.selectbox("일괄 편집 및 다운로드할 시트 선택", ["원재료", "부자재", "디저트&완제품"], key="bulk_cat")
    bulk_sheet = SHEET_MAP[bulk_cat]
    
    if os.path.exists(ORIGINAL_EXCEL_PATH):
        try:
            # 원본 데이터 로드
            df_bulk = pd.read_excel(ORIGINAL_EXCEL_PATH, sheet_name=bulk_sheet, skiprows=2)
            df_bulk = df_bulk.loc[:, ~df_bulk.columns.str.contains('^Unnamed')]
            
            st.info("💡 **일괄 작업 프로세스**\n1. 테이블 내부의 **[금월 입고]** 또는 **[금월 출고]** 입력란을 마우스로 더블클릭하여 오늘 변경된 모든 수량을 채워 넣습니다.\n2. 하단의 **[변경사항 원본 엑셀에 일괄 반영하기]** 버튼을 클릭하여 서버 파일을 저장합니다.\n3. 저장이 끝나면 그 자리에 나타나는 **[다운로드]** 링크를 클릭해 완성된 파일을 로컬 PC로 가져갑니다.")
            
            # 사고 방지용 락(Lock) 설계: 입고, 출고, 유통기한 제외 품목 고유 명세 정보 수정 불가능 고정
            disable_cols = [c for c in df_bulk.columns if "입고" not in c and "출고" not in c and "유통" not in c]
            
            # Streamlit 그리드 편집 인터페이스 오픈
            edited_df = st.data_editor(
                df_bulk,
                use_container_width=True,
                hide_index=True,
                disabled=disable_cols
            )
            
            # 서버 마스터 파일 오버라이트 프로세스 실행 버튼
            if st.button("💾 변경사항 원본 엑셀에 일괄 반영하기", type="primary", key="bulk_save_btn"):
                with st.spinner("서버의 마스터 엑셀 시트에 데이터 일괄 수산 작업 진행 중..."):
                    wb = openpyxl.load_workbook(ORIGINAL_EXCEL_PATH)
                    ws = wb[bulk_sheet]
                    header_row = 3
                    
                    # 칼럼 딕셔너리 빌드
                    col_map = {str(ws.cell(row=header_row, column=c).value).strip().replace(" ", ""): c for c in range(1, ws.max_column + 1)}
                    
                    name_idx = col_map.get("품목명") or col_map.get("품목이름") or col_map.get("구분")
                    inbound_idx = col_map.get("금월입고") or col_map.get("입고")
                    outbound_idx = col_map.get("금월출고") or col_map.get("출고")
                    expiry_idx = col_map.get("유통기한") or col_map.get("유통")
                    
                    if not name_idx:
                        st.error("❌ 엑셀 구조가 기준 규격과 상이하여 '품목명' 열을 찾지 못했습니다.")
                    else:
                        # 사용자가 수정한 데이터프레임의 모든 행을 순회하며 openpyxl 엔진으로 강제 주입
                        for _, row in edited_df.iterrows():
                            # 데이터프레임 상의 품목명 탐색
                            grid_item_name = str(row.get(df_bulk.columns[df_bulk.columns.isin(["품목명","품목이름","구분"])][0])).strip()
                            
                            for r in range(header_row + 1, ws.max_row + 1):
                                excel_item_name = str(ws.cell(row=r, column=name_idx).value).strip()
                                
                                if excel_item_name == grid_item_name:
                                    # 예외 필터링 (결측치 데이터 0 초기화)
                                    in_val = row.get("금월 입고") or row.get("입고") or 0
                                    out_val = row.get("금월 출고") or row.get("출고") or 0
                                    exp_val = row.get("유통기한") or row.get("유통") or ""
                                    
                                    if inbound_idx: ws.cell(row=r, column=inbound_idx).value = int(in_val)
                                    if outbound_idx: ws.cell(row=r, column=outbound_idx).value = int(out_val)
                                    if expiry_idx and exp_val: ws.cell(row=r, column=expiry_idx).value = str(exp_val)
                                    break
                        
                        # 엑셀 파일 저장 및 연결 종료
                        wb.save(ORIGINAL_EXCEL_PATH)
                        wb.close()
                        st.cache_data.clear() # 수정 사항 적용을 위한 캐시 만료 처리
                        
                        st.success("🎉 원본 엑셀 파일이 성공적으로 업데이트되었습니다! 이제 아래 버튼을 눌러 소장용 파일로 받아 가세요.")
                        
                        # 방금 저장된 따끈따끈한 최신 파일을 바이너리로 로드
                        with open(ORIGINAL_EXCEL_PATH, "rb") as f:
                            file_bytes = f.read()
                        
                        # 직관적인 원클릭 다운로드 단추 노출
                        st.download_button(
                            label="📥 즉시 동기화 완료된 원본 엑셀 파일 다운로드 (.xlsx)",
                            data=file_bytes,
                            file_name=f"🟢최신반영_VINI_식자재매출관리_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
        except Exception as e:
            st.error(f"❌ 일괄 편집 스트리밍 처리 오류: {e}")
    else:
        st.error(f"서버 내부 경로에 원본 마스터 엑셀 파일이 존재하지 않습니다: {ORIGINAL_EXCEL_PATH}")

# ------------------------------------------
# TAB 3: 백업 히스토리용 간이 영수증 CSV 로그 뷰어
# ------------------------------------------
with tab3:
    st.header("📜 실시간 단건 입출고 이력 데이터")
    st.caption("※ 이 표는 시스템 로그 컴포넌트이며, 원본 마스터 엑셀 내용과는 별개의 상세 내역 히스토리입니다.")
    if os.path.exists(STOCK_LOG_FILE):
        df_logs = pd.read_csv(STOCK_LOG_FILE, encoding='utf-8-sig')
        st.dataframe(df_logs.sort_index(ascending=False), use_container_width=True, hide_index=True)
    else:
        st.info("기록된 입출고 히스토리가 존재하지 않습니다.")
