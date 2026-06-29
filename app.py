import streamlit as st
import pandas as pd
import openpyxl
import os

# 원본 엑셀 파일 경로 설정 (본인 환경에 맞게 조정)
ORIGINAL_EXCEL_PATH = "VINI_COFFEE_통합_식자재_및_매출관리_시스템_v3_주간체크리스트추가.xlsx"

st.subheader("📊 전품목 일괄 재고 조정 및 다운로드")

# 1. 작업할 대분류 선택
selected_cat = st.selectbox(
    "일괄 수정할 대분류를 선택하세요", 
    ["원재료", "부자재", "디저트&완제품"], 
    key="bulk_cat_select"
)

# 엑셀 시트명 매핑 (기존 시스템 규칙 반영)
sheet_map = {
    "원재료": "원재료(간략)",
    "부자재": "부자재(간략)",
    "디저트&완제품": "디저트&완제품(간략)"
}
target_sheet = sheet_map[selected_cat]

if os.path.exists(ORIGINAL_EXCEL_PATH):
    # 2. 원본 엑셀에서 현재 데이터 읽어오기 (헤더 위치 skiprows=2 가정)
    try:
        df_current = pd.read_excel(ORIGINAL_EXCEL_PATH, sheet_name=target_sheet, skiprows=2)
        
        # 공백 칼럼이나 이름 없는 칼럼 정제
        df_current = df_current.loc[:, ~df_current.columns.str.contains('^Unnamed')]
        
        st.write(f"👉 아래 테이블에서 **[금월 입고]** 또는 **[금월 출고]** 컬럼의 수량을 직접 수정하세요.")
        st.caption("팁: 셀을 더블클릭하면 숫자를 입력할 수 있습니다.")
        
        # 3. Streamlit Data Editor를 통한 사용자 입력 받기
        # '품목명', '현재재고' 등은 수정 못하게 잠그고, 입고/출고 컬럼만 활성화
        edited_df = st.data_editor(
            df_current,
            use_container_width=True,
            hide_index=True,
            disabled=["바코드", "코드", "품목명", "품목이름", "구분", "현재재고", "이월재고"] # 수정 금지할 칼럼들
        )
        
        # 4. 저장 및 동기화 버튼
        if st.button("💾 변경사항 원본 엑셀에 일괄 반영하기", type="primary"):
            with st.spinner("서버의 엑셀 파일을 업데이트하고 다운로드 링크를 생성 중입니다..."):
                try:
                    # openpyxl로 원본 엑셀 로드
                    wb = openpyxl.load_workbook(ORIGINAL_EXCEL_PATH)
                    ws = wb[target_sheet]
                    
                    # 엑셀의 칼럼 매핑 분석 (3번째 행이 헤더인 경우)
                    header_row = 3
                    col_map = {}
                    for col in range(1, ws.max_column + 1):
                        val = str(ws.cell(row=header_row, column=col).value).strip().replace(" ", "")
                        col_map[val] = col
                    
                    # 필수 칼럼 인덱스 확인
                    name_idx = col_map.get("품목명") or col_map.get("품목이름") or col_map.get("구분")
                    stock_idx = col_map.get("현재재고") or col_map.get("재고")
                    inbound_idx = col_map.get("금월입고") or col_map.get("입고")
                    outbound_idx = col_map.get("금월출고") or col_map.get("출고")
                    
                    if not name_idx or not stock_idx:
                        st.error("엑셀 파일에서 '품목명' 또는 '재고' 칼럼 위치를 찾지 못했습니다.")
                    else:
                        # 화면에서 편집된 데이터프레임을 한 행씩 돌며 엑셀 파일에 대입
                        for _, row in edited_df.iterrows():
                            item_name = str(row.get("품목명") or row.get("품목이름") or row.get("구분")).strip()
                            
                            # 엑셀에서 동일한 품목명을 가진 행 찾기
                            for r in range(header_row + 1, ws.max_row + 1):
                                excel_item_name = str(ws.cell(row=r, column=name_idx).value).strip()
                                
                                if excel_item_name == item_name:
                                    # 사용자가 입력한 입고, 출고 값 반영 (비어있으면 0)
                                    in_val = row.get("금월 입고") or row.get("입고") or 0
                                    out_val = row.get("금월 출고") or row.get("출고") or 0
                                    
                                    # 엑셀 셀에 직접 값 주입
                                    if inbound_idx: ws.cell(row=r, column=inbound_idx).value = int(in_val)
                                    if outbound_idx: ws.cell(row=r, column=outbound_idx).value = int(out_val)
                                    
                                    # 계산 공식이 아니라 값으로 연산할 경우 재고 최신화 (옵션)
                                    # (만약 엑셀 내부에 =이월+입고-출고 수식이 걸려있다면 이 단계는 건너뛰어도 됩니다)
                                    break
                        
                        # 변경된 파일 저장
                        wb.save(ORIGINAL_EXCEL_PATH)
                        wb.close()
                        
                        # 캐시 초기화하여 화면 리프레시 준비
                        st.cache_data.clear()
                        st.success("✅ 서버 원본 엑셀 파일 수정이 안전하게 완료되었습니다!")
                        
                        # 5. 수정한 엑셀 파일을 읽어서 바로 다운로드 버튼으로 제공
                        with open(ORIGINAL_EXCEL_PATH, "rb") as f:
                            excel_data = f.read()
                        
                        st.download_button(
                            label="📥 변경된 통합 엑셀 파일 즉시 다운로드",
                            data=excel_data,
                            file_name=f"🟢최신반영_{ORIGINAL_EXCEL_PATH}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                except Exception as e:
                    st.error(f"엑셀 동기화/파일 작업 중 에러 발생: {e}")
                    
    except Exception as e:
        st.error(f"데이터를 읽어오는 중 에러가 발생했습니다: {e}")
else:
    st.error(f"지정된 경로에서 원본 엑셀 파일을 찾을 수 없습니다: {ORIGINAL_EXCEL_PATH}")
