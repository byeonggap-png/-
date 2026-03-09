import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO

# 1. 페이지 설정 및 제목
st.set_page_config(page_title="정문업무일지 생성기", page_icon="📝")
st.title("📝 정문업무일지 자동 생성기")
st.markdown("구글 시트의 데이터를 읽어 엑셀 양식에 자동으로 입력합니다.")

# 2. 설정 정보 (고정값)
TEMPLATE_FILE = '정문업무일지_양식_260211_v2.xlsx'

# 3. 사용자 입력 섹션
st.sidebar.header("⚙️ 설정")
sheet_url = st.sidebar.text_input("구글 시트 주소를 입력하세요", 
                                  placeholder="https://docs.google.com/spreadsheets/d/...")

# 4. 데이터 처리 함수
def process_excel(url):
    try:
        # 시트 ID 추출
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
        if not match:
            st.error("올바른 구글 시트 주소가 아닙니다.")
            return None
        
        sheet_id = match.group(1)
        
        # 엑셀 양식 불러오기
        wb = openpyxl.load_workbook(TEMPLATE_FILE)
        ws_front = wb['전면']

        # --- 데이터 읽기 및 입력 (기본정보) ---
        basic_url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=기본정보'
        df_basic = pd.read_csv(basic_url)
        
        if not df_basic.empty:
            data = df_basic.iloc[0]
            ws_front['B6'] = data.get('날짜', '')
            ws_front['F6'] = data.get('요일', '')
            ws_front['H6'] = data.get('날씨', '')
            ws_front['Y3'] = data.get('근무자(정)', '')
            ws_front['AF3'] = data.get('근무자(부)', '')

        # --- 데이터 읽기 및 입력 (출입사항) ---
        entry_url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=출입사항'
        df_entry = pd.read_csv(entry_url)
        
        if not df_entry.empty:
            start_row = 10
            for i, row in df_entry.iterrows():
                curr = start_row + i
                ws_front[f'B{curr}'] = row.get('성명', '')
                ws_front[f'E{curr}'] = row.get('입소시간', '')
                ws_front[f'H{curr}'] = row.get('퇴소시간', '')
                ws_front[f'K{curr}'] = row.get('목적지', '')
                ws_front[f'R{curr}'] = row.get('사유', '')

        # 메모리에 파일 저장
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
        return None

# 5. 실행 버튼
if st.button("🚀 업무일지 생성하기"):
    if not sheet_url:
        st.warning("먼저 구글 시트 주소를 입력해주세요.")
    else:
        with st.spinner('데이터를 가져와서 엑셀을 만드는 중입니다...'):
            result_file = process_excel(sheet_url)
            
            if result_file:
                st.success("✅ 일지가 성공적으로 생성되었습니다!")
                st.download_button(
                    label="📥 완성된 업무일지 다운로드",
                    data=result_file,
                    file_name="오늘의_정문업무일지.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
