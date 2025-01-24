import streamlit as st
from openpyxl import Workbook
from io import BytesIO
from convert import convert

st.title("출,퇴근 기록 파일 Converter")
st.write("업로드한 XML 파일을 엑셀 파일로 변환한 후 다운로드할 수 있습니다.")

uploaded_file = st.file_uploader("XML 파일을 업로드하세요", type=["xml"])

if uploaded_file:
    try:
        # 임시 경로에 저장
        with open("temp_file.xml", "wb") as temp_file:
            temp_file.write(uploaded_file.read())
        
        # XML을 Excel로 변환
        st.write("📤 XML 파일을 처리 중입니다...")
        workbook = convert("temp_file.xml")  # convert 함수 호출

        # 처리 완료 메시지
        st.success("✅ 처리 완료되었습니다! 아래에서 엑셀 파일을 다운로드하세요.")
        
        # Workbook을 BytesIO 객체로 저장
        output = BytesIO()
        workbook.save(output)
        output.seek(0)

        # 다운로드 버튼
        st.download_button(
            label="엑셀 파일 다운로드",
            data=output,
            file_name="converted_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"파일 처리 중 오류가 발생했습니다: {e}")
