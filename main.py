import streamlit as st
from io import BytesIO
from convert import convert

employee_order = [
        "강희경(Sophie)", "김민경(Ari)", "김민규(Arthur)", "김성준(Alex)",
        "김정한(Hans)", "박주헌(Stark)", "성영아(Amy)", "양은영(Ella)", "오준석(Alex)",
        "유주영(Roxie)", "정기철(Roy)", "정대웅(Henry)", "최정원(Jen)", "김지연(Joanna)", 
        "정석영(Lucas)", "제갈성규(Kai)",
        "박병건(Ben)", "서이현(Zoe)", "정재윤(Rio)"
    ]

st.title("출,퇴근 기록 파일 Converter")
st.write("업로드한 엑셀 파일의 양식을 변환한 후 다운로드할 수 있습니다.")

uploaded_file = st.file_uploader("xlsx 파일을 업로드하세요", type=["xlsx"])

if uploaded_file:
    try:
        # 임시 경로에 저장
        with open("temp_file.xlsx", "wb") as temp_file:
            temp_file.write(uploaded_file.read())
        
        # XML을 Excel로 변환
        st.write("📤 파일을 처리 중입니다...")
        workbook = convert("temp_file.xlsx", employee_order)  # convert 함수 호출

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
