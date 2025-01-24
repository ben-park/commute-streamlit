import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO

# XML 파일을 DataFrame으로 변환하는 함수
def parse_xml_to_dataframe(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # XML 구조를 반복적으로 파싱하여 데이터 저장
    data = []
    for child in root:
        row = {}
        for element in child:
            row[element.tag] = element.text
        data.append(row)
    
    # DataFrame 생성
    df = pd.DataFrame(data)
    return df

# DataFrame을 Excel 파일로 저장하는 함수
def dataframe_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    processed_data = output.getvalue()
    return processed_data

# Streamlit UI 구성
st.title("XML to Excel Converter")
st.write("업로드한 XML 파일을 엑셀 파일로 변환한 후 다운로드할 수 있습니다.")

# 파일 업로드
uploaded_file = st.file_uploader("XML 파일을 업로드하세요", type=["xml"])

if uploaded_file:
    try:
        # XML 파일을 DataFrame으로 변환
        st.write("📤 XML 파일을 처리 중입니다...")
        dataframe = parse_xml_to_dataframe(uploaded_file)

        # DataFrame 미리보기
        st.write("📋 변환된 데이터 미리보기:")
        st.dataframe(dataframe)

        # 엑셀 파일 변환
        excel_file = dataframe_to_excel(dataframe)
        
        # 다운로드 버튼
        st.download_button(
            label="엑셀 파일 다운로드",
            data=excel_file,
            file_name="converted_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"파일 처리 중 오류가 발생했습니다: {e}")

