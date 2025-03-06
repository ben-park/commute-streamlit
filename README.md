# 출퇴근 기록 파일 변환기

KT Telecop 프로그램에서 추출한 엑셀 출퇴근 기록 파일을 사내 업무용 양식으로 변환합니다.

웹페이지를 통해 원본 파일을 업로드 하고 변경된 파일을 다운로드 받을 수 있습니다.

## 주요 기능

- **엑셀 파일 업로드 및 변환**: 사용자가 업로드한 엑셀 파일을 분석하여 출퇴근 시간을 추출하고, 지정된 양식으로 변환.
- **시간 변환**: 오후 시간을 24시간 형식으로 변환하여 데이터의 일관성을 유지.
- **날짜 및 요일 자동 설정**: 날짜별로 출퇴근 시간을 정리하고, 요일을 자동으로 계산하여 표시.
- **주말 제외**: 토요일과 일요일은 결과에서 제외하여 평일 출퇴근 기록만 표시.
- **지각 표시**: 오전 8시 30분 이후 출근 시 해당 셀을 강조하여 지각 여부를 쉽게 확인 가능.
- **파일 다운로드**: 변환된 엑셀 파일을 사용자가 다운로드할 수 있도록 제공.
- **웹 기반 인터페이스**: Streamlit을 사용하여 웹 브라우저에서 편리하게 접근하고 사용 가능.

## 환경 설정

- **Python 3.7+**
- **필요 라이브러리 설치**:

```bash
pip install pandas openpyxl streamlit
```

## 실행 방법

1.  **프로그램 실행**:

```bash
streamlit run main.py
```

위 명령으로 실행 후 http://localhost:8501 접속

2.  **엑셀 파일 업로드**: 웹 브라우저에 표시된 인터페이스에서 "xlsx 파일을 업로드하세요" 버튼을 클릭하고, 변환할 엑셀 파일을 선택합니다.
3.  **파일 처리 및 다운로드**: 파일이 업로드되고 변환이 완료되면, "엑셀 파일 다운로드" 버튼이 활성화됩니다. 이 버튼을 클릭하여 변환된 엑셀 파일을 다운로드합니다.

## 추가 정보

- `convert.py`에서 데이터 변환 및 엑셀 파일 생성을 담당하고, `main.py`에서 Streamlit을 사용하여 인터페이스를 제공합니다.
- 변환된 엑셀 파일은 `converted_file.xlsx`라는 이름으로 저장됩니다.
- 오류 발생 시 웹 인터페이스에 오류 메시지가 표시됩니다.
