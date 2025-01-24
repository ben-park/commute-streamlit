import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from operator import itemgetter
import xml.etree.ElementTree as ET
import json

def get_all_dates(dict):
    sorted_dates = sorted(dict.keys(), key=lambda x: datetime.strptime(x, '%Y-%m-%d'))
    min_date = datetime.strptime(sorted_dates[0], '%Y-%m-%d')
    max_date = datetime.strptime(sorted_dates[-1], '%Y-%m-%d')
    # 최소일과 최대일 사이 모든 날짜 생성
    return [(min_date + timedelta(days=i)).strftime('%Y-%m-%d') for i in range((max_date - min_date).days + 1)]

def xml_to_json(file_path):
    try:
        # XML 파일 파싱
        tree = ET.parse(file_path)
        root = tree.getroot()

        # XML 데이터를 딕셔너리로 변환
        def parse_element(element):
            parsed_data = {}
            
            # 하위 요소가 없으면 텍스트 값을 반환
            if len(element) == 0:
                return element.text.strip() if element.text else None

            # 하위 요소가 있으면 재귀적으로 처리
            for child in element:
                child_data = parse_element(child)
                if child.tag in parsed_data:
                    # 같은 태그 이름이 여러 개 있을 경우 리스트로 저장
                    if not isinstance(parsed_data[child.tag], list):
                        parsed_data[child.tag] = [parsed_data[child.tag]]
                    parsed_data[child.tag].append(child_data)
                else:
                    parsed_data[child.tag] = child_data

            return parsed_data

        # 루트 요소부터 파싱 시작
        json_data = {root.tag: parse_element(root)}

        return json_data

    except ET.ParseError as e:
        print(f"XML 파싱 오류: {e}")
        return None
    except Exception as e:
        print(f"오류 발생: {e}")
        return None

# 사용 예시
if __name__ == "__main__":
    xml_file_path = "example.xml"  # XML 파일 경로
    json_data = xml_to_json(xml_file_path)

    if json_data:
        # JSON 데이터 출력
        print(json.dumps(json_data, indent=4, ensure_ascii=False))


def convert(file_path):
    # 데이터 변환
    data = xml_to_json(file_path)
    work_data = data["NewDataSet"]["tb_workresult"]

    # 직원 이름과 출근/퇴근 시간을 추출
    attendance_dict = {}

    # 날짜별 직원 출근 및 퇴근 정보 분리
    for entry in work_data:
        datetime_str, mode, name = itemgetter('date_Attestation', 'str_Mode', 'str_workempName')(entry)
        date, time = datetime_str.split(' ')[0], datetime_str.split(' ')[2]

        # 날짜별 직원 출근/퇴근 정보 저장
        if date not in attendance_dict:
            attendance_dict[date] = {}

        if name not in attendance_dict[date]:
            attendance_dict[date][name] = {'출근': '', '퇴근': ''}

        # 출근 또는 퇴근 기록 추가
        attendance_dict[date][name][mode] = time

    # 날짜와 직원 정보를 기반으로 데이터프레임 생성
    rows = []

    # 직원들의 이름 목록을 추출 (중복 없는 이름)
    employee_names = sorted(set([entry['str_workempName'] for entry in work_data]))

    # 모든 날짜 추출
    all_dates = list(attendance_dict.keys())

    # 각 날짜에 대해 직원들의 출근/퇴근 시간을 행으로 추가
    for date in all_dates:
        # 날짜 및 요일 계산
        weekday = datetime.strptime(date, '%Y-%m-%d').strftime('%A')

        if weekday == 'Saturday' or weekday == 'Sunday':
            continue

        weekday_kr = {'Monday': '월', 'Tuesday': '화', 'Wednesday': '수',
                      'Thursday': '목', 'Friday': '금'}[weekday]

        # 행 생성
        row = [datetime.strptime(date, '%Y-%m-%d').strftime('%d/%m/%Y'), weekday_kr, '']

        # 직원별 출근/퇴근 시간 추가
        if date in attendance_dict:
            for employee in employee_names:
                if employee in attendance_dict[date]:
                    row.append(attendance_dict[date][employee]['출근'])
                    row.append(attendance_dict[date][employee]['퇴근'])
                else:
                    row.append('')
                    row.append('')

        rows.append(row)

    df = pd.DataFrame(rows)

    # 엑셀 포맷 설정
    workbook = Workbook()
    sheet = workbook.active

    # 날짜, 요일 셀 설정
    sheet['A1'] = '날짜'
    sheet['B1'] = '요일'
    sheet['C1'] = '비고'
    sheet.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    sheet.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
    sheet.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)

    # 직원 이름 셀 2칸씩 병합
    for idx, col in enumerate(employee_names):
        start_column = 4 + idx * 2
        sheet.merge_cells(start_row=1, start_column=start_column, end_row=1, end_column=start_column + 1)
        sheet[get_column_letter(start_column) + '1'] = col

    # 직원 이름 아래 출근, 퇴근 셀 설정
    for idx, employee in enumerate(employee_names, 2):
        sheet[get_column_letter(idx * 2) + '2'] = '출근'
        sheet[get_column_letter(idx * 2 + 1) + '2'] = '퇴근'

    # 출퇴근 시간 데이터프레임 셀 삽입
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=3):
        for c_idx, value in enumerate(row, 1):
            sheet.cell(row=r_idx, column=c_idx, value=value)

    # 전체 스타일 적용
    # 폰트크기 및 가운데정렬
    for cells in sheet.rows:
        for cell in cells:
            cell.font = Font(name='맑은 고딕', size=12)
            cell.alignment = Alignment(horizontal='center', vertical='center')
    # 셀 너비
    sheet.column_dimensions['A'].width = 16
    sheet.column_dimensions['B'].width = 5
    sheet.column_dimensions['C'].width = 30
    for col_idx in range(4, sheet.max_column + 1):
        column_letter = get_column_letter(col_idx)
        sheet.column_dimensions[column_letter].width = 10.5

    # 헤더 고정
    sheet.freeze_panes = sheet['C3']

    # B열이 월요일인 행 위쪽 테두리 추가
    for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row, max_col=sheet.max_column):
        if row[1].value == '월':
            for cell in row:
                cell.border = Border(top=Side(border_style='double'))

    # 지각 처리 (08:30 이후 출근 시 강조 표시)
    late_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    late_font = Font(color='FF0000')

    for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row, min_col=4, max_col=sheet.max_column):
        for idx, cell in enumerate(row):
            if idx % 2 == 0 and cell.value:
                try:
                    if datetime.strptime(cell.value, '%H:%M:%S') > datetime.strptime('08:30:00', '%H:%M:%S'):
                        cell.fill = late_fill
                        cell.font = late_font
                except ValueError:
                    pass

    return workbook
