import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from operator import itemgetter
import json
import sys

def add_12_hours(time_str):
    # 문자열을 datetime 객체로 변환
    time_obj = datetime.strptime(time_str, "%H:%M:%S")
    # 12시간 추가
    new_time_obj = time_obj + timedelta(hours=12)
    # 새로운 시간 문자열 반환
    return new_time_obj.strftime("%H:%M:%S")

def get_all_dates(dict):
    sorted_dates = sorted(dict.keys(), key=lambda x: datetime.strptime(x, '%Y-%m-%d'))
    min_date = datetime.strptime(sorted_dates[0], '%Y-%m-%d')
    max_date = datetime.strptime(sorted_dates[-1], '%Y-%m-%d')
    # 최소일과 최대일 사이 모든 날짜 생성
    return [(min_date + timedelta(days=i)).strftime('%Y-%m-%d') for i in range((max_date - min_date).days + 1)]
    
def xlsx_to_json(file_path):
    excel_df = pd.read_excel(file_path, skiprows=2)
    rename_df = excel_df.rename(columns={"인증일시": "date_Attestation", "요일": "str_Week", "인증번호": "str_tmId", "사원번호": "str_workempNum", "이름": "str_workempName", "리더기 장소": "str_accTerminalPlace", "인증모드": "str_Mode", "인증상태": "str_ValidationStatus", "인증방법": "str_Certificate", "부서": "str_workempPostName", "직위": "str_workempPositionName", "타임테이블": "str_workUserTimetableName", "직원상태": "str_emptmAdmin"})
    json_data = rename_df.to_json(orient='records', force_ascii=False)
    print(json_data)
    return json_data

def convert(file_path):
    # 데이터 변환
    json_data = xlsx_to_json(file_path)
    work_data = json.loads(json_data)
    # 직원 이름과 출근/퇴근 시간을 추출
    attendance_dict = {}
    # 날짜별 직원 출근 및 퇴근 정보 분리
    for entry in work_data:
        datetime_str, mode, name = itemgetter('date_Attestation', 'str_Mode', 'str_workempName')(entry)
        try:
            parts = datetime_str.split(' ')
            if len(parts) == 3:
                date, ampm, time = parts
                if ampm == '오후':
                    time = add_12_hours(time)
            elif len(parts) == 2:
                date, time = parts
            else:
                print(f"Skipping invalid datetime format: {datetime_str}")
                continue
        except ValueError as e:
            print(f"Error parsing datetime: {e}, skipping entry: {datetime_str}")
            continue

        # 날짜별 직원 출근/퇴근 정보 저장
        if date not in attendance_dict:
            attendance_dict[date] = {}

        if name not in attendance_dict[date]:
            attendance_dict[date][name] = {'출근': '', '퇴근': '', 'times': [], '출근_추정': False, '퇴근_추정': False}

        # 모든 시간 기록 저장 및 출근/퇴근 기록 추가
        attendance_dict[date][name]['times'].append(time)
        if mode in ['출근', '퇴근']:
            current_time = attendance_dict[date][name][mode]
            if mode == '출근' and (not current_time or datetime.strptime(time, '%H:%M:%S') < datetime.strptime(current_time, '%H:%M:%S')):
                attendance_dict[date][name][mode] = time
            elif mode == '퇴근' and (not current_time or datetime.strptime(time, '%H:%M:%S') > datetime.strptime(current_time, '%H:%M:%S')):
                attendance_dict[date][name][mode] = time

    # 누락된 출근/퇴근을 추정값으로 채우기
    for date in attendance_dict:
        for name in attendance_dict[date]:
            if attendance_dict[date][name]['times']:
                times = sorted(attendance_dict[date][name]['times'])
                if not attendance_dict[date][name]['출근']:
                    attendance_dict[date][name]['출근'] = times[0]
                    attendance_dict[date][name]['출근_추정'] = True
                if not attendance_dict[date][name]['퇴근']:
                    attendance_dict[date][name]['퇴근'] = times[-1]
                    attendance_dict[date][name]['퇴근_추정'] = True

    # 날짜와 직원 정보를 기반으로 데이터 생성
    rows = []
    employee_names = sorted(set([entry['str_workempName'] for entry in work_data]))
    all_dates = get_all_dates(attendance_dict)  # 모든 날짜 포함

    # 각 날짜에 대해 직원들의 출근/퇴근 시간을 행으로 추가
    for date in all_dates:
        # 날짜 및 요일 계산
        weekday = datetime.strptime(date, '%Y-%m-%d').strftime('%A')

        if weekday == 'Saturday' or weekday == 'Sunday':
            continue

        weekday_kr = {'Monday': '월', 'Tuesday': '화', 'Wednesday': '수',
                      'Thursday': '목', 'Friday': '금'}[weekday]

        # 행 생성
        row = [datetime.strptime(date, '%Y-%m-%d').strftime('%Y/%m/%d'), weekday_kr, '']
        if date in attendance_dict:
            for employee in employee_names:
                if employee in attendance_dict[date]:
                    row.append((attendance_dict[date][employee]['출근'], attendance_dict[date][employee]['출근_추정']))
                    row.append((attendance_dict[date][employee]['퇴근'], attendance_dict[date][employee]['퇴근_추정']))
                else:
                    row.append(('', False))
                    row.append(('', False))
        else:
            for _ in employee_names:
                row.append(('', False))
                row.append(('', False))

        rows.append(row)

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

    # 데이터 직접 삽입
    for r_idx, row in enumerate(rows, start=3):
        for c_idx, value in enumerate(row, 1):
            cell = sheet.cell(row=r_idx, column=c_idx)
            cell.value = value[0] if isinstance(value, tuple) else value

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

    # 지각 및 추정값 스타일 적용
    late_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    late_font = Font(color='FF0000')
    estimated_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')

    for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row, min_col=4, max_col=sheet.max_column):
        for idx, cell in enumerate(row):
            if cell.value:
                row_idx = cell.row - 3
                col_idx = cell.column - 1
                is_estimated = isinstance(rows[row_idx][col_idx], tuple) and rows[row_idx][col_idx][1]
                if idx % 2 == 0:  # 출근 열
                    try:
                        time_val = datetime.strptime(cell.value, '%H:%M:%S')
                        if time_val > datetime.strptime('08:30:59', '%H:%M:%S'):
                            cell.fill = late_fill
                            cell.font = late_font
                    except ValueError:
                        pass
                if is_estimated:
                    cell.fill = estimated_fill

    return workbook

# 사용 예시
if __name__ == "__main__":
    file_path = sys.argv[1] if len(sys.argv) > 1 else "input.xlsx"  # Excel 파일 경로
    workbook = convert(file_path)
    if workbook:
        print("Conversion completed successfully")