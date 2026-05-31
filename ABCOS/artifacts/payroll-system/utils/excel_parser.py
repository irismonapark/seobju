import pandas as pd
import re
from collections import Counter

from utils.errors import UserFacingError


def parse_a_file(file_path):
    """
    A파일(근태 엑셀)을 파싱하여 직원별 일별 근무 데이터를 반환한다.
    반환: {
        'year': 연도,
        'month': 월,
        'employees': [
            {
                'name': 이름,
                'emp_type': 직원형태 (A=사대, D=일용직, F=프리랜서),
                'days': {
                    1: {'출근': '09:00', '퇴근': '18:00', '근무일명칭': '평일'},
                    ...
                }
            },
            ...
        ]
    }
    """
    try:
        df = pd.read_excel(file_path, header=0, dtype=str)
    except Exception:
        try:
            df = pd.read_excel(file_path, header=0, engine="xlrd", dtype=str)
        except Exception as exc:
            raise UserFacingError(
                "엑셀 파일을 읽을 수 없습니다. xls/xlsx 형식과 컬럼(이름, 근무일자, 출근, 퇴근)을 확인해 주세요."
            ) from exc

    df.columns = [str(c).strip() for c in df.columns]

    name_col       = None
    date_col       = None
    type_col       = None
    in_col         = None
    out_col        = None
    emp_type_col   = None
    attendance_col = None

    for c in df.columns:
        cl = c.strip()
        if cl in ('직원형태', '고용형태', '근로형태', '직종'):
            emp_type_col = c
        elif cl in ('이름', '성명'):
            name_col = c
        elif cl in ('근무일자', '일자', '날짜'):
            date_col = c
        elif cl in ('근무일명칭', '근무구분', '구분'):
            type_col = c
        elif cl in ('출근', '출근시간', '시작'):
            in_col = c
        elif cl in ('퇴근', '퇴근시간', '종료'):
            out_col = c
        elif cl == '근태':
            attendance_col = c

    if name_col is None:
        name_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    if date_col is None:
        date_col = df.columns[2] if len(df.columns) > 2 else df.columns[1]
    if type_col is None and len(df.columns) > 3:
        type_col = df.columns[3]
    if in_col is None and len(df.columns) > 4:
        in_col = df.columns[4]
    if out_col is None and len(df.columns) > 5:
        out_col = df.columns[5]

    employees_order = []
    seen_names      = set()
    year_month_list = []

    for _, row in df.iterrows():
        name_val = str(row.get(name_col, '')).strip()
        if not name_val or name_val in ('nan', 'NaN', '이름', '성명', 'None'):
            continue
        if name_val not in seen_names:
            seen_names.add(name_val)
            employees_order.append(name_val)
        date_val = str(row.get(date_col, '')).strip()
        if date_val and date_val not in ('nan', 'NaN', 'None'):
            parsed_date = _parse_date(date_val)
            if parsed_date:
                year_month_list.append((parsed_date[0], parsed_date[1]))

    if year_month_list:
        counter     = Counter(year_month_list)
        most_common = counter.most_common(1)[0][0]
        year, month = most_common
    else:
        import datetime
        now = datetime.date.today()
        year, month = now.year, now.month

    employee_data    = {name: {} for name in employees_order}
    emp_type_map     = {name: '' for name in employees_order}
    attendance_map   = {name: {} for name in employees_order}  # {day: 근태값}

    for _, row in df.iterrows():
        name_val = str(row.get(name_col, '')).strip()
        if not name_val or name_val in ('nan', 'NaN', '이름', '성명', 'None'):
            continue
        if name_val not in employee_data:
            continue

        # 직원형태 수집 (처음 비어있지 않은 값으로)
        if emp_type_col:
            et_val = str(row.get(emp_type_col, '')).strip()
            if et_val and et_val not in ('nan', 'NaN', 'None', '') and not emp_type_map[name_val]:
                emp_type_map[name_val] = et_val.upper()

        date_val = str(row.get(date_col, '')).strip()
        if not date_val or date_val in ('nan', 'NaN', 'None'):
            continue

        parsed_date = _parse_date(date_val)
        if not parsed_date:
            continue

        d_year, d_month, d_day = parsed_date
        if d_year != year or d_month != month:
            continue

        type_val = str(row.get(type_col, '')).strip() if type_col else ''
        if type_val in ('nan', 'NaN', 'None'):
            type_val = ''

        in_val  = str(row.get(in_col,  '')).strip() if in_col  else ''
        out_val = str(row.get(out_col, '')).strip() if out_col else ''
        if in_val  in ('nan', 'NaN', 'None'): in_val  = ''
        if out_val in ('nan', 'NaN', 'None'): out_val = ''

        att_val = ''
        if attendance_col:
            att_val = str(row.get(attendance_col, '')).strip()
            if att_val in ('nan', 'NaN', 'None'):
                att_val = ''

        employee_data[name_val][d_day] = {
            '출근': in_val,
            '퇴근': out_val,
            '근무일명칭': type_val,
        }
        if att_val:
            attendance_map[name_val][d_day] = att_val

    employees = []
    for name in employees_order:
        employees.append({
            'name':       name,
            'emp_type':   emp_type_map.get(name, ''),
            'days':       employee_data[name],
            'attendance': attendance_map.get(name, {}),
        })

    if not employees:
        raise UserFacingError(
            "근태 데이터를 찾을 수 없습니다. 파일에 직원 이름과 근무일자가 포함되어 있는지 확인해 주세요."
        )

    return {
        'year':      year,
        'month':     month,
        'employees': employees,
    }


def extract_company_name(filename):
    """파일명에서 회사명 추출.

    지원 패턴:
    - 어브코스_2026-03.xlsx  → 어브코스  (첫 토큰이 회사명)
    - 어브코스 2026-03.xlsx  → 어브코스
    - 어브코스-2026-03.xlsx  → 어브코스
    - 2026-03_어브코스.xlsx  → 어브코스  (첫 토큰이 날짜이면 그 다음 비날짜 토큰 사용)
    - 2026-03 어브코스.xlsx  → 어브코스
    """
    basename = filename
    for ext in ('.xlsx', '.xls', '.XLSX', '.XLS'):
        if basename.endswith(ext):
            basename = basename[:-len(ext)]
            break

    date_like = re.compile(r'^\d{4}-\d{2}$')
    year_like  = re.compile(r'^\d{4}$')
    num_only   = re.compile(r'^\d+$')

    # 1차 분리: 공백·언더바로 분리
    space_tokens = re.split(r'[ _]', basename)

    # 각 공백토큰 내부의 첫 하이픈 위치를 확인하여 회사명 vs 날짜 결정
    candidates = []
    for tok in space_tokens:
        if not tok:
            continue
        # YYYY-MM 형태 → 날짜 토큰
        if date_like.match(tok):
            continue
        # YYYY 형태 → 날짜 토큰
        if year_like.match(tok):
            continue
        # 토큰 내에 하이픈이 있으면 하이픈 앞 부분 확인
        if '-' in tok:
            prefix = tok.split('-', 1)[0]
            # 앞부분이 4자리 숫자이면 날짜 토큰
            if year_like.match(prefix):
                # 뒷부분을 후보로 추가 (예: 2026-어브코스 → 어브코스)
                suffix = tok.split('-', 1)[1]
                if suffix and not num_only.match(suffix):
                    candidates.append(suffix)
                continue
            # 앞부분이 회사명, 뒷부분이 날짜 (예: 어브코스-2026-03 → 어브코스)
            candidates.append(prefix)
            continue
        candidates.append(tok)

    if candidates:
        return candidates[0]
    return ''


def _parse_date(date_str):
    """날짜 문자열을 (year, month, day) 튜플로 변환"""
    date_str = date_str.strip()
    patterns = [
        r'(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})',
        r'(\d{4})(\d{2})(\d{2})',
        r'(\d{2})[./\-](\d{1,2})[./\-](\d{1,2})',
    ]
    for pat in patterns:
        m = re.search(pat, date_str)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if y < 100:
                y += 2000
            return (y, mo, d)
    return None
