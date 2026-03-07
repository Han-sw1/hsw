import pandas as pd
import xlrd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, timedelta, datetime
import io
import re


def get_week_label(d):
    """전주목요일~금주수요일 기준 주차 레이블 반환 (예: '2월4주')"""
    if not isinstance(d, date):
        return ""
    weekday = d.weekday()  # Mon=0, ..., Thu=3, ..., Wed=2
    days_to_wed = (2 - weekday) % 7
    week_end_wed = d + timedelta(days=days_to_wed)
    week_start_thu = week_end_wed - timedelta(days=6)

    month = week_end_wed.month
    year = week_end_wed.year

    first_day = date(year, month, 1)
    days_to_first_wed = (2 - first_day.weekday()) % 7
    first_wednesday = first_day + timedelta(days=days_to_first_wed)
    first_thursday = first_wednesday - timedelta(days=6)

    week_num = (week_start_thu - first_thursday).days // 7 + 1
    return f"{month}월{week_num}주"


def parse_date_from_str(val):
    """장애접수일시 문자열(yyyymmddHHMMSS)에서 date 반환"""
    try:
        s = str(int(val))
        return date(int(s[:4]), int(s[4:6]), int(s[6:8]))
    except Exception:
        return None


def get_device_col(단말기구분):
    """단말기구분에서 B700/B800 컬럼명 반환"""
    if "B800" in str(단말기구분):
        return "B800"
    elif "B700" in str(단말기구분):
        return "B700"
    return None


def load_error_criteria(criteria_path):
    """오류처리유형 기준 파일 로드 → {코드값명: {B700: 장애여부, B800: 장애여부}} dict"""
    df = pd.read_excel(criteria_path, dtype=str)
    # 컬럼명 공백/개행 정리
    df.columns = [str(c).strip().replace('\n', '') for c in df.columns]

    criteria = {}
    for _, row in df.iterrows():
        name = str(row.get("코드값명", "")).strip()
        if not name or name == "nan":
            continue
        criteria[name] = {
            "B700": str(row.get("B700", "")).strip(),
            "B800": str(row.get("B800", "")).strip(),
            "장애여부": str(row.get("장애여부", "")).strip(),
        }
    return criteria


def load_cits_map(cits_path):
    """CITS 기준 파일 로드 → {차량번호: 설치일(date)} dict"""
    df = pd.read_excel(cits_path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    cits_map = {}
    for _, row in df.iterrows():
        차번 = str(row.get("차량번호", "")).strip()
        설치일 = row.get("설치일", "")
        if 차번 and 차번 != "nan":
            try:
                if 설치일 and 설치일 != "nan":
                    d = pd.to_datetime(설치일).date()
                    cits_map[차번] = d
                else:
                    cits_map[차번] = None
            except Exception:
                cits_map[차번] = None
    return cits_map


def is_비장애(오류유형, device_col, criteria):
    """오류유형이 기준파일에서 비장애인지 확인"""
    if not 오류유형 or str(오류유형).strip() in ("", "nan"):
        return False
    key = str(오류유형).strip()
    if key not in criteria:
        return False
    row = criteria[key]
    # 해당 단말기 컬럼에 'o'가 없으면 적용 안됨 → 장애로 처리
    device_val = row.get(device_col, "")
    if str(device_val).strip().lower() != "o":
        return False
    return row.get("장애여부", "").strip() == "비장애"


def process_fault_file(raw_path, criteria_path, cits_path):
    """
    원본 xls → 필터/가공 → (DataFrame, 메타정보)
    """
    # 원본 로드
    wb = xlrd.open_workbook(raw_path, encoding_override='cp949')
    ws = wb.sheet_by_index(0)
    headers = ws.row_values(0)
    rows = [ws.row_values(i) for i in range(1, ws.nrows)]
    df = pd.DataFrame(rows, columns=headers)

    original_count = len(df)

    # 1. B700 / B800 필터
    df = df[df["단말기구분"].astype(str).str.contains("B700|B800", na=False)]
    b700b800_count = len(df)

    # 2. 증상 재현여부 = "재현" 필터
    df = df[df["증상 재현여부"].astype(str).str.strip() == "재현"]
    재현_count = len(df)

    # 3. 오류처리유형 기준 필터
    criteria = load_error_criteria(criteria_path)
    cits_map = load_cits_map(cits_path)

    def 장애여부_ok(row):
        device_col = get_device_col(row["단말기구분"])
        if not device_col:
            return False
        접수 = str(row.get("접수오류유형", "")).strip()
        현장 = str(row.get("현장처리유형", "")).strip()
        if is_비장애(접수, device_col, criteria):
            return False
        if is_비장애(현장, device_col, criteria):
            return False
        return True

    df = df[df.apply(장애여부_ok, axis=1)]
    장애_count = len(df)

    # 4. 날짜 파싱 및 오름차순 정렬
    df["_날짜_obj"] = df["장애접수일시"].apply(parse_date_from_str)
    df = df.sort_values("_날짜_obj").reset_index(drop=True)

    # 5. 추가 컬럼 계산
    def get_cits(row):
        차번 = str(row.get("차량번호", "")).strip()
        device_col = get_device_col(row["단말기구분"])
        if device_col not in ("B700", "B800"):
            return None
        return cits_map.get(차번, "#N/A")

    날짜_col = df["_날짜_obj"].apply(lambda d: d if d else None)
    cits_col = df.apply(get_cits, axis=1)
    월_col = df["_날짜_obj"].apply(lambda d: f"{d.month}월" if d else "")
    주차_col = df["_날짜_obj"].apply(lambda d: get_week_label(d) if d else "")

    # 6. 컬럼 순서 재배치 (원본 컬럼 사이에 삽입)
    orig_cols = [c for c in df.columns if c != "_날짜_obj"]
    # 장애접수일시 다음에 날짜, cits, 월, 주차 삽입
    idx = orig_cols.index("장애접수일시") + 1
    new_cols_order = orig_cols[:idx] + ["날짜", "cits", "월", "주차"] + orig_cols[idx:]

    df["날짜"] = 날짜_col
    df["cits"] = cits_col
    df["월"] = 월_col
    df["주차"] = 주차_col

    result_df = df[new_cols_order]

    meta = {
        "원본_전체": original_count,
        "B700B800_필터": b700b800_count,
        "재현_필터": 재현_count,
        "장애_최종": 장애_count,
    }

    return result_df, meta


def df_to_excel_bytes(df):
    """DataFrame → 스타일 적용된 xlsx bytes"""
    wb = Workbook()
    ws = wb.active
    ws.title = "장애목록"

    # 색상 정의
    HEADER_BG = "CE0E2D"
    HEADER_FONT = "FFFFFF"
    ROW_ODD = "FFF5F5"
    ROW_EVEN = "FFFFFF"
    HIGHLIGHT_COLS = {"날짜", "cits", "월", "주차"}
    HIGHLIGHT_BG = "FFE8E8"

    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = list(df.columns)

    # 헤더 행
    for col_idx, col_name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        cell.font = Font(bold=True, color=HEADER_FONT, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    ws.row_dimensions[1].height = 30

    # 데이터 행
    for row_idx, (_, row) in enumerate(df.iterrows(), 2):
        bg = ROW_ODD if row_idx % 2 == 0 else ROW_EVEN
        for col_idx, col_name in enumerate(headers, 1):
            val = row[col_name]
            if val is None:
                val = "#N/A"
            elif isinstance(val, float) and pd.isna(val):
                val = ""
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            fill_color = HIGHLIGHT_BG if col_name in HIGHLIGHT_COLS else bg
            cell.fill = PatternFill("solid", fgColor=fill_color)
            cell.font = Font(size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

    # 열 너비 자동 조정
    for col_idx, col_name in enumerate(headers, 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(col_name)), 8)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 30)

    # 헤더 고정
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
