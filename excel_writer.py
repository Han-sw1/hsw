import os
import pandas as pd
from openpyxl import load_workbook

MONTHLY_FILES_DIR = os.path.join(os.path.dirname(__file__), "monthly_files")

# 처리기 탭명 → 월간 파일 숨김 시트명
TAB_TO_RAWSHEET = {
    "서울 B710":           "B710 로우데이터",
    "서울 B800":           "B800로우데이터",
    "서울 B700":           "B700로우데이터",
    "공항 B620":           "공항 B620로우데이터",
    "대전 B650":           "대전 B650 로우데이터",
    "세종 B500":           "세종B500로우데이터",
    "제주 B400":           "제주 B400 로우데이터",
    "포항 B800":           "포항B800로우데이터",
    "상주,영주,예천 B400": "상주.영주.예천 로우데이터",
    "안동 B520D":          "안동B520D 로우데이터",
    "김해 B600":           "김해B600 로우데이터",
}


def list_monthly_files():
    os.makedirs(MONTHLY_FILES_DIR, exist_ok=True)
    return sorted(fn for fn in os.listdir(MONTHLY_FILES_DIR) if fn.endswith("일자별장애현황.xlsx"))


def get_monthly_file_path(filename):
    return os.path.join(MONTHLY_FILES_DIR, filename)


def _read_sheet(ws):
    """시트 전체를 (헤더 리스트, 데이터 행 리스트)로 읽기. 완전 빈 행 제외."""
    all_rows = list(ws.values)
    if not all_rows:
        return [], []
    headers = list(all_rows[0])
    data = [list(r) for r in all_rows[1:] if any(v is not None for v in r)]
    return headers, data


def _build_col_map(sheet_headers, proc_cols):
    """
    processor 컬럼명 → 시트 컬럼 위치(0-based) 매핑.
    특수 규칙:
      "날짜"  → "장애접수일시" 두 번째 등장 위치
      "cits"  → "CITS 설치일" 위치
    """
    name_to_pos = {}
    for i, h in enumerate(sheet_headers):
        if h is None:
            continue
        hs = str(h).strip()
        name_to_pos.setdefault(hs, []).append(i)

    col_map = {}
    used = set()

    for col in proc_cols:
        col_s = str(col).strip()

        if col_s == "날짜":
            positions = name_to_pos.get("장애접수일시", [])
            available = [p for p in positions if p not in used]
            if len(available) >= 2:
                # 첫 번째는 "장애접수일시"(문자열)에 이미 매핑됨 → 두 번째 선택
                target = available[1]
            elif available:
                target = available[0]
            else:
                continue
            col_map[col] = target
            used.add(target)

        elif col_s == "cits":
            for cname in ("CITS 설치일", "CITS설치일"):
                positions = name_to_pos.get(cname, [])
                avail = [p for p in positions if p not in used]
                if avail:
                    col_map[col] = avail[0]
                    used.add(avail[0])
                    break

        else:
            positions = name_to_pos.get(col_s, [])
            avail = [p for p in positions if p not in used]
            if avail:
                col_map[col] = avail[0]
                used.add(avail[0])

    return col_map


def _map_rows(new_df, col_map, num_cols):
    """DataFrame → 시트 컬럼 순서에 맞춘 row 리스트."""
    result = []
    for _, row in new_df.iterrows():
        r = [None] * num_cols
        for col, pos in col_map.items():
            if pos < num_cols:
                val = row.get(col)
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    val = None
                r[pos] = val
        result.append(r)
    return result


def _find_col_pos(sheet_headers, *names):
    for name in names:
        for i, h in enumerate(sheet_headers):
            if h is not None and str(h).strip() == name:
                return i
    return None


def _find_date_pos(sheet_headers):
    pos = _find_col_pos(sheet_headers, "장애접수일")
    if pos is not None:
        return pos
    # "장애접수일시" 두 번째 등장 (날짜 형식)
    occ = [i for i, h in enumerate(sheet_headers) if h == "장애접수일시"]
    if len(occ) >= 2:
        return occ[1]
    return None


def _clear_and_write(ws, headers, rows):
    # 기존 내용 삭제
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None
    # 헤더
    for ci, h in enumerate(headers, 1):
        ws.cell(row=1, column=ci, value=h)
    # 데이터
    for ri, row_data in enumerate(rows, 2):
        for ci, val in enumerate(row_data, 1):
            ws.cell(row=ri, column=ci, value=val)


def insert_into_monthly(final_tabs, monthly_filename, mode="daily", week_label=None):
    """
    처리된 데이터를 월간 파일 숨김 시트에 삽입.
    mode:
      "daily"   - 일별 추가 (접수번호 기준 중복 제외)
      "weekly"  - 주차 확정 (해당 주차 데이터 교체, week_label 필요: 예 "3월2주")
      "monthly" - 월 마감 (전체 교체)
    반환: {tab_name: report}
    """
    file_path = get_monthly_file_path(monthly_filename)
    if not os.path.exists(file_path):
        return {"error": f"파일 없음: {monthly_filename}"}

    wb = load_workbook(file_path)
    report = {}

    for tab_name, new_df in final_tabs.items():
        if new_df is None or new_df.empty:
            report[tab_name] = {"skipped": True, "reason": "데이터 없음"}
            continue

        sheet_name = TAB_TO_RAWSHEET.get(tab_name)
        if not sheet_name:
            report[tab_name] = {"skipped": True, "reason": "시트 매핑 없음"}
            continue
        if sheet_name not in wb.sheetnames:
            report[tab_name] = {"skipped": True, "reason": f"시트 없음: {sheet_name}"}
            continue

        ws = wb[sheet_name]
        sheet_headers, existing_rows = _read_sheet(ws)

        # 시트가 비어있으면 새 데이터 구조로 초기화
        if not sheet_headers:
            sheet_headers = list(new_df.columns)
            existing_rows = []

        num_cols = len(sheet_headers)
        col_map = _build_col_map(sheet_headers, list(new_df.columns))
        new_rows = _map_rows(new_df, col_map, num_cols)

        id_pos = _find_col_pos(sheet_headers, "접수번호", "No")
        week_pos = _find_col_pos(sheet_headers, "주차")
        date_pos = _find_date_pos(sheet_headers)
        before = len(existing_rows)

        if mode == "monthly":
            merged = new_rows

        elif mode == "weekly" and week_label:
            if week_pos is not None:
                keep = [r for r in existing_rows if str(r[week_pos]) != str(week_label)]
            else:
                keep = existing_rows
            merged = keep + new_rows

        else:  # daily
            if id_pos is not None:
                existing_ids = {str(r[id_pos]) for r in existing_rows if r[id_pos] is not None}
                # processor 컬럼 중 id_pos와 매핑된 것
                id_proc_col = next((c for c, p in col_map.items() if p == id_pos), None)
                if id_proc_col and id_proc_col in new_df.columns:
                    new_id_vals = new_df[id_proc_col].astype(str).tolist()
                    to_add = [r for r, nid in zip(new_rows, new_id_vals) if nid not in existing_ids]
                else:
                    to_add = new_rows
                merged = existing_rows + to_add
            else:
                existing_tuples = {tuple(str(v) if v is not None else "" for v in r) for r in existing_rows}
                to_add = [r for r in new_rows
                          if tuple(str(v) if v is not None else "" for v in r) not in existing_tuples]
                merged = existing_rows + to_add

        # 날짜 기준 정렬
        if date_pos is not None and merged:
            try:
                merged.sort(key=lambda r: (r[date_pos] is None, str(r[date_pos])))
            except Exception:
                pass

        _clear_and_write(ws, sheet_headers, merged)

        added = len(merged) - before
        report[tab_name] = {
            "sheet": sheet_name,
            "before": before,
            "added": added,
            "total": len(merged),
        }

    wb.save(file_path)
    return report


def _get_date_series(df):
    """날짜 컬럼을 date 객체 시리즈로 반환."""
    from datetime import date as date_type
    date_col = "날짜" if "날짜" in df.columns else "장애접수일"
    if date_col not in df.columns:
        return None
    def to_date(v):
        if isinstance(v, date_type):
            return v
        try:
            import pandas as pd
            return pd.to_datetime(str(v)).date()
        except Exception:
            return None
    return df[date_col].apply(to_date)


def compute_web_stats(final_tabs):
    """
    처리된 데이터에서 웹 통계 계산.
    건수는 실제 날짜 기준으로 월별 집계 (주차 레이블 기준 X).
    반환: {tab_name: {total, total_primary, top3, by_week, by_month, weeks}}
    """
    import pandas as pd
    stats = {}
    for tab_name, df in final_tabs.items():
        if df is None or df.empty:
            continue

        is_regional = not any(k in tab_name for k in ["B800", "B700", "B710", "B620"])
        fault_col = "단말기접수유형" if is_regional else "접수오류유형"

        # 날짜 시리즈
        date_series = _get_date_series(df)

        # 월별 건수 (실제 날짜 기준)
        by_month = {}
        primary_month = None
        if date_series is not None:
            month_vals = date_series.apply(
                lambda d: f"{d.month}월" if d else None
            ).dropna()
            if not month_vals.empty:
                by_month = {str(k): int(v) for k, v in month_vals.value_counts().sort_index().items()}
                # 가장 많은 월 = 주 대상 월
                primary_month = month_vals.value_counts().idxmax()

        # 주 대상 월의 건수 (해당 월 날짜만)
        total_primary = 0
        if primary_month and date_series is not None:
            pm_num = int(primary_month.replace("월", ""))
            mask = date_series.apply(lambda d: d is not None and d.month == pm_num)
            total_primary = int(mask.sum())
        else:
            total_primary = len(df)

        # TOP3 (해당 월 레코드 기준)
        top3 = ""
        if fault_col in df.columns:
            if primary_month and date_series is not None:
                pm_num = int(primary_month.replace("월", ""))
                mask = date_series.apply(lambda d: d is not None and d.month == pm_num)
                vals = df.loc[mask, fault_col].astype(str).str.strip()
            else:
                vals = df[fault_col].astype(str).str.strip()
            vals = vals[~vals.isin(["", "nan", "None", "NaN"])]
            if not vals.empty:
                top = vals.value_counts().head(3)
                top3 = "/".join(top.index.tolist())

        # 주차별 건수 (주차 레이블 기준 — 표시용)
        by_week = {}
        if "주차" in df.columns:
            by_week = {str(k): int(v) for k, v in df["주차"].value_counts().sort_index().items()}

        weeks = sorted(by_week.keys())

        stats[tab_name] = {
            "total": len(df),
            "total_primary": total_primary,
            "primary_month": primary_month or "",
            "top3": top3,
            "by_week": by_week,
            "by_month": by_month,
            "weeks": weeks,
        }

    return stats
