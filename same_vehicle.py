"""
same_vehicle.py
동일차량 다발 장애현황 - SQLite DB 기반 (빠른 쿼리)

데이터 소스:
  1. data/          폴더 xlsx (25년6월~26년2월 전체)
  2. raw_confirmed/ 폴더 pkl (26년3월~ 주차 확정 원본)
최초 접근 시 DB에 적재, 이후 SQL 쿼리로 빠르게 응답.
"""
import os
import pickle
from datetime import datetime

import pandas as pd
import database as _db

_BASE = os.environ.get("DATA_DIR", os.path.dirname(__file__))
_DATA_DIR = os.path.join(_BASE, "data")
_RAW_CONFIRMED_DIR = os.path.join(_BASE, "raw_confirmed")


# ─── 컬럼 정규화 헬퍼 ────────────────────────────────────────────────────────

def _normalize_terminal(val):
    if not val or pd.isna(val):
        return ""
    s = str(val).upper()
    for code in ["B800", "B700", "B710", "B620", "B650", "B500", "B400", "B600", "B810", "B520D"]:
        if code in s:
            return "포항B800" if code == "B810" else code
    return str(val).strip()


def _parse_date(val):
    if val is None:
        return None
    s = str(val).strip()
    if not s or s in ("None", "nan", "NaT", ""):
        return None
    try:
        digits = s.replace("-", "").replace(" ", "")[:8]
        if len(digits) == 8 and digits.isdigit():
            return datetime.strptime(digits, "%Y%m%d").date()
    except Exception:
        pass
    return None


def _safe(val):
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s in ("None", "nan", "NaT") else s


def _format_date_short(val):
    s = _safe(val)
    if not s:
        return ""
    digits = s.replace("-", "").replace(" ", "")[:8]
    if len(digits) == 8 and digits.isdigit():
        return f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}"
    return s[:10]


def _detect_format(cols):
    col_set = set(cols)
    if "장애접수일시" in col_set:
        return "b_series"
    if "장애접수일" in col_set:
        return "regional"
    return "unknown"


def _df_to_rows(df):
    """DataFrame → DB 삽입용 dict list"""
    cols = list(df.columns)
    fmt = _detect_format(cols)
    if fmt == "unknown":
        return []

    rows = []

    # 날짜
    if fmt == "b_series":
        date_indices = [i for i, c in enumerate(cols) if c == "장애접수일시"]
        if len(date_indices) >= 2:
            raw_dates = df.iloc[:, date_indices[1]]
        else:
            raw_dates = df.get("장애접수일시", pd.Series(dtype=str))
    else:
        raw_dates = df.get("장애접수일", pd.Series(dtype=str))

    dates = raw_dates.apply(_parse_date)

    # 컬럼별 시리즈
    차량번호s = df.get("차량번호", pd.Series(dtype=str)).fillna("").astype(str).str.strip()

    if fmt == "b_series":
        배정부서s = df.get("배정부서", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        접수오류유형s = df.get("접수오류유형", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        교통사업자명s = df.get("교통사업자명", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        처리유형s = df.get("현장처리유형", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        처리자s = df.get("처리자", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        완료s_raw = df.get("처리완료일시", pd.Series(dtype=str))
        완료s = 완료s_raw.apply(_format_date_short)
    else:
        배정부서s = df.get("처리부서", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        접수오류유형s = df.get("단말기접수유형", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        교통사업자명s = df.get("영업소명", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        처리유형s = df.get("단말기현장처리", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        처리자s = df.get("처리자", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        완료s_raw = df.get("처리일시", pd.Series(dtype=str))
        완료s = 완료s_raw.apply(_format_date_short)

    단말기s = df.get("단말기구분", pd.Series(dtype=str)).apply(_normalize_terminal)

    for i in range(len(df)):
        d = dates.iloc[i]
        car = 차량번호s.iloc[i]
        if d is None or not car or car in ("nan", ""):
            continue
        nan_set = {"nan", "None", "NaT", ""}
        rows.append({
            "차량번호": car,
            "날짜": d.strftime("%Y-%m-%d"),
            "배정부서": "" if 배정부서s.iloc[i] in nan_set else 배정부서s.iloc[i],
            "단말기코드": 단말기s.iloc[i],
            "접수오류유형": "" if 접수오류유형s.iloc[i] in nan_set else 접수오류유형s.iloc[i],
            "교통사업자명": "" if 교통사업자명s.iloc[i] in nan_set else 교통사업자명s.iloc[i],
            "처리유형": "" if 처리유형s.iloc[i] in nan_set else 처리유형s.iloc[i],
            "처리자": "" if 처리자s.iloc[i] in nan_set else 처리자s.iloc[i],
            "처리완료일시": 완료s.iloc[i],
        })
    return rows


# ─── DB 동기화 ────────────────────────────────────────────────────────────────

def _sync_data_dir():
    """data/ 폴더 xlsx → DB 최초 1회 적재"""
    source = "history_data"
    if _db.sv_is_source_loaded(source):
        return
    if not os.path.exists(_DATA_DIR):
        return
    all_rows = []
    for fname in sorted(os.listdir(_DATA_DIR)):
        if not (fname.endswith(".xlsx") or fname.endswith(".xls")):
            continue
        path = os.path.join(_DATA_DIR, fname)
        try:
            df = pd.read_excel(path, dtype=str)
            all_rows.extend(_df_to_rows(df))
        except Exception as e:
            print(f"[same_vehicle] data/ 로드 오류 ({fname}): {e}")
    if all_rows:
        _db.sv_insert_rows(all_rows, source)
        print(f"[same_vehicle] data/ → DB 적재 완료: {len(all_rows)}행")


def _sync_raw_confirmed():
    """raw_confirmed/ pkl 파일 → DB 적재 (파일별로 중복 체크)"""
    if not os.path.exists(_RAW_CONFIRMED_DIR):
        return
    for fname in sorted(os.listdir(_RAW_CONFIRMED_DIR)):
        if not fname.endswith(".pkl"):
            continue
        source = f"confirmed_{fname}"
        if _db.sv_is_source_loaded(source):
            continue
        path = os.path.join(_RAW_CONFIRMED_DIR, fname)
        try:
            with open(path, "rb") as f:
                df = pickle.load(f)
            if not isinstance(df, pd.DataFrame) or df.empty:
                continue
            rows = _df_to_rows(df)
            if rows:
                _db.sv_insert_rows(rows, source)
                print(f"[same_vehicle] {fname} → DB 적재 완료: {len(rows)}행")
        except Exception as e:
            print(f"[same_vehicle] raw_confirmed/ 로드 오류 ({fname}): {e}")


def ensure_synced():
    """최초 요청 시 DB 동기화 실행"""
    _sync_data_dir()
    _sync_raw_confirmed()


# ─── 메인 쿼리 ───────────────────────────────────────────────────────────────

def query_same_vehicle(date_from=None, date_to=None, terminal="전체", min_count=2):
    ensure_synced()

    d_from = date_from or "2000-01-01"
    d_to = date_to or "2099-12-31"

    total, same, diff, rows = _db.sv_query(d_from, d_to, terminal, min_count)
    return {
        "ok": True,
        "total": total,
        "same_fault": same,
        "diff_fault": diff,
        "rows": rows,
    }


def add_confirmed_week_raw(df):
    """주차 확정 시 raw_confirmed/ 저장 후 즉시 DB 동기화 (app.py에서 호출)"""
    _sync_raw_confirmed()
