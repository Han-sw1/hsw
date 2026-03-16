"""
로우데이터 파일 파싱 및 통계 집계 모듈
- B800/B700/B710, 공항 B620, 지역버스
"""
import os
import json
import pandas as pd
from datetime import datetime

# 파일 경로 (config.json에서 오버라이드 가능)
_BASE = os.path.dirname(__file__)
_CONFIG_PATH = os.path.join(_BASE, "config.json")

def _get_rawdata_paths():
    paths = {}
    if os.path.exists(_CONFIG_PATH):
        try:
            with open(_CONFIG_PATH, encoding="utf-8") as f:
                cfg = json.load(f)
            rd = cfg.get("rawdata_files", {})
            for key in ("b_series", "regional", "b620"):
                p = rd.get(key, "")
                if p and not os.path.isabs(p):
                    p = os.path.join(_BASE, p)
                paths[key] = p
        except Exception:
            pass
    return paths

# 단말기구분 → 짧은 이름 매핑
B_SERIES_TAB_MAP = {
    "C-ITS통합단말기(B700)":  "B700",
    "차세대통합단말기(B710)":   "B710",
    "통합단말기(B800)":        "B800",
    "공항버스단말기(B620)":     "B620",
    "공항 버스단말기(B620)":    "B620",
    "공항B620":                "B620",
}
REGIONAL_TAB_MAP = {
    "B650":              "B650",
    "신단말기(B400)":    "B400",
    "통합단말기(B600)":  "B600",
    "신단말기(B500)":    "B500",
    "통합표출단말기(B810)": "B810",
    "신단말기(B520D)":   "B520D",
}

def _month_label(dt_val):
    """날짜값 → '2025년 6월' 형식"""
    try:
        dt = pd.to_datetime(dt_val)
        return f"{dt.year}년 {dt.month}월"
    except Exception:
        return None

def _build_type_stats(df, date_col, type_col, months_sorted):
    """유형별 × 월별 건수 집계"""
    types = sorted(df[type_col].dropna().astype(str).unique().tolist())
    by_month = {}
    for t in types:
        sub = df[df[type_col].astype(str) == t]
        by_month[t] = {m: int((sub["_month"] == m).sum()) for m in months_sorted}
    total_by_type = {t: int((df[type_col].astype(str) == t).sum()) for t in types}
    return {"types": types, "by_month": by_month, "total_by_type": total_by_type}

def parse_b_series(path):
    """B800/B700/B710 또는 B620 파일 파싱"""
    df = pd.read_excel(path)
    df["_month"] = df["날짜"].apply(_month_label)
    df = df[df["_month"].notna()]

    # 단말기 이름 정규화
    df["_tab"] = df["단말기구분"].map(B_SERIES_TAB_MAP)
    df = df[df["_tab"].notna()]

    months_sorted = sorted(df["_month"].unique().tolist(),
                           key=lambda x: (int(x.split("년")[0]), int(x.split("년")[1].replace("월","").strip())))

    result = {"months": months_sorted, "tabs": {}}
    for tab in df["_tab"].unique():
        sub = df[df["_tab"] == tab]
        by_month = {m: int((sub["_month"] == m).sum()) for m in months_sorted}
        tab_data = {
            "total": len(sub),
            "by_month": by_month,
        }
        # 접수오류유형
        if "접수오류유형" in df.columns:
            tab_data["접수오류유형"] = _build_type_stats(sub, "날짜", "접수오류유형", months_sorted)
        # 현장처리유형
        if "현장처리유형" in df.columns:
            tab_data["현장처리유형"] = _build_type_stats(sub, "날짜", "현장처리유형", months_sorted)
        result["tabs"][tab] = tab_data

    return result

def parse_regional(path):
    """지역버스 파일 파싱"""
    df = pd.read_excel(path)
    df["_month"] = df["장애접수일"].apply(_month_label)
    df = df[df["_month"].notna()]

    df["_tab"] = df["단말기구분"].map(REGIONAL_TAB_MAP)
    df = df[df["_tab"].notna()]

    months_sorted = sorted(df["_month"].unique().tolist(),
                           key=lambda x: (int(x.split("년")[0]), int(x.split("년")[1].replace("월","").strip())))

    result = {"months": months_sorted, "tabs": {}}
    for tab in df["_tab"].unique():
        sub = df[df["_tab"] == tab]
        by_month = {m: int((sub["_month"] == m).sum()) for m in months_sorted}
        tab_data = {
            "total": len(sub),
            "by_month": by_month,
        }
        if "단말기접수유형" in df.columns:
            tab_data["단말기접수유형"] = _build_type_stats(sub, "장애접수일", "단말기접수유형", months_sorted)
        if "단말기현장처리" in df.columns:
            tab_data["단말기현장처리"] = _build_type_stats(sub, "장애접수일", "단말기현장처리", months_sorted)
        result["tabs"][tab] = tab_data

    return result

import threading as _threading
import database as _db

_parse_lock = _threading.Lock()


def get_rawdata_stats(force_refresh=False):
    paths = _get_rawdata_paths()
    result = {"ok": True}

    for key, parse_fn, path_key in [
        ("b_series", parse_b_series, "b_series"),
        ("b620",     parse_b_series, "b620"),
        ("regional", parse_regional, "regional"),
    ]:
        path = paths.get(path_key, "")
        if not path or not os.path.exists(path):
            result[key] = None
            continue

        mtime = os.path.getmtime(path)
        if not force_refresh:
            cached, cached_mtime = _db.rawdata_cache_get(key)
            if cached is not None and cached_mtime == mtime:
                result[key] = cached
                continue

        # 파일을 하나씩 파싱 (메모리 절약)
        with _parse_lock:
            # 락 안에서 다시 확인 (중복 파싱 방지)
            if not force_refresh:
                cached, cached_mtime = _db.rawdata_cache_get(key)
                if cached is not None and cached_mtime == mtime:
                    result[key] = cached
                    continue
            try:
                data = parse_fn(path)
                _db.rawdata_cache_set(key, data, mtime)
                result[key] = data
            except Exception as e:
                result[key] = {"error": str(e)}

    return result


def prewarm():
    """서버 시작 시 백그라운드에서 캐시 미리 로드."""
    def _run():
        try:
            get_rawdata_stats()
        except Exception:
            pass
    t = _threading.Thread(target=_run, daemon=True)
    t.start()
