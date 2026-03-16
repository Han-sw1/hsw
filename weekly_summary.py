"""
주차별 장애 요약 테이블 집계 모듈

데이터 출처:
  - 전체접수: DB weekly_전체_stats
    - data/ 폴더 파일 변경 시 자동 재계산 → DB에 저장 (이후엔 DB에서 바로 읽기)
    - 처리+확정 시점에도 업로드된 원본 데이터로 DB 갱신 가능
  - 장애접수: DB monthly_tab_stats (주차 확정 때마다 자동 갱신)
  - 주요장애: DB weekly_전체_stats.top_faults
"""
import os
import json
import re
import pandas as pd
from datetime import date
import threading as _threading

from processor import get_week_label

_BASE = os.path.dirname(__file__)
_CONFIG_PATH = os.path.join(_BASE, "config.json")

# ─── 기종별 접수오류유형 제외 (전체접수 계산 시) ──────────────────
EXCLUDE_접수오류유형 = {
    "B800": {"타코/개폐센서이상"},
    "B700": {"타코/개폐센서이상", "통합단말기 CITS불량"},
    "B710": {"타코/개폐센서이상"},
    "B620": {"타코/개폐센서이상"},
}

# ─── 기종별 현장처리유형 제외 (전체접수 계산 시) ──────────────────
_공통_제외 = {
    "고정작업", "기처리건", "단말기 초기화", "단말기사용안내",
    "단말기위치조정", "단말기전원 On/Off", "모니터링", "볼륨조정",
    "사용 안내", "설정값 변경", "운수사 방문", "원인분석",
    "접수취소", "타코메타설정값변경", "단말기 고정작업",
}
EXCLUDE_현장처리유형 = {
    "B800": _공통_제외 | {"전화응대", "터서장비 간섭(이관)"},
    "B700": _공통_제외 | {"타사장비 간섭(이관)"},
    "B710": _공통_제외 | {"타사장비 간섭(이관)"},
    "B620": {"기처리건", "이상없음", "전화응대", "접수취소", "기타"},
}

# ─── 단말기구분 → 기종 코드 ────────────────────────────────────────
DEVICE_NAME_MAP = {
    "통합단말기(B800)":       "B800",
    "C-ITS통합단말기(B700)":  "B700",
    "차세대통합단말기(B710)":  "B710",
    "공항버스단말기(B620)":    "B620",
    "공항 버스단말기(B620)":   "B620",
    "공항B620":               "B620",
}

# DB 탭명 → 기종 코드
TAB_TO_DEVICE = {
    "서울 B800": "B800",
    "서울 B700": "B700",
    "서울 B710": "B710",
    "공항 B620": "B620",
}

DEVICE_ORDER = ["B800", "B620", "B700", "B710"]

ALLOWED_접수구분 = {
    "B800": {"전화접수"},
    "B700": {"전화접수"},
    "B710": {"전화접수"},
    "B620": {"현장접수", "B/S처리"},
}

# ─── 캐시 (data/ 폴더 mtime + 필터 설정 해시 기반) ───────────────────
_sync_lock = _threading.Lock()
_last_mtime = None
_SYNC_STATE_PATH = os.path.join(_BASE, "weekly_전체_sync_state.json")


def _load_sync_state():
    global _last_mtime
    try:
        if os.path.exists(_SYNC_STATE_PATH):
            with open(_SYNC_STATE_PATH, "r", encoding="utf-8") as f:
                _last_mtime = json.load(f).get("cache_key")
    except Exception:
        pass


def _save_sync_state(cache_key):
    try:
        with open(_SYNC_STATE_PATH, "w", encoding="utf-8") as f:
            json.dump({"cache_key": cache_key}, f)
    except Exception:
        pass


_load_sync_state()


def _filter_hash():
    """필터 설정이 바뀌면 해시가 달라져서 DB 자동 재계산"""
    import hashlib, json as _json
    data = {
        "err": {k: sorted(v) for k, v in EXCLUDE_접수오류유형.items()},
        "proc": {k: sorted(v) for k, v in EXCLUDE_현장처리유형.items()},
        "acc": {k: sorted(v) for k, v in ALLOWED_접수구분.items()},
    }
    return hashlib.md5(_json.dumps(data, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:8]


def _get_data_dir():
    return os.environ.get("DATA_DIR", _BASE)


def _find_data_files():
    """data/ 폴더에서 B-시리즈, B620 파일 자동 탐색"""
    data_dir = os.path.join(_get_data_dir(), "data")
    if not os.path.exists(data_dir):
        return [], []
    b_series, b620 = [], []
    for fn in os.listdir(data_dir):
        if fn.startswith("~$") or not (fn.endswith(".xlsx") or fn.endswith(".xls")):
            continue
        path = os.path.join(data_dir, fn)
        if "B620" in fn or "공항" in fn:
            b620.append(path)
        elif any(x in fn for x in ("B800", "B700", "B710")):
            b_series.append(path)
    return b_series, b620


def _files_mtime(paths):
    mt = 0.0
    for p in paths:
        try:
            mt = max(mt, os.path.getmtime(p))
        except Exception:
            pass
    return mt


def _parse_date(val):
    try:
        if isinstance(val, date):
            return val
        if hasattr(val, "date"):
            return val.date()
        return pd.to_datetime(str(val)).date()
    except Exception:
        return None


# ─── 전체접수 계산 (DataFrame → DB 저장) ─────────────────────────────

def _calc_top_faults(wdf):
    """전체접수 레코드에서 접수오류유형 기준 주요장애 계산"""
    if "접수오류유형" not in wdf.columns:
        return []
    series = (
        wdf["접수오류유형"]
        .replace({"nan": None, "None": None, "": None})
        .dropna()
        .str.strip()
    )
    if series.empty:
        return []
    counts = series.value_counts()
    car_col = "차량번호" if "차량번호" in wdf.columns else None
    top = []
    for name, cnt in counts.items():
        fname = str(name).strip()
        if not fname or fname in ("nan", "None"):
            continue
        v2 = 0
        if car_col:
            fault_rows = wdf[wdf["접수오류유형"].astype(str).str.strip() == fname]
            v2 = int((fault_rows[car_col].value_counts() >= 2).sum())
        top.append({"name": fname, "total": int(cnt), "vehicles_2plus": v2})
    top.sort(key=lambda x: x["total"], reverse=True)
    return top[:5]


def compute_전체접수_from_df(df):
    """
    이미 로드된 DataFrame에서 전체접수 통계 계산.
    반환: {device: {week_label: {"total": int, "top_faults": [...], "min_date": date}}}
    """
    if df.empty:
        return {}

    date_col = "날짜" if "날짜" in df.columns else "장애접수일시"
    if date_col not in df.columns:
        return {}

    df = df.copy()
    df["_날짜"] = df[date_col].apply(_parse_date)
    df = df[df["_날짜"].notna()]
    if df.empty:
        return {}

    df["_주차"] = df["_날짜"].apply(get_week_label)
    df["_기종"] = df.get("단말기구분", pd.Series(dtype=str)).map(DEVICE_NAME_MAP)
    df = df[df["_기종"].notna()]
    if df.empty:
        return {}

    # 접수구분 필터
    df = df[df.apply(
        lambda r: str(r.get("접수구분", "")).strip() in ALLOWED_접수구분.get(r["_기종"], {"전화접수"}),
        axis=1
    )]

    # 전체접수 필터 (접수오류유형 + 현장처리유형 제외)
    def _is_전체(row):
        기종 = row["_기종"]
        오류 = str(row.get("접수오류유형", "") or "").strip()
        처리 = str(row.get("현장처리유형", "") or "").strip()
        if 오류 in EXCLUDE_접수오류유형.get(기종, set()):
            return False
        if 처리 in EXCLUDE_현장처리유형.get(기종, set()):
            return False
        if 기종 == "B620" and (not 처리 or 처리 in ("nan", "None", "NaN")):
            return False
        return True

    df = df[df.apply(_is_전체, axis=1)]
    if df.empty:
        return {}

    result = {}
    week_min = df.groupby("_주차")["_날짜"].min().to_dict()

    for device in DEVICE_ORDER:
        sub = df[df["_기종"] == device]
        if sub.empty:
            continue
        for week in sub["_주차"].unique():
            wdf = sub[sub["_주차"] == week]
            result.setdefault(device, {})[week] = {
                "total": len(wdf),
                "top_faults": _calc_top_faults(wdf),
                "min_date": week_min.get(week, date.min),
            }

    return result


def _load_and_combine(file_paths):
    dfs = []
    for path in file_paths:
        try:
            dfs.append(pd.read_excel(path, dtype=str))
        except Exception as e:
            print(f"[weekly_summary] 파일 읽기 오류 {path}: {e}")
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True).drop_duplicates().reset_index(drop=True)


def sync_전체접수_to_db(force=False):
    """
    data/ 폴더 파일을 읽어 weekly_전체_stats DB에 저장.
    파일 변경 없고 필터 설정도 동일하면 스킵.
    """
    global _last_mtime
    import database as _db

    b_series, b620 = _find_data_files()
    all_paths = b_series + b620
    if not all_paths:
        return

    current_mtime = _files_mtime(all_paths)
    fhash = _filter_hash()
    # mtime에 필터 해시를 결합해서 필터 변경도 감지
    cache_key = f"{current_mtime}:{fhash}"

    with _sync_lock:
        if not force and _last_mtime == cache_key:
            return  # 변경 없음 → 스킵

        df = _load_and_combine(all_paths)
        stats = compute_전체접수_from_df(df)

        # data/ 폴더에서 계산된 주차만 업데이트 (수동 업로드 주차는 보존)
        # 필터 해시가 바뀌었으면 data/ 폴더에서 커버하는 주차만 삭제 후 재삽입
        if _last_mtime is not None and _last_mtime.split(":")[1] != fhash:
            # 필터 변경: data/ 파일에서 나온 주차 목록만 삭제
            try:
                import sqlite3
                conn = sqlite3.connect(os.path.join(_get_data_dir(), "atmo.db"))
                for device, weeks in stats.items():
                    for week_label in weeks.keys():
                        conn.execute(
                            "DELETE FROM weekly_전체_stats WHERE week_label=? AND device=?",
                            (week_label, device)
                        )
                conn.commit()
                conn.close()
            except Exception:
                pass

        for device, weeks in stats.items():
            for week_label, data in weeks.items():
                _db.upsert_weekly_전체(
                    week_label, device,
                    data["total"], data["top_faults"]
                )

        # raw_confirmed/ pkl 데이터를 xlsx보다 나중에 적용 (우선순위 높음)
        # xlsx가 날짜 범위 밖의 주차(예: 3월1주 중 3/1~3/4)를 덮어쓰지 않도록 보정
        import pickle as _pkl
        _rc_dir = os.path.join(_get_data_dir(), "raw_confirmed")
        if os.path.exists(_rc_dir):
            for _fname in sorted(os.listdir(_rc_dir)):
                if not _fname.endswith(".pkl"):
                    continue
                try:
                    with open(os.path.join(_rc_dir, _fname), "rb") as _pf:
                        _df = _pkl.load(_pf)
                    _pkl_stats = compute_전체접수_from_df(_df)
                    for _dev, _weeks in _pkl_stats.items():
                        for _wk, _wdata in _weeks.items():
                            _db.upsert_weekly_전체(
                                _wk, _dev,
                                _wdata["total"], _wdata["top_faults"]
                            )
                except Exception:
                    pass

        _last_mtime = cache_key
        _save_sync_state(cache_key)


def upsert_전체접수_from_raw(df):
    """
    처리+확정 시 업로드된 원본 DataFrame으로 전체접수 DB 갱신.
    (3월1주 이후 신규 주차에 사용)
    """
    import database as _db
    stats = compute_전체접수_from_df(df)
    for device, weeks in stats.items():
        for week_label, data in weeks.items():
            _db.upsert_weekly_전체(
                week_label, device,
                data["total"], data["top_faults"]
            )
    return stats


# ─── 장애접수 조회 (DB monthly_tab_stats 기준) ───────────────────────

def _get_장애접수_from_db():
    """DB → {device: {week_label: {"fault": int}}}

    중복 방지: 주차 레이블 월 == 파일 월인 경우만 집계
    (예: 3월1주는 3월 파일에서만, 2월 파일의 3월1주 제외)
    """
    import database as _db
    all_stats = _db.get_all_historical_stats()
    result = {}
    for entry in all_stats.values():
        file_month = entry.get("month")
        for tab, td in entry.get("tabs", {}).items():
            device = TAB_TO_DEVICE.get(tab)
            if not device:
                continue
            for week_label, cnt in td.get("by_week", {}).items():
                if not cnt:
                    continue
                # 주차 레이블에서 월 추출 ("3월1주" → 3)
                m = re.match(r"(\d+)월", str(week_label))
                if m and file_month is not None:
                    wk_month = int(m.group(1))
                    if wk_month != file_month:
                        continue  # 파일 월 != 주차 월 → 중복 방지
                result.setdefault(device, {})
                if week_label not in result[device]:
                    result[device][week_label] = {"fault": 0}
                result[device][week_label]["fault"] += cnt
    return result


# ─── 메인 함수 ───────────────────────────────────────────────────────

def compute_weekly_summary():
    """
    주차별 장애 요약 테이블 데이터 반환.
    DB에서 읽으므로 빠름. data/ 파일 변경 시 자동 갱신.
    """
    import database as _db

    # data/ 폴더 파일 → DB 동기화 (변경 시만)
    try:
        sync_전체접수_to_db()
    except Exception as e:
        print(f"[weekly_summary] DB 동기화 오류: {e}")

    # DB에서 읽기
    전체_db = _db.get_weekly_전체_all()   # {device: {week: {total, top_faults}}}
    장애_db = _get_장애접수_from_db()     # {device: {week: {fault}}}

    # 수동 override 적용 (이미지 기준 보정값)
    _overrides_path = os.path.join(_BASE, "weekly_전체_overrides.json")
    if os.path.exists(_overrides_path):
        try:
            with open(_overrides_path, "r", encoding="utf-8") as _f:
                _overrides = json.load(_f)
            for _dev, _weeks in _overrides.items():
                for _wk, _total in _weeks.items():
                    if _dev in 전체_db and _wk in 전체_db[_dev]:
                        전체_db[_dev][_wk]["total"] = _total
        except Exception:
            pass

    if not 전체_db:
        return {"ok": False, "error": "전체접수 데이터 없음 (data/ 폴더 파일을 확인해주세요)"}

    # 주차 정렬: DB에 min_date가 없으므로 week_label 날짜 근사 정렬
    # 전체 주차 목록 수집
    all_weeks_set = set()
    for dev_weeks in 전체_db.values():
        all_weeks_set.update(dev_weeks.keys())

    def _week_approx_sort(w):
        # "3월1주" → (3, 1), 연도 구분 없으므로 현재 파악 범위(6~12→2025, 1~5→2026) 처리
        m = re.match(r"(\d+)월(\d+)주", str(w))
        if not m:
            return (99, 99)
        mn, wn = int(m.group(1)), int(m.group(2))
        # 6~12월은 2025년 (앞), 1~5월은 2026년 (뒤)
        year_offset = 0 if mn >= 6 else 100
        return (year_offset + mn, wn)

    all_weeks_sorted = sorted(all_weeks_set, key=_week_approx_sort)

    if not all_weeks_sorted:
        return {"ok": False, "error": "주차 데이터 없음"}

    current_week = all_weeks_sorted[-1]
    prev_weeks = all_weeks_sorted[-5:-1]
    needed = set(prev_weeks) | {current_week}

    # ─── 기종별 데이터 조합 ──────────────────────────────────────
    result_data = {}

    for device in DEVICE_ORDER:
        전체 = 전체_db.get(device, {})
        장애 = 장애_db.get(device, {})

        by_week = {}
        for week in needed:
            total = 전체.get(week, {}).get("total", 0)
            fault = 장애.get(week, {}).get("fault", 0)
            fault = min(fault, total)
            nonfault = total - fault
            if total > 0 or fault > 0:
                by_week[week] = {"total": total, "fault": fault, "nonfault": nonfault}

        if not by_week and device not in 전체_db:
            continue

        cur_total = 전체.get(current_week, {}).get("total", 0)
        cur_fault = min(장애.get(current_week, {}).get("fault", 0), cur_total)
        prev_week = prev_weeks[-1] if prev_weeks else None
        prev_total = 전체.get(prev_week, {}).get("total", 0) if prev_week else 0

        result_data[device] = {
            "by_week": by_week,
            "current": {
                "total": cur_total,
                "fault": cur_fault,
                "nonfault": cur_total - cur_fault,
                "prev_diff": cur_total - prev_total,
                "top_faults": 전체.get(current_week, {}).get("top_faults", []),
            },
        }

    if not result_data:
        return {"ok": False, "error": "집계 결과 없음"}

    return {
        "ok": True,
        "devices": [d for d in DEVICE_ORDER if d in result_data],
        "prev_weeks": prev_weeks,
        "current_week": current_week,
        "data": result_data,
    }
