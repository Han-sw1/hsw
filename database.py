"""
database.py - SQLite DB 관리
테이블: confirmed_weeks, monthly_tab_stats
"""
import sqlite3
import os
import json
import threading

_BASE_DIR = os.environ.get("DATA_DIR", os.path.dirname(__file__))
DB_PATH = os.path.join(_BASE_DIR, "atmo.db")

_lock = threading.Lock()


def _connect():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    """테이블 생성 (없으면 생성)."""
    # Migration: same_vehicle_raw 스키마 변경 시 기존 테이블 재생성
    with _connect() as conn:
        cols_rows = conn.execute("PRAGMA table_info(same_vehicle_raw)").fetchall()
        if cols_rows:
            existing = {row[1] for row in cols_rows}
            if '처리유형' not in existing:
                conn.execute("DROP TABLE IF EXISTS same_vehicle_raw")
                conn.execute("DROP TABLE IF EXISTS same_vehicle_sources")
                conn.commit()

    with _connect() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS confirmed_weeks (
                monthly_filename TEXT NOT NULL,
                week_label       TEXT NOT NULL,
                PRIMARY KEY (monthly_filename, week_label)
            );

            CREATE TABLE IF NOT EXISTS monthly_tab_stats (
                filename        TEXT NOT NULL,
                year            INTEGER NOT NULL,
                month           INTEGER NOT NULL,
                tab_name        TEXT NOT NULL,
                count           INTEGER DEFAULT 0,
                운영수량        INTEGER,
                fault_rate      REAL,
                by_week         TEXT DEFAULT '{}',
                by_week_top3    TEXT DEFAULT '{}',
                by_week_faults  TEXT DEFAULT '{}',
                by_month_faults TEXT DEFAULT '{}',
                updated_at      DATETIME DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (filename, tab_name)
            );

            CREATE TABLE IF NOT EXISTS weekly_전체_stats (
                week_label  TEXT NOT NULL,
                device      TEXT NOT NULL,
                total       INTEGER DEFAULT 0,
                top_faults  TEXT DEFAULT '[]',
                updated_at  DATETIME DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (week_label, device)
            );

            CREATE TABLE IF NOT EXISTS users (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                username      TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                name          TEXT NOT NULL DEFAULT '',
                is_admin      INTEGER NOT NULL DEFAULT 0,
                created_at    DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS same_vehicle_raw (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                차량번호      TEXT NOT NULL,
                날짜          TEXT NOT NULL,
                배정부서      TEXT DEFAULT '',
                단말기코드    TEXT DEFAULT '',
                접수오류유형  TEXT DEFAULT '',
                교통사업자명  TEXT DEFAULT '',
                처리유형      TEXT DEFAULT '',
                처리자        TEXT DEFAULT '',
                처리완료일시  TEXT DEFAULT '',
                source        TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_sv_date     ON same_vehicle_raw(날짜);
            CREATE INDEX IF NOT EXISTS idx_sv_car      ON same_vehicle_raw(차량번호);
            CREATE INDEX IF NOT EXISTS idx_sv_terminal ON same_vehicle_raw(단말기코드);
            CREATE INDEX IF NOT EXISTS idx_sv_source   ON same_vehicle_raw(source);

            CREATE TABLE IF NOT EXISTS same_vehicle_sources (
                source     TEXT PRIMARY KEY,
                loaded_at  DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS rawdata_cache (
                key          TEXT PRIMARY KEY,
                data         TEXT NOT NULL,
                source_mtime REAL DEFAULT 0,
                updated_at   DATETIME DEFAULT CURRENT_TIMESTAMP
            );
        """)


# ── confirmed_weeks ──────────────────────────────────────────────────────────

def load_confirmed_weeks():
    """DB → {monthly_filename: [week_label, ...]} 형식 반환."""
    with _connect() as conn:
        rows = conn.execute(
            "SELECT monthly_filename, week_label FROM confirmed_weeks"
        ).fetchall()
    result = {}
    for row in rows:
        result.setdefault(row["monthly_filename"], []).append(row["week_label"])
    return result


def add_confirmed_week(monthly_filename, week_label):
    with _lock:
        with _connect() as conn:
            conn.execute(
                "INSERT OR IGNORE INTO confirmed_weeks (monthly_filename, week_label) VALUES (?, ?)",
                (monthly_filename, week_label),
            )
            conn.commit()


def remove_confirmed_week(monthly_filename, week_label):
    with _lock:
        with _connect() as conn:
            conn.execute(
                "DELETE FROM confirmed_weeks WHERE monthly_filename=? AND week_label=?",
                (monthly_filename, week_label),
            )
            conn.commit()


# ── monthly_tab_stats ─────────────────────────────────────────────────────────

def upsert_file_stats(filename, year, month, tab_stats):
    """단일 파일의 모든 탭 통계를 DB에 저장/갱신."""
    with _lock:
        with _connect() as conn:
            conn.execute("DELETE FROM monthly_tab_stats WHERE filename=?", (filename,))
            for tab_name, ts in tab_stats.items():
                conn.execute(
                    """
                    INSERT INTO monthly_tab_stats
                        (filename, year, month, tab_name, count, 운영수량, fault_rate,
                         by_week, by_week_top3, by_week_faults, by_month_faults)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        filename, year, month, tab_name,
                        ts.get("count", 0),
                        ts.get("운영수량"),
                        ts.get("fault_rate"),
                        json.dumps(ts.get("by_week", {}), ensure_ascii=False),
                        json.dumps(ts.get("by_week_top3", {}), ensure_ascii=False),
                        json.dumps(ts.get("by_week_faults", {}), ensure_ascii=False),
                        json.dumps(ts.get("by_month_faults", {}), ensure_ascii=False),
                    ),
                )
            conn.commit()


def get_all_historical_stats():
    """DB에서 전체 월별 통계 반환. analytics.get_all_historical_stats()와 동일한 형식."""
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM monthly_tab_stats ORDER BY year, month"
        ).fetchall()
    all_stats = {}
    for row in rows:
        key = f"{row['year']}{row['month']:02d}"
        if key not in all_stats:
            all_stats[key] = {
                "filename": row["filename"],
                "year": row["year"],
                "month": row["month"],
                "label": f"{row['year']}년 {row['month']}월",
                "tabs": {},
            }
        all_stats[key]["tabs"][row["tab_name"]] = {
            "year": row["year"],
            "month": row["month"],
            "count": row["count"],
            "운영수량": row["운영수량"],
            "fault_rate": row["fault_rate"],
            "by_week": json.loads(row["by_week"] or "{}"),
            "by_week_top3": json.loads(row["by_week_top3"] or "{}"),
            "by_week_faults": json.loads(row["by_week_faults"] or "{}"),
            "by_month_faults": json.loads(row["by_month_faults"] or "{}"),
        }
    return all_stats


# ── weekly_전체_stats ─────────────────────────────────────────────────────────

def upsert_weekly_전체(week_label, device, total, top_faults):
    """주차별 전체접수 통계 저장/갱신."""
    with _lock:
        with _connect() as conn:
            conn.execute(
                """
                INSERT INTO weekly_전체_stats (week_label, device, total, top_faults)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(week_label, device) DO UPDATE SET
                    total=excluded.total,
                    top_faults=excluded.top_faults,
                    updated_at=CURRENT_TIMESTAMP
                """,
                (week_label, device, total, json.dumps(top_faults, ensure_ascii=False)),
            )
            conn.commit()


def get_weekly_전체_all():
    """전체 주차별 전체접수 통계 반환. {device: {week_label: {total, top_faults}}}"""
    with _connect() as conn:
        rows = conn.execute(
            "SELECT week_label, device, total, top_faults FROM weekly_전체_stats"
        ).fetchall()
    result = {}
    for row in rows:
        result.setdefault(row["device"], {})[row["week_label"]] = {
            "total": row["total"],
            "top_faults": json.loads(row["top_faults"] or "[]"),
        }
    return result


def clear_weekly_전체_by_device_week(week_label, device):
    """특정 주차/기종 전체접수 통계 삭제."""
    with _lock:
        with _connect() as conn:
            conn.execute(
                "DELETE FROM weekly_전체_stats WHERE week_label=? AND device=?",
                (week_label, device),
            )
            conn.commit()


def has_stats_for_file(filename):
    """해당 파일의 통계가 DB에 있는지 확인."""
    with _connect() as conn:
        row = conn.execute(
            "SELECT 1 FROM monthly_tab_stats WHERE filename=? LIMIT 1", (filename,)
        ).fetchone()
    return row is not None


def is_db_empty():
    """monthly_tab_stats 테이블이 비어 있는지 확인."""
    with _connect() as conn:
        row = conn.execute(
            "SELECT 1 FROM monthly_tab_stats LIMIT 1"
        ).fetchone()
    return row is None


# ── users ─────────────────────────────────────────────────────────────────────

def create_user(username, password_hash, name=""):
    with _lock:
        with _connect() as conn:
            try:
                conn.execute(
                    "INSERT INTO users (username, password_hash, name) VALUES (?, ?, ?)",
                    (username, password_hash, name),
                )
                conn.commit()
                return True
            except sqlite3.IntegrityError:
                return False  # 중복 username


def set_admin(username, is_admin=True):
    with _lock:
        with _connect() as conn:
            conn.execute(
                "UPDATE users SET is_admin=? WHERE username=?",
                (1 if is_admin else 0, username),
            )
            conn.commit()


def get_user_by_username(username):
    with _connect() as conn:
        return conn.execute(
            "SELECT * FROM users WHERE username=?", (username,)
        ).fetchone()


def get_user_by_id(user_id):
    with _connect() as conn:
        return conn.execute(
            "SELECT * FROM users WHERE id=?", (user_id,)
        ).fetchone()


def count_users():
    with _connect() as conn:
        row = conn.execute("SELECT COUNT(*) as cnt FROM users").fetchone()
    return row["cnt"] if row else 0


# ── migrate ───────────────────────────────────────────────────────────────────

def migrate_excel_stats():
    """모든 월간 Excel 파일에서 통계를 읽어 DB에 저장 (1회 마이그레이션)."""
    from analytics import read_monthly_stats, _parse_ym
    from excel_writer import list_monthly_files

    files = list_monthly_files()
    for fn in files:
        if has_stats_for_file(fn):
            continue  # 이미 있으면 스킵
        year, month = _parse_ym(fn)
        if not year:
            continue
        try:
            stats = read_monthly_stats(fn)
            if stats:
                upsert_file_stats(fn, year, month, stats)
                print(f"[DB] 마이그레이션 완료: {fn}")
        except Exception as e:
            print(f"[DB] 마이그레이션 오류 ({fn}): {e}")


# ── same_vehicle_raw ──────────────────────────────────────────────────────────

def sv_delete_source(source):
    """특정 source의 raw 데이터 및 소스 기록 삭제."""
    with _lock:
        with _connect() as conn:
            conn.execute("DELETE FROM same_vehicle_raw WHERE source=?", (source,))
            conn.execute("DELETE FROM same_vehicle_sources WHERE source=?", (source,))
            conn.commit()


def sv_is_source_loaded(source):
    with _connect() as conn:
        row = conn.execute(
            "SELECT 1 FROM same_vehicle_sources WHERE source=?", (source,)
        ).fetchone()
    return row is not None


def sv_insert_rows(rows, source):
    """rows: list of dict. source: 'history_B시리즈' 등."""
    with _lock:
        with _connect() as conn:
            conn.executemany(
                """INSERT INTO same_vehicle_raw
                   (차량번호, 날짜, 배정부서, 단말기코드, 접수오류유형, 교통사업자명, 처리유형, 처리자, 처리완료일시, source)
                   VALUES (?,?,?,?,?,?,?,?,?,?)""",
                [
                    (r["차량번호"], r["날짜"], r["배정부서"], r["단말기코드"],
                     r["접수오류유형"], r["교통사업자명"], r["처리유형"], r["처리자"], r["처리완료일시"], source)
                    for r in rows
                ],
            )
            conn.execute(
                "INSERT OR REPLACE INTO same_vehicle_sources (source) VALUES (?)", (source,)
            )
            conn.commit()


def sv_query(date_from, date_to, terminal, min_count):
    """
    날짜 범위 + 단말기 필터 → min_count 이상 차량의 모든 행 반환.
    반환: (total_cars, same_fault_cars, diff_fault_cars, rows)
    """
    REGIONAL = ("B650", "B500", "B400", "포항B800", "B520D", "B600")

    # 단말기 WHERE 절
    if terminal == "전체":
        term_sql = ""
        term_params = []
    elif terminal == "지역버스":
        placeholders = ",".join("?" * len(REGIONAL))
        term_sql = f"AND 단말기코드 IN ({placeholders})"
        term_params = list(REGIONAL)
    else:
        term_sql = "AND 단말기코드 = ?"
        term_params = [terminal]

    base_params = [date_from, date_to] + term_params

    with _connect() as conn:
        # 1) min_count 이상인 차량 목록
        cars_sql = f"""
            SELECT 차량번호, COUNT(*) AS cnt
            FROM same_vehicle_raw
            WHERE 날짜 >= ? AND 날짜 <= ? {term_sql}
            GROUP BY 차량번호
            HAVING cnt >= ?
        """
        car_rows = conn.execute(cars_sql, base_params + [min_count]).fetchall()
        if not car_rows:
            return 0, 0, 0, []

        car_counts = {r["차량번호"]: r["cnt"] for r in car_rows}
        car_list = list(car_counts.keys())

        # 2) 해당 차량들의 모든 행
        placeholders = ",".join("?" * len(car_list))
        rows_sql = f"""
            SELECT 차량번호, 날짜, 배정부서, 단말기코드, 접수오류유형, 교통사업자명, 노선명, 처리완료일시
            FROM same_vehicle_raw
            WHERE 날짜 >= ? AND 날짜 <= ? {term_sql}
              AND 차량번호 IN ({placeholders})
            ORDER BY cnt_order DESC, 차량번호 ASC, 날짜 ASC
        """
        # cnt_order를 ORDER BY에 쓰려면 서브쿼리 필요
        rows_sql = f"""
            SELECT r.차량번호, r.날짜, r.배정부서, r.단말기코드, r.접수오류유형,
                   r.교통사업자명, r.처리유형, r.처리자, r.처리완료일시
            FROM same_vehicle_raw r
            WHERE r.날짜 >= ? AND r.날짜 <= ? {term_sql}
              AND r.차량번호 IN ({placeholders})
            ORDER BY (SELECT COUNT(*) FROM same_vehicle_raw
                      WHERE 차량번호=r.차량번호 AND 날짜>=? AND 날짜<=? {term_sql}
                     ) DESC,
                     r.차량번호 ASC, r.날짜 ASC
        """
        all_params = [date_from, date_to] + term_params + car_list + \
                     [date_from, date_to] + term_params
        all_rows = conn.execute(rows_sql, all_params).fetchall()

        # 3) 동일/다중장애 분류
        from collections import defaultdict
        car_faults = defaultdict(set)
        for r in all_rows:
            ft = (r["접수오류유형"] or "").strip()
            if ft:
                car_faults[r["차량번호"]].add(ft)

        car_type = {
            car: "동일장애" if len(fts) <= 1 else "다중장애"
            for car, fts in car_faults.items()
        }
        # fault 없는 차량
        for car in car_list:
            if car not in car_type:
                car_type[car] = "동일장애"

        same_cnt = sum(1 for t in car_type.values() if t == "동일장애")
        diff_cnt = sum(1 for t in car_type.values() if t == "다중장애")

        result_rows = [
            {
                "차량번호": r["차량번호"],
                "날짜": r["날짜"],
                "배정부서": r["배정부서"] or "",
                "단말기구분": r["단말기코드"] or "",
                "접수오류유형": r["접수오류유형"] or "",
                "교통사업자명": r["교통사업자명"] or "",
                "처리유형": r["처리유형"] or "",
                "처리자": r["처리자"] or "",
                "처리완료일시": r["처리완료일시"] or "",
                "유형": car_type.get(r["차량번호"], ""),
                "차량건수": car_counts.get(r["차량번호"], 0),
            }
            for r in all_rows
        ]

    return len(car_list), same_cnt, diff_cnt, result_rows


# ── rawdata_cache ─────────────────────────────────────────────────────────────

def rawdata_cache_get(key):
    """key에 해당하는 캐시 데이터와 mtime 반환. 없으면 (None, 0)."""
    with _connect() as conn:
        row = conn.execute(
            "SELECT data, source_mtime FROM rawdata_cache WHERE key=?", (key,)
        ).fetchone()
    if row:
        return json.loads(row["data"]), row["source_mtime"]
    return None, 0


def rawdata_cache_set(key, data, source_mtime):
    """key에 대한 캐시 저장/갱신."""
    with _lock:
        with _connect() as conn:
            conn.execute(
                """INSERT INTO rawdata_cache (key, data, source_mtime)
                   VALUES (?, ?, ?)
                   ON CONFLICT(key) DO UPDATE SET
                       data=excluded.data,
                       source_mtime=excluded.source_mtime,
                       updated_at=CURRENT_TIMESTAMP""",
                (key, json.dumps(data, ensure_ascii=False), source_mtime),
            )
            conn.commit()
