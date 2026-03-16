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
