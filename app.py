from flask import Flask, render_template, request, jsonify, send_file, make_response, redirect, url_for, session
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
import os
import shutil
import re as _re_auth

_BASE_DIR = os.environ.get("DATA_DIR", os.path.dirname(__file__))
_RESULTS_DIR = os.path.join(_BASE_DIR, "results")

# Railway Volume 초기화: 월간 파일 및 rawdata가 없으면 repo에서 복사
_VOLUME_MONTHLY = os.path.join(_BASE_DIR, "monthly_files")
_REPO_MONTHLY = os.path.join(os.path.dirname(__file__), "monthly_files")
if _BASE_DIR != os.path.dirname(__file__):
    os.makedirs(_VOLUME_MONTHLY, exist_ok=True)
    for _f in os.listdir(_REPO_MONTHLY):
        _dst = os.path.join(_VOLUME_MONTHLY, _f)
        if not os.path.exists(_dst):
            shutil.copy2(os.path.join(_REPO_MONTHLY, _f), _dst)
    # data/ 폴더 (rawdata xlsx): repo → volume 복사
    _repo_data = os.path.join(os.path.dirname(__file__), "data")
    _vol_data = os.path.join(_BASE_DIR, "data")
    if os.path.exists(_repo_data):
        os.makedirs(_vol_data, exist_ok=True)
        for _f in os.listdir(_repo_data):
            _dst = os.path.join(_vol_data, _f)
            if not os.path.exists(_dst):
                shutil.copy2(os.path.join(_repo_data, _f), _dst)
import tempfile
import json
import pickle
from datetime import datetime
from processor import process_all_files, tabs_to_excel_bytes, build_preview
from excel_writer import list_monthly_files, insert_into_monthly, compute_web_stats, get_monthly_file_path, delete_weekly, cleanup_monthly_duplicates
from analytics import (
    get_all_historical_stats, generate_comments, generate_comparison,
    read_운영수량_only, TAB_ORDER as ANALYTICS_TAB_ORDER, _parse_ym,
)
import database as _db
import threading

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")
REFERENCE_DIR = os.path.join(os.path.dirname(__file__), "reference_files")

SUPER_ADMIN = "sw_han"

DEFAULT_CONFIG = {
    "criteria_path": "",
    "cits_path": "",
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB
app.secret_key = os.environ.get("SECRET_KEY", "atmo-secret-key-change-in-prod-2024")

# ── Flask-Login 설정 ──────────────────────────────────────────────────────────
login_manager = LoginManager(app)
login_manager.login_view = "login_page"
login_manager.login_message = None


class User(UserMixin):
    def __init__(self, row):
        self.id = row["id"]
        self.username = row["username"]
        self.name = row["name"]
        self.is_admin = bool(row["is_admin"])

    def get_id(self):
        return str(self.id)


@login_manager.user_loader
def load_user(user_id):
    row = _db.get_user_by_id(int(user_id))
    return User(row) if row else None


@app.context_processor
def inject_super_admin():
    from flask_login import current_user
    try:
        is_super = current_user.is_authenticated and current_user.username == SUPER_ADMIN
    except Exception:
        is_super = False
    return {"is_super_admin": is_super}

# DB 초기화 (백그라운드)
def _startup_init():
    try:
        _db.init_db()
        if _db.is_db_empty():
            _db.migrate_excel_stats()
        # 슈퍼관리자 계정 자동 생성 (환경변수 SUPER_ADMIN_PW 설정 시)
        _super_pw = os.environ.get("SUPER_ADMIN_PW", "")
        if _super_pw and not _db.get_user_by_username(SUPER_ADMIN):
            from werkzeug.security import generate_password_hash as _gph
            _db.create_user(SUPER_ADMIN, _gph(_super_pw), "한상우")
            _db.set_admin(SUPER_ADMIN, True)
            print(f"[DB] 슈퍼관리자 '{SUPER_ADMIN}' 자동 생성 완료")
    except Exception as e:
        print(f"[DB] 초기화 오류: {e}")

threading.Thread(target=_startup_init, daemon=True).start()

# 서버 시작 시 rawdata 캐시 미리 로드 (백그라운드)
try:
    from rawdata import prewarm as _rawdata_prewarm
    _rawdata_prewarm()
except Exception:
    pass


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_confirmed_weeks():
    """DB에서 confirmed_weeks 반환."""
    return _db.load_confirmed_weeks()


def get_reference_paths():
    """reference_files/ 폴더에서 기준파일 자동 탐지. config.json 경로를 fallback으로 사용."""
    criteria, cits = None, None
    if os.path.exists(REFERENCE_DIR):
        for fn in os.listdir(REFERENCE_DIR):
            if fn.startswith('.'):
                continue
            lower = fn.lower()
            if '오류처리유형' in fn and (lower.endswith('.xls') or lower.endswith('.xlsx')):
                criteria = os.path.join(REFERENCE_DIR, fn)
            elif 'cits' in lower and (lower.endswith('.xls') or lower.endswith('.xlsx')):
                cits = os.path.join(REFERENCE_DIR, fn)
    cfg = load_config()
    criteria = criteria or cfg.get("criteria_path", "")
    cits = cits or cfg.get("cits_path", "")
    return criteria, cits


# ── 인증 라우트 ───────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login_page():
    if current_user.is_authenticated:
        return redirect(url_for("index"))
    error = None
    success = request.args.get("signup_success")
    signup_msg = f"'{success}' 님, 회원가입이 완료되었습니다. 로그인해주세요." if success else None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        row = _db.get_user_by_username(username)
        if row and check_password_hash(row["password_hash"], password):
            login_user(User(row), remember=False)
            return redirect(request.args.get("next") or url_for("index"))
        error = "아이디 또는 비밀번호가 올바르지 않습니다."
    return render_template("login.html", error=error, signup_msg=signup_msg)


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login_page"))


@app.route("/check_username")
def check_username():
    username = request.args.get("username", "").strip()
    if not _re_auth.match(r'^[a-zA-Z0-9_]{4,20}$', username):
        return {"available": False, "message": "아이디 형식이 올바르지 않습니다."}
    user = _db.get_user_by_username(username)
    if user:
        return {"available": False, "message": "이미 사용 중인 아이디입니다."}
    return {"available": True, "message": "사용 가능한 아이디입니다."}


@app.route("/signup", methods=["GET", "POST"])
def signup_page():
    if current_user.is_authenticated:
        return redirect(url_for("index"))
    error = None
    success = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        name = request.form.get("name", "").strip()
        password = request.form.get("password", "")
        password2 = request.form.get("password2", "")

        if not _re_auth.match(r'^[a-zA-Z0-9_]{4,20}$', username):
            error = "아이디는 영문/숫자/언더스코어 4~20자로 입력해주세요."
        elif not name:
            error = "이름을 입력해주세요."
        elif len(password) < 6:
            error = "비밀번호는 6자 이상이어야 합니다."
        elif password != password2:
            error = "비밀번호가 일치하지 않습니다."
        else:
            ok = _db.create_user(username, generate_password_hash(password), name)
            if ok:
                if username == SUPER_ADMIN:
                    _db.set_admin(username, True)
                return redirect(url_for("login_page", signup_success=name))
            else:
                error = "이미 사용 중인 아이디입니다."
    return render_template("signup.html", error=error, success=success)


# ── 메인 페이지 ───────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    cfg = load_config()
    return render_template("index.html", config=cfg, active_page="index")


@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify(load_config())


@app.route("/api/reference-files")
def get_reference_files():
    """현재 서버에 저장된 기준파일 목록 반환."""
    criteria, cits = get_reference_paths()
    return jsonify({
        "criteria": os.path.basename(criteria) if criteria and os.path.exists(criteria) else None,
        "cits": os.path.basename(cits) if cits and os.path.exists(cits) else None,
    })


@app.route("/api/upload-reference", methods=["POST"])
def upload_reference():
    """기준파일 서버 업로드 (reference_files/ 폴더에 저장)."""
    file_type = request.form.get("type")  # "criteria" or "cits"
    f = request.files.get("file")
    if not f or f.filename == "":
        return jsonify({"error": "파일 없음"}), 400
    if file_type not in ("criteria", "cits"):
        return jsonify({"error": "type 파라미터 오류"}), 400

    os.makedirs(REFERENCE_DIR, exist_ok=True)
    ext = os.path.splitext(f.filename)[1]
    save_name = f"criteria{ext}" if file_type == "criteria" else f"cits{ext}"
    save_path = os.path.join(REFERENCE_DIR, save_name)
    f.save(save_path)
    return jsonify({"ok": True, "filename": f.filename, "saved_as": save_name})


@app.route("/api/config", methods=["POST"])
def update_config():
    data = request.json
    cfg = load_config()
    if "criteria_path" in data:
        cfg["criteria_path"] = data["criteria_path"]
    if "cits_path" in data:
        cfg["cits_path"] = data["cits_path"]
    save_config(cfg)
    return jsonify({"ok": True})


def _is_confirmed(filename):
    """확정(잠금) 상태 여부: 하드코딩 날짜 범위 또는 DB에 월마감 마커가 있으면 True."""
    import re
    m = re.search(r'(\d{4})년\s*(\d{2})월', filename)
    if not m:
        return False
    year, month = int(m.group(1)), int(m.group(2))
    if year < 2025:
        return True
    if year == 2025 and month >= 6:
        return True
    if year == 2026 and month <= 2:
        return True
    # 그 이후 월은 DB에서 월마감 마커 확인
    cw = load_confirmed_weeks()
    return "__월마감__" in cw.get(filename, [])


@app.route("/api/monthly-files")
def get_monthly_files():
    files = list_monthly_files()
    return jsonify({
        "files": [
            {"name": f, "confirmed": _is_confirmed(f)}
            for f in files
        ]
    })


@app.route("/api/process", methods=["POST"])
def process():
    cfg = load_config()

    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "파일이 없습니다."}), 400

    auto_criteria, auto_cits = get_reference_paths()
    criteria_path = request.form.get("criteria_path") or auto_criteria
    cits_path = request.form.get("cits_path") or auto_cits

    if not criteria_path or not os.path.exists(criteria_path):
        return jsonify({"error": "오류처리유형 기준 파일 경로를 확인해주세요. (⚙ 설정)"}), 400
    if not cits_path or not os.path.exists(cits_path):
        return jsonify({"error": "CITS 기준 파일 경로를 확인해주세요. (⚙ 설정)"}), 400

    tmp_paths = []
    try:
        for f in files:
            if f.filename == "":
                continue
            suffix = os.path.splitext(f.filename)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                f.save(tmp.name)
                tmp_paths.append(tmp.name)

        # 원본 raw DataFrame 저장 (전체접수 집계에 사용)
        import pandas as _pd_raw
        raw_dfs = []
        for p in tmp_paths:
            try:
                raw_dfs.append(_pd_raw.read_excel(p, dtype=str))
            except Exception:
                pass
        raw_df_combined = _pd_raw.concat(raw_dfs, ignore_index=True) if raw_dfs else _pd_raw.DataFrame()

        final_tabs, summary, meta = process_all_files(tmp_paths, criteria_path, cits_path)

        if not final_tabs:
            return jsonify({"error": "처리된 데이터가 없습니다. 파일 형식을 확인해주세요."}), 400

        excel_bytes = tabs_to_excel_bytes(final_tabs)

        result_dir = _RESULTS_DIR
        os.makedirs(result_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        result_filename = f"장애목록_처리결과_{timestamp}.xlsx"
        result_path = os.path.join(result_dir, result_filename)
        with open(result_path, "wb") as f:
            f.write(excel_bytes)

        # 처리 결과 저장 (월간 삽입에서 재사용)
        pkl_path = result_path + ".pkl"
        with open(pkl_path, "wb") as f:
            pickle.dump(final_tabs, f)

        # 원본 raw DataFrame 저장 (주차 확정 시 전체접수 계산에 사용)
        if not raw_df_combined.empty:
            raw_pkl_path = result_path + ".raw.pkl"
            with open(raw_pkl_path, "wb") as f:
                pickle.dump(raw_df_combined, f)


        previews = build_preview(final_tabs)
        stats = compute_web_stats(final_tabs)

        # 주차 목록 추출
        all_weeks = set()
        for tab_stats in stats.values():
            all_weeks.update(tab_stats.get("weeks", []))

        # 주 대상 월 감지 → 해당 월간 파일에서 운영수량 읽어 장애율 추가
        primary_months = {ts["primary_month"] for ts in stats.values() if ts.get("primary_month")}
        운영수량_map = {}
        if primary_months:
            try:
                pm_num = int(sorted(primary_months, key=lambda x: int(x.replace("월", "")))[-1].replace("월", ""))
                target_file = next(
                    (f for f in list_monthly_files() if f" {pm_num:02d}월" in f),
                    None
                )
                if target_file:
                    운영수량_map = read_운영수량_only(target_file)
            except Exception:
                pass

        for tab_name, ts in stats.items():
            op = 운영수량_map.get(tab_name)
            if op and op > 0:
                cnt = ts.get("total_primary", ts.get("total", 0))
                ts["운영수량"] = op
                ts["fault_rate"] = round(cnt / op * 100, 2)
            else:
                ts["운영수량"] = None
                ts["fault_rate"] = None

        # 스마트 비교 (전일/전주)
        comparison = {"comparisons": []}
        try:
            hist = get_all_historical_stats()
            comparison = generate_comparison(stats, hist)
        except Exception:
            pass

        return jsonify({
            "ok": True,
            "summary": summary,
            "filename": result_filename,
            "previews": previews,
            "stats": stats,
            "weeks": sorted(all_weeks),
            "comparison": comparison,
            "file_types": {
                os.path.basename(p): m.get("type", "unknown")
                for p, m in zip(tmp_paths, meta.values())
            },
        })

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500
    finally:
        for p in tmp_paths:
            try:
                os.unlink(p)
            except Exception:
                pass


@app.route("/api/insert-monthly", methods=["POST"])
@login_required
def insert_monthly():
    if not current_user.is_admin:
        return jsonify({"error": "관리자만 사용할 수 있습니다."}), 403
    data = request.json
    result_key = data.get("result_key")
    monthly_filename = data.get("monthly_filename")
    mode = data.get("mode", "daily")
    week_label = data.get("week_label", "")

    if not result_key or not monthly_filename:
        return jsonify({"error": "result_key 또는 monthly_filename 누락"}), 400

    if _is_confirmed(monthly_filename):
        return jsonify({"error": f"확정 완료된 파일입니다. 삽입이 불가합니다.\n({monthly_filename})"}), 403

    # 주차 확정 여부 체크
    if mode == "weekly" and week_label:
        cw = load_confirmed_weeks()
        if week_label in cw.get(monthly_filename, []):
            return jsonify({"error": f"'{week_label}'은 확정된 주차입니다. 삽입이 불가합니다."}), 403

    result_dir = _RESULTS_DIR
    pkl_path = os.path.join(result_dir, result_key + ".pkl")

    if not os.path.exists(pkl_path):
        return jsonify({"error": "처리된 데이터가 없습니다. 먼저 처리를 실행해주세요."}), 400

    try:
        import gc
        with open(pkl_path, "rb") as f:
            final_tabs = pickle.load(f)
        gc.collect()

        report = insert_into_monthly(
            final_tabs,
            monthly_filename,
            mode=mode,
            week_label=week_label if mode == "weekly" else None,
        )

        if "error" in report:
            return jsonify({"error": report["error"]}), 400

        # 월 마감 시 DB에 확정 마커 저장 → _is_confirmed() 에서 확인
        if mode == "monthly":
            _db.add_confirmed_week(monthly_filename, "__월마감__")

        # monthly_tab_stats DB 갱신 (월별 장애현황 동기화)
        try:
            from analytics import read_monthly_stats, _parse_ym
            _stats = read_monthly_stats(monthly_filename)
            if _stats:
                _year, _month = _parse_ym(monthly_filename)
                if _year:
                    _db.upsert_file_stats(monthly_filename, _year, _month, _stats)
        except Exception as _e:
            print(f"[insert_monthly] stats DB 갱신 오류: {_e}")

        # 주차 확정 시 원본 raw 데이터로 전체접수 DB 업데이트 + raw_confirmed/ 영구 저장
        if mode == "weekly" and week_label:
            try:
                raw_pkl_path = os.path.join(_RESULTS_DIR, result_key + ".raw.pkl")
                if os.path.exists(raw_pkl_path):
                    with open(raw_pkl_path, "rb") as _f:
                        _raw_df = pickle.load(_f)
                    from weekly_summary import upsert_전체접수_from_raw
                    upsert_전체접수_from_raw(_raw_df)
                    # raw_confirmed/ 영구 저장 (동일차량 장애현황용)
                    _rc_dir = os.path.join(_BASE_DIR, "raw_confirmed")
                    os.makedirs(_rc_dir, exist_ok=True)
                    _safe_label = week_label.replace(" ", "_").replace("/", "_")
                    _rc_path = os.path.join(_rc_dir, f"{_safe_label}.pkl")
                    with open(_rc_path, "wb") as _f2:
                        pickle.dump(_raw_df, _f2)
            except Exception as _e:
                print(f"[weekly_summary] 전체접수 DB 갱신 오류: {_e}")

        # 마지막 업데이트 기록
        action_label = {"daily": "일별 추가", "weekly": "주차 확정", "monthly": "월 마감"}.get(mode, mode)
        detail = week_label if mode == "weekly" else monthly_filename
        _save_last_update(action_label, detail)

        return jsonify({"ok": True, "report": report, "filename": monthly_filename})

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


@app.route("/api/result-data/<result_key>")
def get_result_data(result_key):
    """처리된 탭 데이터를 JSON으로 반환 (클라이언트 사이드 삽입용)."""
    import gc
    result_dir = _RESULTS_DIR
    pkl_path = os.path.join(result_dir, result_key + ".pkl")
    if not os.path.exists(pkl_path):
        return jsonify({"error": "처리된 데이터가 없습니다. 다시 처리를 실행해주세요."}), 404
    try:
        with open(pkl_path, "rb") as f:
            final_tabs = pickle.load(f)
        tabs_json = {}
        for tab_name, df in final_tabs.items():
            if df is None or df.empty:
                continue
            columns = list(df.columns)
            rows = []
            for vals in df.itertuples(index=False, name=None):
                r = []
                for v in vals:
                    if v is None:
                        r.append(None)
                    elif isinstance(v, float) and v != v:
                        r.append(None)
                    elif hasattr(v, "strftime"):
                        r.append(v.strftime("%Y-%m-%d"))
                    elif isinstance(v, (int, float, str, bool)):
                        r.append(v)
                    else:
                        r.append(str(v))
                rows.append(r)
            tabs_json[tab_name] = {"columns": columns, "rows": rows}
        del final_tabs
        gc.collect()
        return jsonify({"ok": True, "tabs": tabs_json})
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


@app.route("/api/confirmed-weeks")
def get_confirmed_weeks():
    return jsonify(load_confirmed_weeks())


# ─── 마지막 업데이트 기록 ─────────────────────────────────
_LAST_UPDATE_PATH = os.path.join(os.path.dirname(__file__), "last_update.json")

def _save_last_update(action, detail=""):
    try:
        payload = {
            "action": action,
            "detail": detail,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        with open(_LAST_UPDATE_PATH, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
    except Exception:
        pass

@app.route("/api/last-update")
@login_required
def get_last_update():
    try:
        if os.path.exists(_LAST_UPDATE_PATH):
            with open(_LAST_UPDATE_PATH, "r", encoding="utf-8") as f:
                return jsonify(json.load(f))
    except Exception:
        pass
    return jsonify({})


@app.route("/api/confirm-week", methods=["POST"])
@login_required
def confirm_week():
    if not current_user.is_admin:
        return jsonify({"error": "관리자만 사용할 수 있습니다."}), 403
    data = request.json
    filename = data.get("monthly_filename")
    week = data.get("week_label")
    if not filename or not week:
        return jsonify({"error": "파라미터 누락"}), 400
    _db.add_confirmed_week(filename, week)
    _save_last_update("주차 확정", week)
    return jsonify({"ok": True})


@app.route("/api/unconfirm-week", methods=["POST"])
@login_required
def unconfirm_week():
    if current_user.username != SUPER_ADMIN:
        return jsonify({"error": "슈퍼관리자만 사용할 수 있습니다."}), 403
    data = request.json
    filename = data.get("monthly_filename")
    week = data.get("week_label")
    if not filename or not week:
        return jsonify({"error": "파라미터 누락"}), 400
    _db.remove_confirmed_week(filename, week)
    # 주차별 장애 요약 테이블에서 해당 주차 삭제 (슬라이딩 윈도우)
    try:
        from weekly_summary import DEVICE_ORDER as _ws_devices
        for _dev in _ws_devices:
            _db.clear_weekly_전체_by_device_week(week, _dev)
    except Exception as _e:
        print(f"[weekly_summary] 전체접수 삭제 오류: {_e}")
    # 동일차량 장애현황 DB에서 해당 주차 삭제
    try:
        _safe_label = week.replace(" ", "_").replace("/", "_")
        _rc_path = os.path.join(_BASE_DIR, "raw_confirmed", f"{_safe_label}.pkl")
        if os.path.exists(_rc_path):
            os.remove(_rc_path)
        _db.sv_delete_source(f"confirmed_{_safe_label}.pkl")
    except Exception as _e:
        print(f"[same_vehicle] 주차 취소 삭제 오류: {_e}")
    return jsonify({"ok": True})


@app.route("/api/cleanup-monthly", methods=["POST"])
def cleanup_monthly_route():
    data = request.json
    filename = data.get("monthly_filename")
    if not filename:
        return jsonify({"error": "파라미터 누락"}), 400
    try:
        report = cleanup_monthly_duplicates(filename)
        if isinstance(report, dict) and "error" in report:
            return jsonify({"error": report["error"]}), 400
        return jsonify({"ok": True, "report": report})
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


@app.route("/api/delete-weekly", methods=["POST"])
@login_required
def delete_weekly_route():
    if not current_user.is_admin:
        return jsonify({"error": "관리자만 사용할 수 있습니다."}), 403
    data = request.json
    filename = data.get("monthly_filename")
    week = data.get("week_label")
    if not filename or not week:
        return jsonify({"error": "파라미터 누락"}), 400
    cw = load_confirmed_weeks()
    if week in cw.get(filename, []):
        return jsonify({"error": f"'{week}'은 확정된 주차입니다. 삭제하려면 먼저 확정을 취소하세요."}), 403
    try:
        report = delete_weekly(filename, week)
        if "error" in report:
            return jsonify({"error": report["error"]}), 400
        return jsonify({"ok": True, "report": report})
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


@app.route("/api/confirmed-weeks-stats")
def confirmed_weeks_stats():
    """주차별 탭별 건수+TOP3 반환. DB에서 직접 조회.

    파일 월 == 주차 레이블 월인 경우만 집계.
    예) 3월1주(2/26~3/4) 데이터는 3월 파일에 전부 삽입되므로 3월 파일만 기준으로 함.
    (2월 파일에 3월1주 레이블이 있어도 제외 → 중복 방지)
    """
    try:
        import re as _re
        from analytics import get_all_historical_stats

        cw = load_confirmed_weeks()
        all_stats = get_all_historical_stats()

        week_map = {}

        for k, entry in all_stats.items():
            monthly_filename = entry.get("filename", "")
            file_year  = entry.get("year")
            file_month = entry.get("month")
            year_short = str(file_year)[2:] if file_year else ""
            confirmed_for_file = cw.get(monthly_filename, [])

            for tab, td in entry.get("tabs", {}).items():
                for week_label, cnt in td.get("by_week", {}).items():
                    if cnt <= 0:
                        continue
                    wk_m = _re.search(r'(\d+)월', str(week_label))
                    wk_month = int(wk_m.group(1)) if wk_m else None

                    # 파일 월 != 주차 월이면 제외 (중복 방지)
                    # 3월1주 데이터는 3월 파일에만 있으므로 2월 파일의 3월1주는 무시
                    if wk_month is not None and file_month is not None and wk_month != file_month:
                        continue

                    display_label = f"{year_short}년 {week_label}" if year_short else week_label

                    if display_label not in week_map:
                        week_map[display_label] = {
                            "tab_counts": {},
                            "tab_faults": {},
                            "tab_ops": {},
                            "confirmed": False,
                        }
                    wm = week_map[display_label]
                    wm["tab_counts"][tab] = wm["tab_counts"].get(tab, 0) + cnt
                    # 운영수량 (주차별로 동일하므로 덮어쓰기)
                    ops = td.get("운영수량")
                    if ops:
                        wm["tab_ops"][tab] = ops

                    # fault 합산
                    raw_faults = td.get("by_week_faults", {}).get(week_label, {})
                    if raw_faults:
                        tf = wm["tab_faults"].setdefault(tab, {})
                        for fv, fc in raw_faults.items():
                            tf[fv] = tf.get(fv, 0) + fc

                    # 확정 여부
                    if _is_confirmed(monthly_filename) or (week_label in confirmed_for_file):
                        wm["confirmed"] = True

        # 정렬 + 결과 변환
        def _display_sort_key(label):
            # "26년 3월1주" → (26, 3, 1)
            m = _re.search(r'(\d+)년\s*(\d+)월(\d+)주', label)
            if m:
                return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return (99, 99, 99)

        result = []
        for display_label in sorted(week_map.keys(), key=_display_sort_key):
            wm = week_map[display_label]
            if not wm["tab_counts"]:
                continue
            if not wm["confirmed"]:
                continue
            # TOP3 계산 (합산된 fault 카운터 기준)
            tab_top3 = {
                tab: "/".join(sorted(fc, key=fc.get, reverse=True)[:3])
                for tab, fc in wm["tab_faults"].items() if fc
            }
            # 장애율 = 건수 / 운영수량 * 100
            tab_rates = {
                tab: round(cnt / wm["tab_ops"][tab] * 100, 2)
                for tab, cnt in wm["tab_counts"].items()
                if wm["tab_ops"].get(tab)
            }
            result.append({
                "week_label": display_label,
                "tab_counts": wm["tab_counts"],
                "tab_top3": tab_top3,
                "tab_rates": tab_rates,
                "tab_faults": wm["tab_faults"],
                "total": sum(wm["tab_counts"].values()),
                "confirmed": wm["confirmed"],
            })

        payload = {"ok": True, "weeks": result}
        from flask import make_response
        resp = make_response(jsonify(payload))
        resp.headers["Cache-Control"] = "no-store"
        return resp
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


@app.route("/api/weekly-summary-table")
@login_required
def weekly_summary_table():
    """주차별 장애 요약 테이블 (B800/B620/B700/B710)"""
    try:
        from weekly_summary import compute_weekly_summary
        data = compute_weekly_summary()
        from flask import make_response
        resp = make_response(jsonify(data))
        resp.headers["Cache-Control"] = "no-store"
        return resp
    except Exception as e:
        import traceback
        return jsonify({"ok": False, "error": str(e), "detail": traceback.format_exc()}), 500


@app.route("/analysis")
@login_required
def analysis_page():
    return redirect(url_for("analysis_monthly_page"))


@app.route("/analysis/monthly")
@login_required
def analysis_monthly_page():
    return render_template("analysis_monthly.html", active_page="analysis_monthly")


@app.route("/analysis/weekly")
@login_required
def analysis_weekly_page():
    return render_template("analysis_weekly.html", active_page="analysis_weekly")


@app.route("/analysis/fault-types")
@login_required
def analysis_faulttype_page():
    return render_template("analysis_faulttype.html", active_page="analysis_faulttype")


@app.route("/api/analysis-data")
def analysis_data():
    from flask import make_response
    try:
        all_stats = get_all_historical_stats()

        comments = generate_comments(all_stats)

        # 차트용 데이터 변환
        sorted_keys = sorted(all_stats.keys())
        labels = [all_stats[k]["label"] for k in sorted_keys]

        chart_datasets = {}
        for tab in ANALYTICS_TAB_ORDER:
            counts = []
            rates = []
            for k in sorted_keys:
                tab_data = all_stats[k]["tabs"].get(tab, {})
                counts.append(tab_data.get("count", 0))
                rates.append(tab_data.get("fault_rate"))
            chart_datasets[tab] = {"counts": counts, "rates": rates}

        # 각 월별 확정 여부 플래그 추가 + 불필요한 대용량 데이터 제거
        for k, v in all_stats.items():
            v["confirmed"] = _is_confirmed(v.get("filename", ""))
            for tab_data in v.get("tabs", {}).values():
                tab_data.pop("by_week_faults", None)  # 분석 대시보드에서 불필요

        resp = make_response(jsonify({
            "ok": True,
            "labels": labels,
            "chart_datasets": chart_datasets,
            "all_stats": all_stats,
            "comments": comments,
        }))
        resp.headers["Cache-Control"] = "no-store"
        return resp
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


@app.route("/api/analysis-comments")
def analysis_comments():
    """선택한 기간(from/to)에 대한 코멘트 생성."""
    try:
        from_key = request.args.get("from")
        to_key = request.args.get("to")

        all_stats = get_all_historical_stats()
        comments = generate_comments(all_stats, from_key=from_key, to_key=to_key)
        return jsonify({"ok": True, "comments": comments})
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


@app.route("/api/download/<filename>")
def download(filename):
    result_dir = _RESULTS_DIR
    path = os.path.join(result_dir, filename)
    if not os.path.exists(path):
        return jsonify({"error": "파일 없음"}), 404
    return send_file(path, as_attachment=True, download_name=filename)


@app.route("/api/download-monthly/<filename>")
def download_monthly(filename):
    file_path = get_monthly_file_path(filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "파일 없음"}), 404
    return send_file(file_path, as_attachment=True, download_name=filename)


@app.route("/rawdata")
@login_required
def rawdata_page():
    return render_template("rawdata.html", active_page="rawdata")


@app.route("/same-vehicle")
@login_required
def same_vehicle_page():
    return render_template("same_vehicle.html", active_page="same_vehicle")


@app.route("/api/same-vehicle-stats", methods=["POST"])
@login_required
def same_vehicle_stats():
    from same_vehicle import query_same_vehicle
    data = request.json or {}
    result = query_same_vehicle(
        date_from=data.get("date_from"),
        date_to=data.get("date_to"),
        terminal=data.get("terminal", "전체"),
        min_count=int(data.get("min_count", 2)),
    )
    resp = make_response(jsonify(result))
    resp.headers["Cache-Control"] = "no-store"
    return resp

@app.route("/api/admin/search-user")
@login_required
def admin_search_user():
    if not current_user.is_admin:
        return jsonify({"error": "관리자만 사용 가능합니다."}), 403
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"error": "아이디를 입력하세요."}), 400
    user = _db.get_user_by_username(q)
    if not user:
        return jsonify({"error": f"'{q}' 아이디를 찾을 수 없습니다."}), 404
    return jsonify({"ok": True, "username": user["username"], "name": user["name"], "is_admin": bool(user["is_admin"])})


@app.route("/api/admin/set-admin", methods=["POST"])
@login_required
def admin_set_admin():
    if not current_user.is_admin:
        return jsonify({"error": "관리자만 사용 가능합니다."}), 403
    data = request.json or {}
    username = data.get("username", "").strip()
    is_admin = bool(data.get("is_admin", False))
    if not username:
        return jsonify({"error": "아이디 누락"}), 400
    if username == current_user.username:
        return jsonify({"error": "자기 자신의 권한은 변경할 수 없습니다."}), 400
    if username == SUPER_ADMIN:
        return jsonify({"error": "슈퍼관리자 계정은 변경할 수 없습니다."}), 403
    user = _db.get_user_by_username(username)
    if not user:
        return jsonify({"error": f"'{username}' 아이디를 찾을 수 없습니다."}), 404
    _db.set_admin(username, is_admin)
    return jsonify({"ok": True, "username": username, "is_admin": is_admin})


@app.route("/api/admin/resync-weekly", methods=["POST"])
@login_required
def admin_resync_weekly():
    """주차별 전체접수 DB 강제 재동기화 (data/ + raw_confirmed/ pkl)"""
    if current_user.username != SUPER_ADMIN:
        return jsonify({"error": "슈퍼관리자만 사용 가능합니다."}), 403
    try:
        from weekly_summary import sync_전체접수_to_db, upsert_전체접수_from_raw
        # data/ 폴더 → DB 강제 갱신
        sync_전체접수_to_db(force=True)
        # raw_confirmed/ pkl → DB 갱신
        rc_dir = os.path.join(_BASE_DIR, "raw_confirmed")
        count = 0
        if os.path.exists(rc_dir):
            for fname in sorted(os.listdir(rc_dir)):
                if not fname.endswith(".pkl"):
                    continue
                try:
                    with open(os.path.join(rc_dir, fname), "rb") as f:
                        df = pickle.load(f)
                    upsert_전체접수_from_raw(df)
                    count += 1
                except Exception as _e:
                    print(f"[resync] {fname} 오류: {_e}")
        return jsonify({"ok": True, "pkl_count": count})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/upload-rawdata", methods=["POST"])
@login_required
def upload_rawdata():
    """전체 전화접수 현황용 원본 xlsx 업로드 (DATA_DIR/data/ 저장)."""
    if current_user.username != SUPER_ADMIN:
        return jsonify({"error": "슈퍼관리자만 사용 가능합니다."}), 403
    file_type = request.form.get("type")  # b_series | regional | b620
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "파일 없음"}), 400
    if file_type not in ("b_series", "regional", "b620"):
        return jsonify({"error": "type 파라미터 오류"}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xls", ".xlsx"):
        return jsonify({"error": "xlsx 파일만 가능합니다."}), 400

    save_dir = os.path.join(_BASE_DIR, "data")
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, f.filename)
    f.save(save_path)

    # config.json rawdata_files 갱신
    cfg = load_config()
    cfg.setdefault("rawdata_files", {})[file_type] = f"data/{f.filename}"
    save_config(cfg)

    # DB 캐시 무효화
    try:
        import database as _dbi
        _dbi.rawdata_cache_set(file_type, None, 0)
    except Exception:
        pass

    return jsonify({"ok": True, "filename": f.filename})


@app.route("/api/rawdata-stats")
def rawdata_stats():
    from rawdata import get_rawdata_stats
    result = get_rawdata_stats()
    resp = make_response(jsonify(result))
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/api/fault-type-stats")
def fault_type_stats():
    """월별 장애파일(monthly_files)에서 단말기별 접수오류유형/현장처리유형 월별 집계 반환."""
    from analytics import get_all_historical_stats, TAB_ORDER
    all_stats = get_all_historical_stats()

    # 탭 그룹 정의
    GROUPS = {
        "b_series":  {"label": "B시리즈", "tabs": ["서울 B800", "서울 B700", "서울 B710"]},
        "airport":   {"label": "공항 B620", "tabs": ["공항 B620"]},
        "regional":  {"label": "지역버스", "tabs": [
            "대전 B650", "세종 B500", "제주 B400", "포항 B800",
            "상주,영주,예천 B400", "안동 B520D", "김해 B600"
        ]},
    }

    # 월 레이블 정렬
    sorted_keys = sorted(all_stats.keys())
    months = []
    for k in sorted_keys:
        lbl = all_stats[k].get("label", "")
        if lbl and lbl not in months:
            months.append(lbl)

    result = {"ok": True, "months": months}

    for group_key, group_info in GROUPS.items():
        tabs_data = {}
        for tab in group_info["tabs"]:
            # 월별 건수
            by_month = {}
            # 장애유형 컬럼별: {col: {type: {month: count}}}
            fault_cols = {}
            for k in sorted_keys:
                entry = all_stats[k]
                lbl = entry.get("label", "")
                tab_data = entry.get("tabs", {}).get(tab)
                if not tab_data:
                    continue
                cnt = tab_data.get("count", 0)
                if cnt:
                    by_month[lbl] = by_month.get(lbl, 0) + cnt
                bmf = tab_data.get("by_month_faults", {})
                for col, type_counts in bmf.items():
                    if col not in fault_cols:
                        fault_cols[col] = {}
                    for ft, fc in type_counts.items():
                        if ft not in fault_cols[col]:
                            fault_cols[col][ft] = {}
                        fault_cols[col][ft][lbl] = fault_cols[col][ft].get(lbl, 0) + fc

            if not by_month:
                continue

            # 컬럼별 집계 정리
            col_result = {}
            for col, types_data in fault_cols.items():
                total_by_type = {ft: sum(mv.values()) for ft, mv in types_data.items()}
                col_result[col] = {
                    "types": sorted(types_data.keys(), key=lambda t: -total_by_type[t]),
                    "by_month": types_data,
                    "total_by_type": total_by_type,
                }

            tabs_data[tab] = {"by_month": by_month, **col_result}

        if tabs_data:
            result[group_key] = {"months": months, "tabs": tabs_data}

    resp = make_response(jsonify(result))
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/api/export-confirmed")
def export_confirmed():
    """confirmed_weeks DB 내용을 JSON으로 다운로드."""
    data = _db.load_confirmed_weeks()
    response = make_response(json.dumps(data, ensure_ascii=False, indent=2))
    response.headers["Content-Type"] = "application/json"
    response.headers["Content-Disposition"] = "attachment; filename=confirmed_weeks.json"
    return response


@app.route("/api/import-confirmed", methods=["POST"])
def import_confirmed():
    """confirmed_weeks.json 업로드 → DB에 저장."""
    import json as _json
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "파일 없음"}), 400
    try:
        data = _json.loads(f.read().decode("utf-8"))
    except Exception as e:
        return jsonify({"ok": False, "error": f"JSON 파싱 오류: {e}"}), 400
    # DB에 반영 (기존 데이터 유지하며 추가)
    for filename, weeks in data.items():
        for week in weeks:
            _db.add_confirmed_week(filename, week)
    return jsonify({"ok": True, "imported": len(data)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "true").lower() == "true"
    app.run(host="0.0.0.0", port=port, debug=debug)
