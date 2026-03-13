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
    get_all_historical_stats, generate_comments, generate_upload_comparison,
    generate_week_comparison,
    read_운영수량_only, TAB_ORDER as ANALYTICS_TAB_ORDER, _parse_ym,
)
import database as _db
import threading

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")
REFERENCE_DIR = os.path.join(os.path.dirname(__file__), "reference_files")

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

# DB 초기화 (백그라운드)
def _startup_init():
    try:
        _db.init_db()
        if _db.is_db_empty():
            _db.migrate_excel_stats()
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
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        row = _db.get_user_by_username(username)
        if row and check_password_hash(row["password_hash"], password):
            login_user(User(row), remember=True)
            return redirect(request.args.get("next") or url_for("index"))
        error = "아이디 또는 비밀번호가 올바르지 않습니다."
    return render_template("login.html", error=error)


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login_page"))


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
                success = f"'{name}' 님, 회원가입이 완료되었습니다. 로그인해주세요."
            else:
                error = "이미 사용 중인 아이디입니다."
    return render_template("signup.html", error=error, success=success)


# ── 메인 페이지 ───────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    cfg = load_config()
    return render_template("index.html", config=cfg)


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
    """25년6월~26년2월 파일은 확정(잠금) 상태."""
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
    return False


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

        # 신규 업로드 vs 직전 확정 월/주차 비교 코멘트
        upload_comments = []
        week_comments = []
        try:
            hist = get_all_historical_stats()
            upload_comments = generate_upload_comparison(stats, hist)
            week_comments = generate_week_comparison(stats, hist)
        except Exception:
            pass

        return jsonify({
            "ok": True,
            "summary": summary,
            "filename": result_filename,
            "previews": previews,
            "stats": stats,
            "weeks": sorted(all_weeks),
            "upload_comments": upload_comments,
            "week_comments": week_comments,
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
    return jsonify({"ok": True})


@app.route("/api/unconfirm-week", methods=["POST"])
@login_required
def unconfirm_week():
    if not current_user.is_admin:
        return jsonify({"error": "관리자만 사용할 수 있습니다."}), 403
    data = request.json
    filename = data.get("monthly_filename")
    week = data.get("week_label")
    if not filename or not week:
        return jsonify({"error": "파라미터 누락"}), 400
    _db.remove_confirmed_week(filename, week)
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


@app.route("/analysis")
@login_required
def analysis_page():
    return render_template("analysis.html")


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
    return render_template("rawdata.html")

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
