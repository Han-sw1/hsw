from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
import json
from datetime import datetime
from processor import process_all_files, tabs_to_excel_bytes, build_preview

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")

DEFAULT_CONFIG = {
    "criteria_path": "",
    "cits_path": "",
}


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


@app.route("/")
def index():
    cfg = load_config()
    return render_template("index.html", config=cfg)


@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify(load_config())


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


@app.route("/api/process", methods=["POST"])
def process():
    cfg = load_config()

    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "파일이 없습니다."}), 400

    criteria_path = request.form.get("criteria_path") or cfg.get("criteria_path")
    cits_path = request.form.get("cits_path") or cfg.get("cits_path")

    if not criteria_path or not os.path.exists(criteria_path):
        return jsonify({"error": "오류처리유형 기준 파일 경로를 확인해주세요. (⚙ 설정)"}), 400
    if not cits_path or not os.path.exists(cits_path):
        return jsonify({"error": "CITS 기준 파일 경로를 확인해주세요. (⚙ 설정)"}), 400

    # 임시 파일 저장
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

        # 결과 저장
        result_dir = os.path.join(os.path.dirname(__file__), "results")
        os.makedirs(result_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        result_filename = f"장애목록_처리결과_{timestamp}.xlsx"
        result_path = os.path.join(result_dir, result_filename)
        with open(result_path, "wb") as f:
            f.write(excel_bytes)

        previews = build_preview(final_tabs)

        return jsonify({
            "ok": True,
            "summary": summary,
            "filename": result_filename,
            "previews": previews,
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


@app.route("/api/download/<filename>")
def download(filename):
    result_dir = os.path.join(os.path.dirname(__file__), "results")
    path = os.path.join(result_dir, filename)
    if not os.path.exists(path):
        return jsonify({"error": "파일 없음"}), 404
    return send_file(path, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
