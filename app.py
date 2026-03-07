from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
import json
from datetime import datetime
from processor import process_fault_file, df_to_excel_bytes

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

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

    if "file" not in request.files:
        return jsonify({"error": "파일이 없습니다."}), 400

    raw_file = request.files["file"]
    criteria_path = request.form.get("criteria_path") or cfg.get("criteria_path")
    cits_path = request.form.get("cits_path") or cfg.get("cits_path")

    if not criteria_path or not os.path.exists(criteria_path):
        return jsonify({"error": "오류처리유형 기준 파일 경로를 확인해주세요."}), 400
    if not cits_path or not os.path.exists(cits_path):
        return jsonify({"error": "CITS 기준 파일 경로를 확인해주세요."}), 400

    # 임시 파일로 저장
    suffix = os.path.splitext(raw_file.filename)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        raw_file.save(tmp.name)
        tmp_path = tmp.name

    try:
        result_df, meta = process_fault_file(tmp_path, criteria_path, cits_path)
        excel_bytes = df_to_excel_bytes(result_df)

        # 결과 임시 저장
        result_dir = os.path.join(os.path.dirname(__file__), "results")
        os.makedirs(result_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        orig_name = os.path.splitext(raw_file.filename)[0]
        result_filename = f"{orig_name}_가공_{timestamp}.xlsx"
        result_path = os.path.join(result_dir, result_filename)

        with open(result_path, "wb") as f:
            f.write(excel_bytes)

        # 미리보기 데이터 (최대 50행)
        preview_cols = ["날짜", "월", "주차", "차량번호", "단말기구분", "접수오류유형", "현장처리유형", "cits"]
        available = [c for c in preview_cols if c in result_df.columns]
        preview = result_df[available].head(50).fillna("").astype(str).to_dict(orient="records")

        return jsonify({
            "ok": True,
            "meta": meta,
            "preview": preview,
            "preview_cols": available,
            "filename": result_filename,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp_path)


@app.route("/api/download/<filename>")
def download(filename):
    result_dir = os.path.join(os.path.dirname(__file__), "results")
    path = os.path.join(result_dir, filename)
    if not os.path.exists(path):
        return jsonify({"error": "파일 없음"}), 404
    return send_file(path, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
