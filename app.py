from flask import Flask, render_template, request, redirect, jsonify, Response, send_file, abort
from werkzeug.utils import safe_join
import pandas as pd
import subprocess
import threading
import time
import os
import shutil
import zipfile
from lxml import etree
from openpyxl import load_workbook

app = Flask(__name__)

# ==========================================================
# PATHS (Auto-adjust for local or Render)
# ==========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

data_dir = os.path.join(BASE_DIR, "data")
specimen_dir = os.path.join(data_dir, "NEWW", "Specimen")
experience_dir = os.path.join(data_dir, "Experience Letters", "Exp1")

excel_path = os.path.join(data_dir, "MainData.xlsx")
exp_samples_path = os.path.join(data_dir, "ExpManual.xlsx")

logs = []
process_running = False

# ==========================================================
# ROUTE: Experience Letter Samples
# ==========================================================
@app.route('/exp-samples')
def exp_samples():
    try:
        df = pd.read_excel(exp_samples_path).fillna("")
        required_cols = ["File Name", "Country", "Company Type", "Company", "Project"]
        for c in required_cols:
            if c not in df.columns:
                return jsonify({"status": "error", "message": f"Missing column: {c}"})
        rows = df[required_cols].to_dict(orient="records")
        return jsonify({"status": "ok", "rows": rows})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

# ==========================================================
# ROUTE: Main Page
# ==========================================================
@app.route('/')
def index():
    try:
        df = pd.read_excel(excel_path)
        records = df.to_dict(orient='records')
        columns = df.columns.tolist()
    except Exception as e:
        print(f"⚠️ Error reading Excel: {e}")
        records, columns = [], []
    return render_template('form.html', records=records, columns=columns)

# ==========================================================
# ROUTE: Submit Data
# ==========================================================
@app.route('/submit', methods=['POST'])
def submit():
    data = request.form.to_dict()
    try:
        df = pd.read_excel(excel_path)
    except FileNotFoundError:
        df = pd.DataFrame(columns=data.keys())

    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    df.to_excel(excel_path, index=False)
    return redirect('/')

# ==========================================================
# ROUTE: Clear All Data
# ==========================================================
@app.route('/clear-data', methods=['POST'])
def clear_data():
    try:
        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path)
            empty_df = pd.DataFrame(columns=df.columns)
            empty_df.to_excel(excel_path, index=False)
        else:
            pd.DataFrame().to_excel(excel_path, index=False)
        return jsonify({"status": "ok", "message": "All data cleared successfully."})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

# ==========================================================
# ROUTE: Generate PDF (runs bulk script)
# ==========================================================
@app.route('/generate-pdf')
def generate_pdf():
    global process_running, logs
    if process_running:
        return jsonify({"status": "already_running"})

    logs = []
    process_running = True

    def run_process():
        global process_running
        try:
            process = subprocess.Popen(
                ["python", "bulk_cv_auto.py"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )
            for line in process.stdout:
                logs.append(line.strip())
            process.wait()
            logs.append("🚀 Complete process finished successfully!")
        except Exception as e:
            logs.append(f"❌ Error: {e}")
        process_running = False

    threading.Thread(target=run_process, daemon=True).start()
    return jsonify({"status": "started"})

# ==========================================================
# ROUTE: Stream Logs (real-time)
# ==========================================================
@app.route('/logs')
def stream_logs():
    def event_stream():
        last_len = 0
        while process_running or last_len < len(logs):
            if last_len < len(logs):
                for i in range(last_len, len(logs)):
                    yield f"data: {logs[i]}\n\n"
                last_len = len(logs)
            time.sleep(0.5)
    return Response(event_stream(), mimetype="text/event-stream")

# ==========================================================
# ROUTE: Download final PDF (single file)
# ==========================================================
@app.route('/download/<filename>')
def download_file(filename):
    # Sanitize filename
    safe_name = os.path.basename(filename)
    file_path = os.path.join(data_dir, "output", safe_name)
    if os.path.exists(file_path):
        # send file as download
        return send_file(file_path, as_attachment=True, download_name=safe_name)
    else:
        abort(404)

# ==========================================================
# ROUTE: Certificate Submit Page
# ==========================================================
@app.route('/submit-certificate')
def submitcertificate():
    return render_template('submit-certificate.html')

# (Certificate submission code unchanged...)

# ==========================================================
# ROUTE: Photocopies Page
# ==========================================================
@app.route('/photocopies')
def photocopies():
    return render_template('photocopies.html')

# ==========================================================
# RUN APP
# ==========================================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
