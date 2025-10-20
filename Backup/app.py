from pydoc import HTMLDoc

from flask import Flask, render_template, request, redirect, jsonify, Response
import pandas as pd
import subprocess
import threading
import time
import os

app = Flask(__name__)

excel_path = r'F:\CV_SOFTWARE\MainData.xlsx'
exp_samples_path = r'F:\CV_SOFTWARE\ExpSamples.xlsx'
logs = []
process_running = False

# ==========================================================
# ROUTE: Expereince Letter Samples
# ==========================================================
@app.route('/exp-samples')
def exp_samples():
    try:
        exp_path = r"F:\CV_SOFTWARE\ExpManual.xlsx"
        df = pd.read_excel(exp_path).fillna("")

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
# ROUTE: Clear All Data from MainData.xlsx
# ==========================================================
@app.route('/clear-data', methods=['POST'])
def clear_data():
    try:
        # Create empty DataFrame with same columns (keep headers)
        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path)
            empty_df = pd.DataFrame(columns=df.columns)
            empty_df.to_excel(excel_path, index=False)
        else:
            # If file doesn’t exist, just create blank one
            pd.DataFrame().to_excel(excel_path, index=False)
        return jsonify({"status": "ok", "message": "All data cleared successfully."})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

# ==========================================================
# ROUTE: Generate PDF (runs bulk_cv_auto.py)
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

#================================================
#NEW CERTIFICATE HTML
#================================================
@app.route('/newcertificate')
def newcertificate():
    return render_template('newcertificate.html')


from flask import request, jsonify
from docx import Document
import os

@app.route('/submit-certificate', methods=['POST'])
def submit_certificate():
    try:
        data = request.form.to_dict()

        template_name = data.get('template', 'Page 1') + '.docx'
        template_path = os.path.join(r'F:\CV_SOFTWARE\NEWW\Specimen', template_name)
        output_dir = r'F:\CV_SOFTWARE\NEWW'

        if not os.path.exists(template_path):
            return render_template('success.html', filename=output_filename)

        doc = Document(template_path)

        # Replace placeholders in all paragraphs
        for p in doc.paragraphs:
            for key, value in data.items():
                placeholder = f"<{key}>"
                if placeholder in p.text:
                    p.text = p.text.replace(placeholder, value)

        # Also check headers, footers, and tables if needed (optional — I can help if needed)

        # Output file path
        output_filename = f"{data.get('Name', 'Certificate')}_Certificate.docx"
        output_path = os.path.join(output_dir, output_filename)

        doc.save(output_path)

        return jsonify({"status": "ok", "message": "Certificate generated successfully.", "file": output_path})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})








if __name__ == '__main__':
    app.run(debug=True)
