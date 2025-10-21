import os
import random
import threading
import time
import shutil
import tempfile
import zipfile
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

from flask import Flask, render_template, request, redirect, jsonify, Response, send_file, abort
import pandas as pd
from openpyxl import load_workbook
from docx import Document
from lxml import etree
import io

app = Flask(__name__)

# ==========================================================
# BASE PATHS - Adjust as per your environment
# ==========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
SPECIMEN_DIR = os.path.join(DATA_DIR, "NEWW", "Specimen")
EXP1_FOLDER = os.path.join(DATA_DIR, "experience", "Exp1")
EXP2_FOLDER = os.path.join(DATA_DIR, "experience", "Exp2")
CV_FOLDER = os.path.join(DATA_DIR, "cv_templates")

EXCEL_PATH = os.path.join(DATA_DIR, "MainData.xlsx")
EXP_AUTO_PATH = os.path.join(DATA_DIR, "ExpAuto.xlsx")
EXP_MANUAL_PATH = os.path.join(DATA_DIR, "ExpManual.xlsx")
UPDATED_PATH = os.path.join(DATA_DIR, "MainData_Updated.xlsx")

OUTPUT_DIR = os.path.join(DATA_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

logs = []
process_running = False

# ==========================================================
# Utility: Case-insensitive file finder for docx templates
# ==========================================================
def find_file_case_insensitive(folder, filename):
    if not filename or not os.path.exists(folder):
        return None
    filename = str(filename).strip()
    if not filename.lower().endswith(".docx"):
        filename += ".docx"
    filename_lower = filename.lower()
    for f in os.listdir(folder):
        if f.lower() == filename_lower:
            return os.path.join(folder, f)
    return None

# ==========================================================
# Date Adjustment and Experience Info Generation
# ==========================================================
def adjust_dates_and_generate_experience(logs):
    logs.append("[Step] Adjusting dates and experience info...")
    try:
        exp_auto_samples = pd.read_excel(EXP_AUTO_PATH)
    except Exception as e:
        logs.append(f"[ERROR] Could not read ExpAuto.xlsx: {e}")
        exp_auto_samples = pd.DataFrame()

    min_exp_years, max_exp_years = 2, 3
    min_gap_months, max_gap_months = 12, 36
    date_format = "%d-%m-%Y"
    today = datetime.today()

    try:
        df_main = pd.read_excel(EXCEL_PATH).replace("nan", pd.NA).astype("object")
    except Exception as e:
        logs.append(f"[ERROR] Could not read MainData.xlsx: {e}")
        return None

    def random_date_between(start, end):
        delta_days = (end - start).days
        return start if delta_days <= 0 else start + timedelta(days=random.randint(0, delta_days))

    def realistic_experience(start_date):
        years = random.randint(min_exp_years, max_exp_years)
        months = random.randint(0, 11)
        end_date = start_date + relativedelta(years=years, months=months)
        if end_date > today - timedelta(days=15):
            end_date = today - timedelta(days=random.randint(15, 90))
        return end_date

    for i in range(len(df_main)):
        dob_raw = df_main.loc[i, "dob"] if "dob" in df_main.columns else ""
        try:
            dob = datetime.strptime(str(dob_raw).strip(), "%d-%m-%Y")
        except:
            logs.append(f"[WARN] Row {i+1}: invalid DOB '{dob_raw}' — skipped")
            continue

        career_start = dob + relativedelta(years=18)
        if career_start.year < 2015:
            career_start = datetime(2015, 1, 1)

        exp1_start = random_date_between(career_start, career_start + relativedelta(days=365))
        exp1_end = realistic_experience(exp1_start)
        gap_months = random.randint(min_gap_months, max_gap_months)
        exp2_start = exp1_end + relativedelta(months=gap_months)
        exp2_end = realistic_experience(exp2_start)

        try:
            samples = exp_auto_samples.sample(2).reset_index(drop=True)
            df_main.loc[i, "Exp1 Company"] = str(samples.loc[0, "Company"])
            df_main.loc[i, "Exp1 Project"] = str(samples.loc[0, "Project"])
            df_main.loc[i, "From"] = exp1_start.strftime(date_format)
            df_main.loc[i, "To"] = exp1_end.strftime(date_format)
            df_main.loc[i, "Exp2 Company"] = str(samples.loc[1, "Company"])
            df_main.loc[i, "Exp2 Project"] = str(samples.loc[1, "Project"])
            df_main.loc[i, "From2"] = exp2_start.strftime(date_format)
            df_main.loc[i, "To2"] = exp2_end.strftime(date_format)
            if "File Name" in samples.columns:
                df_main.loc[i, "exp1"] = str(samples.loc[0, "File Name"])
                df_main.loc[i, "exp2"] = str(samples.loc[1, "File Name"])
            logs.append(f"[Info] Row {i+1}: Experience data generated.")
        except Exception as ex:
            logs.append(f"[WARN] Row {i+1}: Could not sample experience — {ex}")

    df_main.to_excel(UPDATED_PATH, index=False)
    logs.append(f"[OK] Updated Excel saved at: {UPDATED_PATH}")
    return df_main

# ==========================================================
# DOCX Placeholder Replacement
# ==========================================================
def replace_in_paragraph(paragraph, replacements):
    full_text = "".join([r.text for r in paragraph.runs])
    new_text = full_text
    for key, val in replacements.items():
        placeholder = f"<{key}>"
        replacement_text = "" if pd.isna(val) else str(val)
        new_text = new_text.replace(placeholder, replacement_text)
    if new_text != full_text:
        paragraph.clear()
        paragraph.add_run(new_text)

def replace_placeholders(doc_path, replacements, output_path):
    try:
        doc = Document(doc_path)
        for para in doc.paragraphs:
            replace_in_paragraph(para, replacements)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        replace_in_paragraph(para, replacements)
        doc.save(output_path)
        return True
    except Exception as e:
        return False

# ==========================================================
# Merge multiple DOCX files into one
# ==========================================================
def merge_docx(files, output_path):
    merged_document = Document()
    # Remove the default empty paragraph created by Document()
    if merged_document.paragraphs:
        p = merged_document.paragraphs[0]
        p._element.getparent().remove(p._element)
        merged_document._body.clear_content()

    for file in files:
        sub_doc = Document(file)
        for element in sub_doc.element.body:
            merged_document.element.body.append(element)
    merged_document.save(output_path)

# ==========================================================
# MAIN processing function called by /generate-docx
# ==========================================================
def process_all(logs):
    global process_running
    try:
        df = adjust_dates_and_generate_experience(logs)
        if df is None:
            logs.append("[ERROR] Date adjustment failed. Aborting.")
            process_running = False
            return

        for idx, row in df.iterrows():
            name = str(row.get("name", f"Person_{idx+1}")).strip()
            logs.append(f"[Processing] Preparing docs for: {name}")

            exp1_file = find_file_case_insensitive(EXP1_FOLDER, str(row.get("exp1", "")))
            exp2_file = find_file_case_insensitive(EXP2_FOLDER, str(row.get("exp2", "")))
            cv_file = find_file_case_insensitive(CV_FOLDER, str(row.get("cv", "")))

            temp_dir = tempfile.mkdtemp(prefix=f"tmp_{name}_")

            replaced_files = []
            for src in [exp1_file, exp2_file, cv_file]:
                if src and os.path.exists(src):
                    dest = os.path.join(temp_dir, os.path.basename(src))
                    success = replace_placeholders(src, row.to_dict(), dest)
                    if success:
                        replaced_files.append(dest)
                    else:
                        logs.append(f"[WARN] Failed to replace placeholders in {src}")
                else:
                    logs.append(f"[WARN] Missing source file for {name}: {src or 'Unknown'}")

            if not replaced_files:
                logs.append(f"[WARN] No DOCX files replaced for {name}, skipping merge.")
                shutil.rmtree(temp_dir)
                continue

            output_docx = os.path.join(OUTPUT_DIR, f"{name}.docx")
            merge_docx(replaced_files, output_docx)
            logs.append(f"[OK] Merged DOCX created: {output_docx}")

            shutil.rmtree(temp_dir)
        logs.append("[DONE] All documents processed.")
    except Exception as e:
        logs.append(f"[ERROR] Unexpected error: {e}")
    finally:
        process_running = False

# ==========================================================
# ROUTES
# ==========================================================
@app.route('/')
def index():
    try:
        df = pd.read_excel(EXCEL_PATH)
        records = df.to_dict(orient='records')
        columns = df.columns.tolist()
    except Exception as e:
        print(f"⚠️ Error reading Excel: {e}")
        records, columns = [], []
    return render_template('form.html', records=records, columns=columns)

@app.route('/submit', methods=['POST'])
def submit():
    data = request.form.to_dict()
    try:
        df = pd.read_excel(EXCEL_PATH)
    except FileNotFoundError:
        df = pd.DataFrame(columns=data.keys())

    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    df.to_excel(EXCEL_PATH, index=False)
    return redirect('/')

@app.route('/clear-data', methods=['POST'])
def clear_data():
    try:
        if os.path.exists(EXCEL_PATH):
            df = pd.read_excel(EXCEL_PATH)
            empty_df = pd.DataFrame(columns=df.columns)
            empty_df.to_excel(EXCEL_PATH, index=False)
        else:
            pd.DataFrame().to_excel(EXCEL_PATH, index=False)
        return jsonify({"status": "ok", "message": "All data cleared successfully."})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

@app.route('/generate-docx')
def generate_docx():
    global process_running, logs
    if process_running:
        return jsonify({"status": "already_running"})

    logs = []
    process_running = True

    threading.Thread(target=process_all, args=(logs,), daemon=True).start()
    return jsonify({"status": "started"})

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

@app.route('/download/<filename>')
def download_file(filename):
    safe_name = os.path.basename(filename)
    file_path = os.path.join(OUTPUT_DIR, safe_name)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name=safe_name)
    else:
        abort(404)

# ==========================================================
# ROUTE: Submit Certificate (Updated as per requirements)
# ==========================================================
@app.route('/submit-certificate', methods=['GET', 'POST'])
def submit_certificate():
    if request.method == 'GET':
        return render_template('submit-certificate.html')

    try:
        data = request.form.to_dict()

        template_name = data.get('template', 'Page 1') + '.docx'
        template_path = os.path.join(SPECIMEN_DIR, template_name)

        if not os.path.exists(template_path):
            return render_template('submit-certificate.html', message="❌ Template not found.")

        # Next available page number
        existing_files = [f for f in os.listdir(EXP1_FOLDER) if f.startswith('Page (') and f.endswith('.docx')]
        max_page = 0
        for f in existing_files:
            try:
                num = int(f.split('(')[1].split(')')[0])
                max_page = max(max_page, num)
            except:
                continue
        next_page_num = max_page + 1
        next_filename = f"Page ({next_page_num}).docx"
        output_path = os.path.join(EXP1_FOLDER, next_filename)

        # Unzip template, replace placeholders
        with tempfile.TemporaryDirectory() as temp_dir:
            with zipfile.ZipFile(template_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
            parser = etree.XMLParser(ns_clean=True, recover=True)
            tree = etree.parse(document_xml_path, parser)
            root = tree.getroot()

            NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            text_nodes = root.xpath('.//w:t', namespaces=NS)

            paragraphs = {}
            for t in text_nodes:
                p = t.getparent()
                while p is not None and p.tag != f"{{{NS['w']}}}p":
                    p = p.getparent()
                if p is None:
                    continue
                paragraphs.setdefault(p, []).append(t)

            def replace_placeholders_in_runs(runs, replacements):
                full_text = ''.join([r.text or '' for r in runs])
                for key, val in replacements.items():
                    placeholder = f"[{key}]"
                    if placeholder in full_text:
                        start = full_text.index(placeholder)
                        end = start + len(placeholder)
                        new_text = full_text[:start] + val + full_text[end:]
                        runs[0].text = new_text
                        for run in runs[1:]:
                            run.text = ''
                        full_text = new_text

            for para, runs in paragraphs.items():
                replace_placeholders_in_runs(runs, data)

            tree.write(document_xml_path, xml_declaration=True, encoding='UTF-8', standalone="yes")

            shutil.make_archive('modified_doc', 'zip', temp_dir)
            modified_zip = 'modified_doc.zip'

            if os.path.exists(output_path):
                os.remove(output_path)
            shutil.move(modified_zip, output_path)

        # Update Excel
        if not os.path.exists(EXP_MANUAL_PATH):
            return render_template('submit-certificate.html', message="❌ Excel file not found.")

        wb = load_workbook(EXP_MANUAL_PATH)
        ws = wb.active

        new_row = ws.max_row + 1
        ws.cell(row=new_row, column=1, value=f"Page ({next_page_num})")
        ws.cell(row=new_row, column=2, value=data.get('companyname', ''))
        ws.cell(row=new_row, column=3, value=data.get('companyproject', ''))
        ws.cell(row=new_row, column=4, value=data.get('country', ''))

        wb.save(EXP_MANUAL_PATH)

        return render_template('submit-certificate.html', message=f"✅ Certificate saved as <b>{next_filename}</b> and Excel updated.")

    except Exception as e:
        return render_template('submit-certificate.html', message=f"❌ Error: {str(e)}")

# ==========================================================
# ROUTE: Photocopies Page (static template rendering)
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
