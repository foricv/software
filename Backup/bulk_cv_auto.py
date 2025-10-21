import os
import re
import random
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from docx import Document
from PyPDF2 import PdfMerger
import subprocess
import shutil

# ==========================================================
# PATHS
# ==========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MAIN_PATH = os.path.join(BASE_DIR, "data", "MainData.xlsx")
EXP_AUTO_PATH = os.path.join(BASE_DIR, "data", "ExpAuto.xlsx")
EXP_MANUAL_PATH = os.path.join(BASE_DIR, "data", "ExpManual.xlsx")
UPDATED_PATH = os.path.join(BASE_DIR, "data", "MainData_Updated.xlsx")

EXP1_FOLDER = os.path.join(BASE_DIR, "experience", "Exp1")
EXP2_FOLDER = os.path.join(BASE_DIR, "experience", "Exp2")
CV_FOLDER   = os.path.join(BASE_DIR, "cv_templates")
TEMP_FOLDER = os.path.join(BASE_DIR, "temp")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "data", "output")

CCC_TEMPLATE = os.path.join(BASE_DIR, "data", "ccc_template.docx")

os.makedirs(TEMP_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ==========================================================
# CASE-INSENSITIVE FILE FINDER ✅
# ==========================================================
def find_file_case_insensitive(folder, filename):
    """Find file in folder ignoring case, and auto-handle missing .docx"""
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
# LIBREOFFICE CONVERSION HELPER
# ==========================================================
def docx_to_pdf_with_libreoffice(docx_path, out_dir):
    """Convert docx -> pdf using LibreOffice CLI"""
    if not os.path.exists(docx_path):
        return None
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("[WARN] LibreOffice not found on PATH — skipping PDF conversion.")
        return None
    try:
        os.makedirs(out_dir, exist_ok=True)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
        )
        pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        return pdf_path if os.path.exists(pdf_path) else None
    except Exception as e:
        print(f"[ERROR] LibreOffice conversion failed: {e}")
        return None


# ==========================================================
# STEP 1: Adjust Dates
# ==========================================================
def adjust_dates():
    print("[Step 1] Adjusting dates and experience info...")

    try:
        exp_auto_samples = pd.read_excel(EXP_AUTO_PATH)
    except Exception as e:
        print(f"[ERROR] Could not read ExpAuto.xlsx: {e}")
        exp_auto_samples = pd.DataFrame()

    try:
        exp_manual_samples = pd.read_excel(EXP_MANUAL_PATH)
    except Exception as e:
        print(f"[ERROR] Could not read ExpManual.xlsx: {e}")
        exp_manual_samples = pd.DataFrame()

    min_exp_years, max_exp_years = 2, 3
    min_gap_months, max_gap_months = 12, 36
    date_format = "%d-%m-%Y"
    today = datetime.today()

    df_main = pd.read_excel(MAIN_PATH).replace("nan", pd.NA).astype("object")

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
            print(f"[WARN] Row {i+1}: invalid DOB '{dob_raw}' — skipped")
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
        except Exception as ex:
            print(f"[WARN] Row {i+1}: Could not sample experience — {ex}")

    df_main.to_excel(UPDATED_PATH, index=False)
    print(f"[OK] Updated Excel saved at: {UPDATED_PATH}")


# ==========================================================
# STEP 2: Replace Placeholders
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
    except Exception as e:
        print(f"Error processing {doc_path}: {e}")


# ==========================================================
# STEP 3: Merge and Export PDFs
# ==========================================================
def merge_pdfs(pdf_list, output_path):
    merger = PdfMerger()
    for pdf in pdf_list:
        if os.path.exists(pdf):
            merger.append(pdf)
    merger.write(output_path)
    merger.close()


# ==========================================================
# STEP 4: MAIN PDF MAKER (Case-insensitive search fixed ✅)
# ==========================================================
def make_pdfs():
    print("[Step 2] Filling templates, converting to PDF, and adding CCC if needed…")

    try:
        df = pd.read_excel(UPDATED_PATH)
    except Exception as e:
        print(f"[ERROR] Could not read Excel: {e}")
        return

    for idx, row in df.iterrows():
        name = str(row.get("name", "Unknown")).strip()
        print(f"[{idx + 1}/{len(df)}] Preparing docs for: {name}")

        exp1_file = find_file_case_insensitive(EXP1_FOLDER, str(row.get("exp1", "")))
        exp2_file = find_file_case_insensitive(EXP2_FOLDER, str(row.get("exp2", "")))
        cv_file   = find_file_case_insensitive(CV_FOLDER, str(row.get("cv", "")))

        temp_dir = os.path.join(TEMP_FOLDER, name)
        os.makedirs(temp_dir, exist_ok=True)

        replacements = {col: str(val) for col, val in row.items() if pd.notna(val)}

        for src in [exp1_file, exp2_file, cv_file]:
            if src and os.path.exists(src):
                dest = os.path.join(temp_dir, os.path.basename(src))
                replace_placeholders(src, replacements, dest)
            else:
                print(f"[WARN] Missing source file: {src or 'Unknown'}")

        # Convert all DOCXs in temp folder → PDFs
        pdf_files = []
        for docx in os.listdir(temp_dir):
            if docx.lower().endswith(".docx"):
                pdf = docx_to_pdf_with_libreoffice(os.path.join(temp_dir, docx), temp_dir)
                if pdf:
                    pdf_files.append(pdf)

        if not pdf_files:
            print(f"[WARN] No PDFs to merge for: {name}")
            continue

        output_pdf = os.path.join(OUTPUT_FOLDER, f"{name}.pdf")
        merge_pdfs(pdf_files, output_pdf)
        print(f"[OK] PDF created: {output_pdf}")

        shutil.rmtree(temp_dir, ignore_errors=True)

    print("[DONE] All documents processed and temp cleaned up.")


# ==========================================================
# RUN
# ==========================================================
if __name__ == "__main__":
    adjust_dates()
    make_pdfs()
