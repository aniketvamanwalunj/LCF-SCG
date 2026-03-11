import os
import io
import time
import uuid
import tempfile
import zipfile
import re
import subprocess

from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import pandas as pd
from docxtpl import DocxTemplate

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "super-secret-key"

GENERATED_ZIPS = {}

# ----------------------------
# DOCX → PDF using LibreOffice
# ----------------------------
def convert_to_pdf(docx_path, output_dir):

    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_dir,
        docx_path
    ], check=True)


# ----------------------------
# Safe filename
# ----------------------------
def sanitize_filename(name: str):
    name = re.sub(r"[\\/]+", "-", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = re.sub(r"[:\*\?\"<>\|]+", "", name)
    return name or "file"


# ----------------------------
# Render filename
# ----------------------------
def render_filename(template, row, idx):

    try:
        date_val = pd.to_datetime(row.get("date", "")).strftime("%b/%Y")
    except:
        date_val = str(row.get("date", ""))

    try:
        issue_date_val = pd.to_datetime(row.get("issue_date", "")).strftime("%d-%m-%Y")
    except:
        issue_date_val = str(row.get("issue_date", ""))

    values = {
        "name": str(row.get("name", "")).strip(),
        "course": str(row.get("course", "")).strip(),
        "grade": str(row.get("grade", "")).strip(),
        "date": date_val,
        "place": str(row.get("place", "")).strip(),
        "issue_date": issue_date_val,
        "index": str(idx),
    }

    result = template

    for k, v in values.items():
        result = result.replace("{" + k + "}", v)

    return sanitize_filename(result)


# ----------------------------
# Main Route
# ----------------------------
@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        start_time = time.perf_counter()

        excel_file = request.files.get("employees_file")
        template_file = request.files.get("template_file")

        filename_template = request.form.get(
            "filename_template", "{name} - {course} Certificate"
        ).strip()

        if not excel_file or not template_file:
            flash("Please upload Excel file and Certificate Template.")
            return redirect(url_for("index"))

        with tempfile.TemporaryDirectory() as tmpdir:

            excel_path = os.path.join(tmpdir, "data.xlsx")
            template_path = os.path.join(tmpdir, "template.docx")

            excel_file.save(excel_path)
            template_file.save(template_path)

            try:
                df = pd.read_excel(excel_path)
            except Exception as e:
                flash(f"Unable to read Excel file: {e}")
                return redirect(url_for("index"))

            required_cols = [
                "name",
                "course",
                "grade",
                "date",
                "place",
                "issue_date",
            ]

            missing = [c for c in required_cols if c not in df.columns]

            if missing:
                flash(f"Missing required Excel columns: {', '.join(missing)}")
                return redirect(url_for("index"))

            pdf_dir = os.path.join(tmpdir, "PDF")
            os.makedirs(pdf_dir, exist_ok=True)

            success = 0
            errors = 0
            report_rows = []

            for i, (_, row) in enumerate(df.iterrows(), start=1):

                name = str(row["name"])
                course = str(row["course"])
                grade = str(row["grade"])
                place = str(row["place"])

                # Format date -> Feb/2026
                try:
                    date = pd.to_datetime(row["date"]).strftime("%b/%Y")
                except:
                    date = str(row["date"])

                try:
                    issue_date = pd.to_datetime(row["issue_date"]).strftime("%d-%m-%Y")
                except:
                    issue_date = str(row["issue_date"])

                doc = DocxTemplate(template_path)

                context = {
                    "name": name,
                    "course": course,
                    "grade": grade,
                    "date": date,
                    "place": place,
                    "issue_date": issue_date,
                }

                base_name = render_filename(filename_template, row, i)

                docx_path = os.path.join(tmpdir, base_name + ".docx")

                try:

                    doc.render(context)
                    doc.save(docx_path)

                    # Convert DOCX → PDF
                    convert_to_pdf(docx_path, pdf_dir)

                    success += 1
                    status = "Success"

                except Exception as e:

                    errors += 1
                    status = f"Error: {e}"

                report_rows.append(
                    {
                        "name": name,
                        "filename": base_name + ".pdf",
                        "status": status,
                    }
                )

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:

                for f in os.listdir(pdf_dir):
                    zipf.write(os.path.join(pdf_dir, f), f)

                report_df = pd.DataFrame(report_rows)

                report_io = io.BytesIO()

                with pd.ExcelWriter(report_io, engine="openpyxl") as writer:
                    report_df.to_excel(writer, index=False)

                report_io.seek(0)

                zipf.writestr("Certificate_Report.xlsx", report_io.read())

            zip_buffer.seek(0)

            zip_id = str(uuid.uuid4())
            GENERATED_ZIPS[zip_id] = zip_buffer

            elapsed = time.perf_counter() - start_time

            result = {
                "total": len(df),
                "success": success,
                "error": errors,
                "time": f"{elapsed:.2f} sec",
                "zip_id": zip_id,
                "filename_template": filename_template,
            }

            return render_template("index.html", result=result)

    return render_template("index.html")


# ----------------------------
# Download ZIP
# ----------------------------
@app.route("/download/<zip_id>")
def download_zip(zip_id):

    if zip_id not in GENERATED_ZIPS:
        return "Invalid download link", 404

    GENERATED_ZIPS[zip_id].seek(0)

    return send_file(
        GENERATED_ZIPS[zip_id],
        as_attachment=True,
        download_name="Certificates.zip",
        mimetype="application/zip",
    )


# ----------------------------
# Start Server
# ----------------------------
if __name__ == "__main__":

    port = int(os.environ.get("PORT", 5000))

    app.run(host="0.0.0.0", port=port)
