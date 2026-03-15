import os
import io
import time
import uuid
import tempfile
import zipfile
import re

from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import pandas as pd
from weasyprint import HTML

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "super-secret-key"

GENERATED_ZIPS = {}


# ----------------------------
# Safe filename
# ----------------------------
def sanitize_filename(name: str):
    name = re.sub(r"[\\/]+", "-", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = re.sub(r"[:\*\?\"<>\|]+", "", name)
    return name or "file"


# ----------------------------
# Date formatter
# ----------------------------
def format_date(value, format_type="month"):

    if pd.isna(value) or str(value).strip() == "":
        return ""

    try:
        if format_type == "month":
            return pd.to_datetime(value).strftime("%B %Y")
        elif format_type == "issue":
            return pd.to_datetime(value).strftime("%d-%m-%Y")
    except:
        return str(value)


# ----------------------------
# Check missing values
# ----------------------------
def is_missing(value):

    if pd.isna(value):
        return True

    if str(value).strip() == "":
        return True

    return False


# ----------------------------
# Filename rendering
# ----------------------------
def render_filename(template, row, idx):

    values = {
        "name": str(row.get("name", "")).strip(),
        "course": str(row.get("course", "")).strip(),
        "index": str(idx),
    }

    result = template

    for k, v in values.items():
        result = result.replace("{" + k + "}", v)

    return sanitize_filename(result)


# ----------------------------
# Create HTML certificate
# ----------------------------
def build_certificate_html(context):

    return f"""
    <html>
    <head>
    <style>
    body {{
        text-align:center;
        font-family:Arial;
        padding-top:150px;
    }}

    h1 {{
        font-size:40px;
    }}

    .name {{
        font-size:46px;
        font-weight:bold;
        margin:20px 0;
    }}

    .course {{
        font-size:26px;
    }}

    .small {{
        font-size:18px;
        margin-top:20px;
    }}

    </style>
    </head>

    <body>

    <h1>Certificate of Completion</h1>

    <p>This is to certify that</p>

    <div class="name">{context['name']}</div>

    <p>has successfully completed</p>

    <div class="course">{context['course']}</div>

    <div class="small">Grade: {context['grade']}</div>

    <div class="small">{context['place']}</div>

    <div class="small">{context['date']}</div>

    <div class="small">{context['issue_date']}</div>

    </body>
    </html>
    """


# ----------------------------
# Main Route
# ----------------------------
@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        start_time = time.perf_counter()

        excel_file = request.files.get("employees_file")

        filename_template = request.form.get(
            "filename_template", "{name} Certificate"
        ).strip()

        if not excel_file:
            flash("Please upload Excel file.")
            return redirect(url_for("index"))

        with tempfile.TemporaryDirectory() as tmpdir:

            excel_path = os.path.join(tmpdir, "data.xlsx")
            excel_file.save(excel_path)

            df = pd.read_excel(excel_path)

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

                name = row["name"]
                course = row["course"]
                grade = row["grade"]
                date_val = row["date"]
                place = row["place"]

                if (
                    is_missing(name)
                    or is_missing(course)
                    or is_missing(grade)
                    or is_missing(date_val)
                    or is_missing(place)
                ):

                    errors += 1

                    report_rows.append(
                        {
                            "name": str(name),
                            "filename": "Not Generated",
                            "status": "Error: Missing required value",
                        }
                    )

                    continue

                name = str(name)
                course = str(course)
                grade = str(grade)
                place = str(place)

                date = format_date(date_val, "month")
                issue_date = format_date(row["issue_date"], "issue")

                context = {
                    "name": name,
                    "course": course,
                    "grade": grade,
                    "date": date,
                    "place": place,
                    "issue_date": issue_date,
                }

                html_content = build_certificate_html(context)

                base_name = render_filename(filename_template, row, i)

                pdf_path = os.path.join(pdf_dir, base_name + ".pdf")

                try:

                    HTML(string=html_content).write_pdf(pdf_path)

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
            }

            return render_template("index.html", result=result)

    return render_template("index.html")


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


if __name__ == "__main__":

    port = int(os.environ.get("PORT", 5000))

    app.run(host="0.0.0.0", port=port)
