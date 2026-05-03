import os
import tempfile
from pathlib import Path

from flask import Flask, after_this_request, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from ams_processor import process_ams_update


app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-this-in-production")
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process_files():
    source_file = request.files.get("source_file")
    input_excel = request.files.get("input_excel")
    selected_class = request.form.get("selected_class", "1")

    if not source_file or not input_excel:
        flash("Please upload both AMS source and Input Excel file.")
        return redirect(url_for("index"))

    source_name = secure_filename(source_file.filename or "")
    excel_name = secure_filename(input_excel.filename or "")
    source_ext = Path(source_name).suffix.lower()
    excel_ext = Path(excel_name).suffix.lower()

    if source_ext not in {".pdf", ".xlsx", ".xlsm"}:
        flash("AMS source must be PDF, XLSX, or XLSM.")
        return redirect(url_for("index"))
    if excel_ext not in {".xlsx", ".xlsm"}:
        flash("Input file must be XLSX or XLSM.")
        return redirect(url_for("index"))

    temp_paths = []
    tmp_root = tempfile.mkdtemp(prefix="ams-web-")

    source_path = os.path.join(tmp_root, source_name)
    excel_path = os.path.join(tmp_root, excel_name)
    temp_paths.extend([source_path, excel_path])

    source_file.save(source_path)
    input_excel.save(excel_path)

    try:
        output_file, _ = process_ams_update(source_path, excel_path, selected_class)
        temp_paths.append(output_file)
    except Exception as exc:
        flash(str(exc))
        return redirect(url_for("index"))

    @after_this_request
    def cleanup_temp_files(response):
        for path in temp_paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except OSError:
                pass
        try:
            os.rmdir(tmp_root)
        except OSError:
            pass
        return response

    return send_file(
        output_file,
        as_attachment=True,
        download_name=Path(output_file).name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
