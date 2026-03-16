from flask import Flask, render_template, request, send_file
import pandas as pd
import os

from formatter import normalize_marks, generate_mgu_cca

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/download-template")
def download_template():
    return send_file(
        "static/marks_template.xlsx",
        as_attachment=True
    )


@app.route("/process", methods=["POST"])
def process():

    college = request.form["college"]
    course_code = request.form["course_code"]
    semester = request.form["semester"]
    year = request.form["year"]
    course_name = request.form["course_name"]

    file = request.files["marks_file"]

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    # Normalization module
    normalized_df, co_columns, eval_methods, max_marks = normalize_marks(filepath)

    output_file = os.path.join(OUTPUT_FOLDER, "MGU_CCA_Output.xlsx")

    generate_mgu_cca(
        normalized_df,
        co_columns,
        eval_methods,
        max_marks,
        college,
        course_code,
        semester,
        year,
        course_name,
        output_file
    )

    return send_file(output_file, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
