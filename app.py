from flask import Flask, render_template, request, send_file
import os

from formatter import (
    normalize_marks,
    generate_part_tables,
    export_excel,
    export_pdf
)

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
    return send_file("static/marks_template.xlsx", as_attachment=True)


@app.route("/process", methods=["POST"])
def process():

    college = request.form["college"]
    course_code = request.form["course_code"]
    course_name = request.form["course_name"]
    semester = request.form["semester"]
    year = request.form["year"]

    file = request.files["marks_file"]

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    normalized_df, co_columns, eval_methods, max_marks = normalize_marks(filepath)

    part1_df, part2_df = generate_part_tables(normalized_df, co_columns)

    excel_file = os.path.join(OUTPUT_FOLDER, "MGU_CCA_Output.xlsx")
    pdf_file = os.path.join(OUTPUT_FOLDER, "MGU_CCA_Output.pdf")

    export_excel_template(
    part1_df,
    part2_df,
    college,
    course_code,
    course_name,
    semester,
    year,
    "static/mgu_cca_template.xlsx",
    excel_file
	)
    export_pdf(part2_df, pdf_file)

    return render_template(
        "preview.html",
        normalized=normalized_df.to_html(index=False),
        part2=part2_df.to_html(index=False),
        college=college,
        course_code=course_code,
        course_name=course_name,
        semester=semester,
        year=year
    )


@app.route("/download")
def download():
    return send_file("outputs/MGU_CCA_Output.xlsx", as_attachment=True)


@app.route("/download_pdf")
def download_pdf():
    return send_file("outputs/MGU_CCA_Output.pdf", as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
