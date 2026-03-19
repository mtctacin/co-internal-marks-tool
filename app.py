from flask import Flask, render_template, request, send_file
import os

from formatter import process_file

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/download-template")
def download_template():
    return send_file("static/marks_template.xlsx", as_attachment=True)


@app.route("/process", methods=["POST"])
def process():

    form_data = {
        "college": request.form["college"],
        "course_code": request.form["course_code"],
        "course_name": request.form["course_name"],
        "year": request.form["year"],
        "semester": request.form["semester"],
        "category": request.form["category"]
    }

    file = request.files["marks_file"]
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    result = process_file(path)

    return render_template(
        "preview.html",
        table2=result["table2"].to_html(index=False),
        co_columns=result["co_columns"],
        max_marks=result["max_marks"],
        num_cos=len(result["co_columns"]),
        num_students=result["num_students"],
        form=form_data
    )


if __name__ == "__main__":
    app.run(debug=True)
