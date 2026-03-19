from flask import Flask, render_template, request, send_from_directory
import os

from formatter import process_file

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


# ✅ Download Template
@app.route("/download-template")
def download_template():
    return send_from_directory(
        os.path.join(app.root_path, "static"),
        "marks_template.xlsx",
        as_attachment=True
    )


# ✅ Process File
@app.route("/process", methods=["POST"])
def process():
    try:
        # Form data
        form_data = {
            "college": request.form["college"],
            "course_code": request.form["course_code"],
            "course_name": request.form["course_name"],
            "year": request.form["year"],
            "semester": request.form["semester"],
            "category": request.form["category"]
        }

        file = request.files.get("marks_file")

        if not file or file.filename == "":
            return "<h3>No file uploaded</h3>"

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        result = process_file(filepath)

        return render_template(
            "preview.html",
            table1=result["table1"].to_html(index=False),
            table2=result["table2"],
            co_columns=result["co_columns"],
            max_marks=result["max_marks"],
            eval_list=result["eval_list"],
            total_cca_max=result["total_cca_max"],
            num_cos=len(result["co_columns"]),
            num_students=result["num_students"],
            form=form_data
        )

    except Exception as e:
        return f"<h3>Error:</h3><pre>{str(e)}</pre>"


if __name__ == "__main__":
    app.run(debug=True)
