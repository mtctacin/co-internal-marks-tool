from flask import Flask, render_template, request
import os

from formatter import normalize_marks, build_tables

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():

    file = request.files["marks_file"]

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    table1, table2 = normalize_marks(filepath)
    table2 = build_tables(table2)

    return render_template(
        "preview.html",
        table1=table1.to_html(index=False),
        table2=table2.to_html(index=False)
    )


if __name__ == "__main__":
    app.run(debug=True)
