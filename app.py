from flask import Flask, render_template, request, send_file
import pandas as pd
import uuid

app = Flask(__name__)

TARGET_MAX = 10

@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        file = request.files["file"]

        input_path = f"/tmp/{uuid.uuid4()}.xlsx"
        file.save(input_path)

        df = pd.read_excel(input_path)

        # Extract max marks from first row
        max_row = df.iloc[0]

        # Student data starts from second row
        students = df.iloc[1:].copy()

        co_data = {}

        for column in students.columns:

            if "_CO" in column:

                component, co = column.split("_")

                max_mark = max_row[column]

                if co not in co_data:

                    co_data[co] = {
                        "scored": students[column].astype(float),
                        "max": max_mark
                    }

                else:

                    co_data[co]["scored"] += students[column].astype(float)
                    co_data[co]["max"] += max_mark

        result = students[["RegNo", "Name"]].copy()

        for co in co_data:

            scored = co_data[co]["scored"]
            max_total = co_data[co]["max"]

            result[co] = (scored / max_total) * TARGET_MAX

        output_path = f"/tmp/result_{uuid.uuid4()}.xlsx"
        result.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run()
