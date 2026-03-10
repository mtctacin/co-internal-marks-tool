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

        co_data = {}

        for column in df.columns:

            if "_CO" in column:

                component = column.split("_")[0]
                co = column.split("_")[1]

                max_col = f"{component}_Max"

                if max_col not in df.columns:
                    continue

                if co not in co_data:
                    co_data[co] = {
                        "scored": df[column].copy(),
                        "max": df[max_col].copy()
                    }
                else:
                    co_data[co]["scored"] += df[column]
                    co_data[co]["max"] += df[max_col]

        result = df[["RegNo", "Name"]].copy()

        for co in co_data:

            total_scored = co_data[co]["scored"]
            total_max = co_data[co]["max"]

            result[co] = (total_scored / total_max) * TARGET_MAX

        output_path = f"/tmp/result_{uuid.uuid4()}.xlsx"
        result.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run()
