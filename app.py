from flask import Flask, render_template, request, send_file
import pandas as pd
import matplotlib.pyplot as plt
import uuid
import os

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():

    table_html = None
    chart_path = None
    download_file = None

    if request.method == "POST":

        file = request.files["file"]
        path = f"/tmp/{uuid.uuid4()}.xlsx"
        file.save(path)

        df = pd.read_excel(path)

        max_row = df.iloc[0]
        norm_row = df.iloc[1]
        students = df.iloc[2:].copy()

        result = students[["RegNo","Name"]].copy()

        co_totals = {}
        co_max_marks = {}

        for col in students.columns:

            if "_CO" in col:

            max_mark = max_row[col]
            norm_mark = norm_row[col]

            norm_col = col + "_N"

            result[norm_col] = (
                (students[col].astype(float) / max_mark) * norm_mark
            ).round(2)

            co = col.split("_")[1]

            if co not in co_totals:

                co_totals[co] = result[norm_col]
                co_max_marks[co] = norm_mark

            else:

                co_totals[co] += result[norm_col]
                co_max_marks[co] += norm_mark

        for co in co_totals:
            result[co] = co_totals[co].round(2)

        # CO SUMMARY
        summary = pd.DataFrame({
            "CO": list(co_totals.keys()),
            "Max Marks": [co_max_marks[c] for c in co_totals],
            "Average": [round(co_totals[c].mean(),2) for c in co_totals],
            "Highest": [round(co_totals[c].max(),2) for c in co_totals],
            "Lowest": [round(co_totals[c].min(),2) for c in co_totals]
        })
        
        # CHART
        plt.figure()

        x = summary["CO"]
        avg = summary["Average"]
        max_marks = summary["Max Marks"]

        plt.bar(x, avg)

        for i in range(len(x)):
            plt.text(i, avg[i] + 0.2, f"{avg[i]}/{max_marks[i]}", ha='center')

        plt.ylabel("Average Marks")
        plt.title("CO Performance (Average / Max)")

        chart_path = "static/chart.png"
        plt.savefig(chart_path)
        plt.close()

        # CREATE EXCEL
        output = f"/tmp/result_{uuid.uuid4()}.xlsx"

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result.to_excel(writer, sheet_name="Student Marks", index=False)
            summary.to_excel(writer, sheet_name="CO Summary", index=False)

        download_file = output

        table_html = result.to_html(classes="table", index=False)

    return render_template(
        "index.html",
        table=table_html,
        chart=chart_path,
        download=download_file
    )


@app.route("/download")
def download():
    path = request.args.get("file")
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    app.run()

@app.route("/download-template")
def download_template():
    return send_file("static/marks_template.xlsx", as_attachment=True)
