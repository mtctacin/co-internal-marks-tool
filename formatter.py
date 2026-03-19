import pandas as pd


def process_file(file_path):

    df = pd.read_excel(file_path)

    raw_max = df.iloc[0]
    norm_max = df.iloc[1]
    students = df.iloc[2:].reset_index(drop=True)

    co_map = {}

    for col in df.columns:
        if "_CO" in col:
            comp, co = col.split("_")
            co_map.setdefault(co, []).append(col)

    co_columns = sorted(co_map.keys())

    table2 = pd.DataFrame()
    table2["SL. No."] = range(1, len(students)+1)
    table2["Register Number"] = students["RegNo"]
    table2["Name of the Student"] = students["Name"]

    max_marks = []

    for co in co_columns:

        total = 0
        co_max = 0

        for col in co_map[co]:
            total += (students[col] / raw_max[col]) * norm_max[col]
            co_max += norm_max[col]

        table2[f"Marks Obtained in {co}"] = total.round(2)
        max_marks.append(co_max)

    table2["Total Marks Obtained in CCA"] = table2[
        [f"Marks Obtained in {co}" for co in co_columns]
    ].sum(axis=1).round(2)

    return {
        "table2": table2,
        "co_columns": co_columns,
        "max_marks": max_marks,
        "num_students": len(students)
    }
