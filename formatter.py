import pandas as pd


def process_file(file_path):

    df = pd.read_excel(file_path)

    # ✅ Validate structure
    if df.shape[0] < 3:
        raise ValueError("Invalid template. Must contain MAX, Normalised, and student rows.")

    # ✅ Correct row mapping (based on your template)
    raw_max = df.iloc[0]        # MAX row
    norm_max = df.iloc[1]       # Normalised row
    students = df.iloc[2:].reset_index(drop=True)

    co_map = {}
    eval_methods = {}

    # ✅ Detect COs and evaluation methods
    for col in df.columns:
        if "_CO" in col:

            comp, co = col.split("_")

            co_map.setdefault(co, []).append(col)
            eval_methods.setdefault(co, []).append(comp)

    if not co_map:
        raise ValueError("No CO columns found. Ensure columns like Test1_CO1 exist.")

    co_columns = sorted(co_map.keys())

    # ---------------- TABLE 1 ----------------
    table1 = pd.DataFrame()

    table1["Register Number"] = students["RegNo"]
    table1["Name of the Student"] = students["Name"]

    for col in df.columns:
        if "_CO" in col:
            table1[col] = ((students[col] / raw_max[col]) * norm_max[col]).round(2)

    # ---------------- TABLE 2 ----------------
    table2 = pd.DataFrame()

    table2["SL. No."] = range(1, len(students) + 1)
    table2["Register Number"] = students["RegNo"]
    table2["Name of the Student"] = students["Name"]

    max_marks = []
    eval_list = []

    for co in co_columns:

        total = 0
        co_max = 0

        for col in co_map[co]:
            total += (students[col] / raw_max[col]) * norm_max[col]
            co_max += norm_max[col]

        table2[f"Marks Obtained in {co}"] = total.round(2)

        max_marks.append(co_max)
        eval_list.append(", ".join(eval_methods[co]))

    # Total CCA
    table2["Total Marks Obtained in CCA"] = table2[
        [f"Marks Obtained in {co}" for co in co_columns]
    ].sum(axis=1).round(2)

    total_cca_max = sum(max_marks)

    return {
        "table1": table1,
        "table2": table2,
        "co_columns": co_columns,
        "max_marks": max_marks,
        "eval_list": eval_list,
        "total_cca_max": total_cca_max,
        "num_students": len(students)
    }
