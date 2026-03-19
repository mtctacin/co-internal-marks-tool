import pandas as pd


def normalize_marks(file_path):

    df = pd.read_excel(file_path)

    raw_max = df.iloc[0]
    norm_max = df.iloc[1]

    students = df.iloc[2:].reset_index(drop=True)

    # -------- TABLE 1 (Component-wise normalized) --------

    table1 = pd.DataFrame()
    table1["Register Number"] = students["RegNo"]
    table1["Name of the Student"] = students["Name"]

    co_map = {}

    for col in df.columns:

        if "_CO" in col:

            comp, co = col.split("_")

            normalized_col = (students[col] / raw_max[col]) * norm_max[col]

            table1[col] = normalized_col.round(2)

            co_map.setdefault(co, []).append(col)

    # -------- TABLE 2 (CO totals) --------

    table2 = pd.DataFrame()
    table2["Register Number"] = students["RegNo"]
    table2["Name of the Student"] = students["Name"]

    co_columns = sorted(co_map.keys())

    for co in co_columns:

        total = 0

        for col in co_map[co]:

            total += (students[col] / raw_max[col]) * norm_max[col]

        table2[f"{co}"] = total.round(2)

    table2["Total Marks Obtained in CCA"] = table2[co_columns].sum(axis=1).round(2)

    return table1, table2


def build_tables(table2):

    table2.insert(0, "SL. No.", range(1, len(table2)+1))

    # Rename CO columns nicely
    for col in table2.columns:
        if col.startswith("CO"):
            table2.rename(columns={
                col: f"Marks Obtained in {col}"
            }, inplace=True)

    return table2
