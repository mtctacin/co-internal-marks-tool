import pandas as pd
from openpyxl import Workbook


def normalize_marks(file_path):

    df = pd.read_excel(file_path)

    raw_max = df.iloc[0]
    norm_max = df.iloc[1]

    students = df.iloc[2:].reset_index(drop=True)

    co_map = {}
    eval_methods = {}

    for col in df.columns:

        if "_CO" in col:

            comp, co = col.split("_")

            co_map.setdefault(co, []).append(col)

            eval_methods.setdefault(co, []).append(comp)

    co_columns = sorted(co_map.keys())

    normalized = pd.DataFrame()
    normalized["RegNo"] = students["RegNo"]
    normalized["Name"] = students["Name"]

    max_marks = []

    for co in co_columns:

        total = 0

        for col in co_map[co]:

            total += (students[col] / raw_max[col]) * norm_max[col]

        normalized[co] = total.round(2)

        max_marks.append(sum(norm_max[c] for c in co_map[co]))

    eval_list = [", ".join(eval_methods[co]) for co in co_columns]

    return normalized, co_columns, eval_list, max_marks


def generate_mgu_cca(df, co_columns, eval_methods, max_marks,
                     college, course_code, semester, year,
                     course_name, output_file):

    wb = Workbook()

    ws1 = wb.active
    ws1.title = "CCA Form A Part I"

    ws2 = wb.create_sheet("CCA Form A Part II")

    # ---------- HEADER ----------

    ws1["A1"] = "MAHATMA GANDHI UNIVERSITY, KOTTAYAM"
    ws2["A1"] = "MAHATMA GANDHI UNIVERSITY, KOTTAYAM"

    ws1["A2"] = "Name of the College:"
    ws1["C2"] = college
    ws1["K2"] = "Academic Year"
    ws1["L2"] = year

    ws2["A2"] = "Name of the College:"
    ws2["C2"] = college
    ws2["K2"] = "Academic Year"
    ws2["L2"] = year

    ws1["A3"] = "Course Code:"
    ws1["C3"] = course_code
    ws1["H3"] = "Name of the Course:"
    ws1["K3"] = course_name

    ws2["A3"] = "Course Code:"
    ws2["C3"] = course_code
    ws2["H3"] = "Name of the Course:"
    ws2["K3"] = course_name

    ws1["A4"] = "Semester:"
    ws1["C4"] = semester

    ws2["A4"] = "Semester:"
    ws2["C4"] = semester

    ws1["H4"] = "Number of Course Outcomes:"
    ws1["L4"] = len(co_columns)

    ws2["H4"] = "Number of Course Outcomes:"
    ws2["L4"] = len(co_columns)

    ws1["H5"] = "Total Number of Students Enrolled for this Course:"
    ws1["L5"] = len(df)

    ws2["H5"] = "Total Number of Students Enrolled for this Course:"
    ws2["L5"] = len(df)

    # ---------- CO HEADER ----------

    start_col = 4

    ws2.cell(row=6, column=1).value = "CO Number"
    ws2.cell(row=7, column=1).value = "Evaluation Method(s) Used"
    ws2.cell(row=8, column=1).value = "Maximum Marks Allocated"

    for i, co in enumerate(co_columns):

        ws2.cell(row=6, column=start_col+i).value = co
        ws2.cell(row=7, column=start_col+i).value = eval_methods[i]
        ws2.cell(row=8, column=start_col+i).value = max_marks[i]

    ws2.cell(row=6, column=start_col+len(co_columns)).value = "Total Marks Allocated for CCA"
    ws2.cell(row=8, column=start_col+len(co_columns)).value = sum(max_marks)

    # ---------- STUDENTS ----------

    start_row = 10

    for i, row in df.iterrows():

        r = start_row + i

        ws1.cell(r,1).value = i+1
        ws1.cell(r,2).value = row["RegNo"]
        ws1.cell(r,3).value = row["Name"]

        ws2.cell(r,1).value = i+1
        ws2.cell(r,2).value = row["RegNo"]
        ws2.cell(r,3).value = row["Name"]

        for j, co in enumerate(co_columns):

            ws1.cell(r,6+(2*j)).value = row[co]
            ws2.cell(r,4+j).value = row[co]

    wb.save(output_file)
