# CCA / Internal Marks Normalization Tool

## Overview

This tool helps teachers prepare **Internal Assessment / Continuous Comprehensive Assessment (CCA) forms** for courses that contain multiple **Course Outcomes (COs)** and several **formative or summative assessment components** mapped to those COs.

In many cases, assessments are conducted for **larger raw marks** (for example, a test conducted for **30 marks**), while the marks must be entered in the official **CCA form in a smaller scale** (for example, **4 marks**).

This tool automatically **normalizes the raw marks to the required CCA scale** and produces an output that can be directly used while preparing the **final CCA statement**.

---

## Features

* Supports **multiple Course Outcomes (COs)**
* Handles **multiple assessment components per CO**
* Automatically **normalizes marks to the required scale**
* Reduces **manual calculation errors**
* Generates **student-wise normalized marks**
* Upload Excel file with student marks  
* Generate:
  - **Part I – Normalized Marks (Component-wise)**
  - **Part II – CCA Form A Report**
* Clean, printable report format (University style)
* Download reports as **PDF (via browser print)**
* Simple and user-friendly interface

---

## How Normalization Works

If an assessment is conducted for **Raw Marks** and must be converted to **CCA Marks**, the tool uses the formula:

```
Normalized Mark = (Student Mark / Maximum Raw Mark) × Maximum CCA Mark
```

Example:

| Raw Test Marks | Raw Max | CCA Max | Normalized Mark |
| -------------- | ------- | ------- | --------------- |
| 24             | 30      | 4       | 3.20            |

---

## Workflow

1. Download the **Excel template**.
2. Enter the following details:

   * Student details
   * Raw marks obtained
   * Maximum marks for each component
3. Upload the completed Excel file into the tool.
4. The tool will:

   * Normalize all marks
   * Generate **CO-wise results**
   * Prepare a **CCA-ready format**
5. Download the processed reports as **PDF (via browser print)**.

---

## Benefits

* Saves time while preparing CCA forms
* Eliminates manual normalization errors
* Handles large student datasets easily
* Produces **ready-to-use output for official records**

---

## Typical Use Case

Example course structure:

| CO  | Assessment | Raw Marks | CCA Marks |
| --- | ---------- | --------- | --------- |
| CO1 | Test       | 30        | 4         |
| CO2 | Assignment | 20        | 3         |
| CO3 | Quiz       | 10        | 2         |

The tool automatically converts the raw marks to the corresponding **CCA scale**.

---

## Technologies Used

* Python
* Pandas
* HTML / JavaScript

---

## Future Improvements

* **Graphical CO attainment analysis**
* **Batch processing for multiple courses**
* **PDF report generation**

---

## License

This project is intended for **academic use** by teachers preparing internal assessment records.
