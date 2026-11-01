::: {align="center"}
# 🎓 STUDENT MARKS ANALYSIS

### Excel Data Lookup & Filtering Project

**Organising student scores and practising Excel data-analysis
functions**

![Microsoft
Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Project](https://img.shields.io/badge/Project-Student%20Data%20Analysis-2563EB?style=for-the-badge)
![Level](https://img.shields.io/badge/Level-Beginner-16A34A?style=for-the-badge)
:::

------------------------------------------------------------------------

## 👨‍💻 Author

**Abhishek Choudhary**

## 📌 Project Overview

This project uses Microsoft Excel to organise and explore student marks
across three subjects: Mathematics, Science and English. The workbook
contains student IDs, names, subject-wise scores, and dedicated columns
for practising the `VLOOKUP` and `FILTER` functions.

The project is designed as a practical beginner-level exercise in
spreadsheet management, data retrieval and conditional filtering.

## 🎯 Project Objectives

-   Organise student information in a structured Excel table.
-   Understand how student IDs can be used to retrieve student details.
-   Practise the `VLOOKUP` function to search for a student's name.
-   Practise the `FILTER` function to display records that meet selected
    conditions.
-   Explore subject-wise marks and identify students based on score
    criteria.
-   Improve accuracy when writing formulas and selecting data ranges.
-   Develop basic Excel skills useful for data analysis.

## ❓ Problem Statements

The workbook is intended to help answer these practical questions:

1.  How can student records be stored in a structured table?
2.  How can a student's name be retrieved using their Student ID?
3.  How can all students scoring above 50 in Mathematics be displayed?
4.  How can students be filtered based on Science or English marks?
5.  How can records be extracted when more than one condition is
    applied?
6.  How can the result of a lookup be checked against the original
    student table?

## 🗂️ Dataset Description

The workbook `PR3.xlsx` contains one worksheet named **Sheet1** with 20
student records.

  Column       Description
  ------------ ----------------------------------------------------
  Student ID   Unique ID assigned to each student
  Name         Student name
  Math         Mathematics marks
  Science      Science marks
  English      English marks
  VLOOKUP      Practice column for retrieving student information
  FILTER       Practice column for displaying filtered records

**Dataset size:** 20 students\
**Subjects:** Mathematics, Science and English

## 🧮 Excel Functions Practised

### 1. VLOOKUP --- Find a Student Name

Use a Student ID to retrieve the matching student's name.

``` excel
=VLOOKUP(A7,$A$2:$E$21,2,FALSE)
```

-   `A7` is the Student ID to search for.
-   `$A$2:$E$21` is the student data range.
-   `2` returns the value from the second column (Name).
-   `FALSE` requests an exact match.

### 2. FILTER --- Show Students Scoring Above 50 in Mathematics

``` excel
=FILTER(A2:G21,C2:C21>50,"No records found")
```

This returns the complete rows for students whose Mathematics marks are
greater than 50.

### 3. FILTER --- Show Students Scoring Above 50 in Mathematics and Science

``` excel
=FILTER(A2:G21,(C2:C21>50)*(D2:D21>50),"No records found")
```

The `*` symbol combines conditions using **AND**.

### 4. FILTER --- Show Students Scoring Above 80 in Mathematics OR English

``` excel
=FILTER(A2:G21,(C2:C21>80)+(E2:E21>80),"No records found")
```

The `+` symbol combines conditions using **OR**.

> **Compatibility:** `FILTER` is available in Microsoft 365 and
> supported newer Excel versions, including Excel 2021 and later. Its
> results spill into neighbouring cells, so keep the output area clear.

## 🔍 Suggested Analysis Tasks

Try completing these tasks in Excel:

  -----------------------------------------------------------------------
                                       No. Task
  ---------------------------------------- ------------------------------
                                         1 Find the name of the student
                                           with ID 1007 using VLOOKUP.

                                         2 Display all students who
                                           scored above 50 in
                                           Mathematics.

                                         3 Display all students who
                                           scored 75 or more in Science.

                                         4 Find students who scored above
                                           80 in both Mathematics and
                                           English.

                                         5 Find students who scored below
                                           50 in Mathematics.

                                         6 Display students who scored
                                           above 90 in any one of the
                                           three subjects.

                                         7 Retrieve a student's name
                                           using an exact Student ID
                                           match.
  -----------------------------------------------------------------------

## 🛠️ Tools Used

-   **Microsoft Excel** --- data organisation and formula practice.
-   **VLOOKUP** --- retrieving student details by ID.
-   **FILTER** --- extracting records based on conditions.
-   **Markdown** --- project documentation.

## 📈 Expected Learning Outcomes

After completing this project, a learner should be able to:

-   Understand a simple student dataset.
-   Use exact-match VLOOKUP to retrieve information.
-   Apply FILTER with one or multiple conditions.
-   Understand AND (`*`) and OR (`+`) logic in FILTER formulas.
-   Check formula results and troubleshoot common errors.
-   Explain how Excel supports basic data analysis.

## 📁 Project Structure

``` text
Student-Marks-Analysis/
│
├── README.md
└── PR3.xlsx
```

## 🚀 How to Use

1.  Open `PR3.xlsx` in Microsoft Excel.
2.  Review the 20 student records in `Sheet1`.
3.  Enter the VLOOKUP formula in the VLOOKUP practice column.
4.  Try the FILTER examples in an empty area or a separate worksheet.
5.  Complete the suggested analysis tasks.
6.  Save your completed workbook with a new filename.

## ⚠️ Notes

-   The student names and marks are sample practice data, not verified
    academic records.
-   The workbook currently includes a VLOOKUP formula example; the
    FILTER column is provided for practice.
-   Use `FALSE` in VLOOKUP when you need an exact Student ID match.
-   The examples use the data range `A2:G21`. Adjust the range if you
    add or remove records.
-   Do not place a FILTER formula where its returned results would
    overwrite existing data.

------------------------------------------------------------------------

::: {align="center"}
### ✨ LEARN • PRACTISE • FILTER • ANALYSE

**Created with Microsoft Excel by Abhishek Choudhary**
:::
