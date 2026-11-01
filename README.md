::: {align="center"}
# 📊 SALES DATA ANALYSIS --- EXCEL PROJECT

### Turning raw sales records into useful business insights

**A beginner-friendly Microsoft Excel practice project**

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Project](https://img.shields.io/badge/Type-Data%20Analysis-2563EB?style=for-the-badge)
![Level](https://img.shields.io/badge/Level-Beginner-16A34A?style=for-the-badge)
:::

------------------------------------------------------------------------

## 👨‍💻 Author

**Abhishek Choudhary**

## 📌 Project Overview

This project uses a sample sales dataset to practise data handling and
analysis in Microsoft Excel. It contains sales transaction records with
product, region, salesperson, amount and date information. The workbook
also includes a small employee joining-date section and formula practice
examples.

The main purpose is to learn how Excel formulas can help organise
information, apply business rules, retrieve records and perform basic
analysis.

## 🎯 Project Objectives

-   Understand and organise sales data in a structured worksheet.
-   Calculate discount values based on sales amount.
-   Identify records that meet a defined eligibility condition.
-   Practise extracting day, month and year from dates.
-   Round values using different rounding functions.
-   Retrieve information using lookup and reference functions.
-   Filter records based on one or more conditions.
-   Build confidence in using Excel for entry-level data analysis.

## ❓ Problem Statements

The workbook is designed around the following practical questions:

1.  How can sales transactions be organised by product, region and
    salesperson?
2.  How can a discount value be calculated based on the transaction
    amount?
3.  How can transactions be marked as eligible or not eligible using a
    condition?
4.  How can the day number be extracted from an employee's joining date?
5.  How can a date be constructed from separate year, month and day
    values?
6.  How can amounts be rounded to a required number of digits or to a
    specified multiple?
7.  How can a matching sales record or value be retrieved from a data
    range?
8.  How can the required rows be extracted using conditions such as
    region, product or amount?

## 🗂️ Dataset Description

The Excel workbook (`Sales.PR.xlsx`) contains the following sections:

### 1. Sales Transactions

  Column                    Description
  ------------------------- -----------------------------------
  Sales ID                  Unique transaction identifier
  Product                   Product sold
  Region                    Sales region
  Salesperson               Person responsible for the sale
  Amount                    Transaction amount
  Date                      Transaction date
  Discount                  Calculated discount value
  Eligible / Not Eligible   Result of the eligibility formula

**Dataset size:** 20 sales transaction records.

### 2. Employee Joining-Date Practice

  Column          Description
  --------------- -------------------------------------
  Employee ID     Employee identifier
  Employee Name   Employee name
  Joining Date    Employee joining date
  Day             Day extracted from the joining date
  New Date        Practice column for creating a date

**Dataset size:** 10 employee records.

### 3. Lookup and Formula Practice

The workbook also contains examples for practising functions such as
`VLOOKUP`, `MATCH`, `INDEX`, `XMATCH`, `ROUND` and `CEILING`.

## 🧮 Excel Functions Practised

  --------------------------------------------------------------------------------------------
  Function                Purpose                 Example
  ----------------------- ----------------------- --------------------------------------------
  `IF`                    Apply a condition       `=IF(E2>20000,"High","Low")`

  `OR`                    Check whether at least  `=OR(E2>6000,E2>7000)`
                          one condition is true   

  `DATE`                  Create a date from      `=DATE(2026,11,1)`
                          year, month and day     

  `DAY`                   Extract the day from a  `=DAY(Q2)`
                          date                    

  `ROUND`                 Round to a specified    `=ROUND(1256.78,2)`
                          number of digits        

  `CEILING`               Round up to a specified `=CEILING(1256,100)`
                          multiple                

  `FILTER`                Return rows that meet a `=FILTER(A2:H21,C2:C21="South","No data")`
                          condition               

  `VLOOKUP`               Search vertically for a `=VLOOKUP(A2,A2:E21,5,FALSE)`
                          value                   

  `MATCH`                 Find the position of a  `=MATCH("Person 5",D2:D21,0)`
                          value                   

  `INDEX`                 Return a value at a     `=INDEX(E2:E21,3)`
                          given position          

  `XMATCH`                Find the relative       `=XMATCH("Mouse",B2:B21)`
                          position of a value     
  --------------------------------------------------------------------------------------------

> **Note:** `FILTER` and some of the examples above are included as
> learning exercises in this README; check your Excel version and
> worksheet ranges before using them. `FILTER` and `XMATCH` require a
> supported modern version of Excel, such as Microsoft 365 or Excel 2021
> and later.

## 🔍 Analysis Workflow

1.  **Inspect the data** --- review column headings, records and data
    types.
2.  **Apply formulas** --- calculate discount and eligibility results.
3.  **Work with dates** --- use `DATE` and `DAY` to create and extract
    date components.
4.  **Round values** --- practise `ROUND` and `CEILING` for numeric
    calculations.
5.  **Retrieve records** --- use lookup and reference functions to find
    specific values.
6.  **Filter records** --- extract rows that match selected business
    conditions.
7.  **Validate results** --- check formula outputs against the source
    records.

## 💡 Example Business Questions to Explore

-   Which products appear in the sales records?
-   Which regions have recorded transactions?
-   Which transactions have an amount above ₹20,000?
-   Which sales records meet the eligibility condition?
-   What is the total sales amount for a selected region and product?
-   Which employee joined on a particular day of the month?
-   What are the results when transaction amounts are rounded to the
    nearest 10, 100 or 1,000?

## 📁 Project Files

``` text
Sales-Data-Analysis/
│
├── README.md
└── Sales.PR.xlsx
```

## 🛠️ Tools Used

-   **Microsoft Excel** --- data organisation, formulas and analysis.
-   **Excel worksheet functions** --- conditional, date, rounding,
    lookup and filtering operations.
-   **Markdown** --- project documentation.

## 🚀 How to Use This Project

1.  Download or open `Sales.PR.xlsx` in Microsoft Excel.
2.  Review the sales transaction and employee sections.
3.  Check the formulas already entered in the workbook.
4.  Practise the listed functions using the supplied data.
5.  Try the business questions and compare your results with the source
    records.
6.  Save your completed workbook with a new filename so the original
    data remains unchanged.

## 📚 Learning Outcomes

After completing this project, a learner should be able to:

-   Work with structured sales and employee data.
-   Apply basic Excel formulas and logical conditions.
-   Extract and construct dates.
-   Round numeric values using different methods.
-   Search, retrieve and filter data.
-   Explain how spreadsheet functions support basic business analysis.

## ⚠️ Notes and Limitations

-   The workbook contains sample/practice data; it should not be treated
    as verified company sales information.
-   The employee section is a separate practice dataset, not a confirmed
    employee register.
-   Review formula conditions before using them for real business
    decisions. For example, the workbook's current eligibility formula
    uses `OR(E2>6000,E2>7000)`, where the second condition is already
    covered by the first. If the intended rule is simply an amount above
    ₹6,000, `=IF(E2>6000,"Eligible","Not Eligible")` is clearer.
-   The discount formula returns numeric values based on its current
    calculation; confirm the intended discount rate and amount before
    interpreting it as a currency discount.

------------------------------------------------------------------------

::: {align="center"}
### ✨ Learn • Practise • Analyse • Improve

**Created by Abhishek Choudhary**
:::
