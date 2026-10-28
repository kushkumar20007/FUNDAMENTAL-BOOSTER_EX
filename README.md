# 📊 PR. 1 — Fundamental Booster | Excel Formula Mastery


## 📌 Project Overview

**Fundamental Booster** is a hands-on Excel practice project designed to strengthen core spreadsheet skills through realistic, formula-based tasks. It combines student scores, employee information, product prices, and sales data to demonstrate how Excel can automate calculations, retrieve records, classify results, and analyze information.

The workbook is organized around practical exercises, with an emphasis on writing formulas, understanding cell references, and interpreting results.

## 🎯 Objectives

- Apply relative and absolute cell references correctly.
- Format data using appropriate number, date, currency, and text formats.
- Classify student grades with `IF` and nested `IF` formulas.
- Evaluate multiple conditions using `AND` and `OR`.
- Summarize data using `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS`.
- Retrieve records with `VLOOKUP`, `XLOOKUP`, and `INDEX`/`MATCH`.
- Find positions with `XMATCH`.
- Clean and transform text using `LEFT`, `FIND`, `UPPER`, and `LOWER`.
- Create dynamic references with `INDIRECT` and `OFFSET`.
- Work with dates, time, and mathematical rounding functions.
- Return filtered lists with `FILTER`.

## 🧰 Skills & Functions Covered

| Category | Functions / Concepts | Practice Goal |
|---|---|---|
| Cell references | Relative (`A1`), absolute (`$A$1`), mixed references | Build formulas that copy correctly |
| Formatting & data input | Bold, italic, currency, dates, custom number formats | Present data clearly |
| Logical functions | `IF`, nested `IF`, `AND`, `OR` | Assign grades and evaluate eligibility |
| Conditional analysis | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` | Count, total, and average matching records |
| Lookup functions | `VLOOKUP`, `XLOOKUP` | Retrieve names, product prices, and salaries |
| Dynamic lookup | `INDEX`, `MATCH`, `XMATCH` | Find records and positions flexibly |
| Text functions | `LEFT`, `FIND`, `UPPER`, `LOWER` | Extract names and standardize text |
| Dynamic references | `INDIRECT`, `OFFSET` | Build flexible cell and range references |
| Date & time | `TODAY`, `DATEDIF`, date subtraction | Calculate age and date differences |
| Math functions | `ROUND`, `CEILING`, `FLOOR` | Apply rounding rules |
| Dynamic arrays | `FILTER` | Extract students meeting score criteria |

## 🧪 Main Tasks

1. **Student grade classification** — use `IF` or nested `IF` formulas to assign grades based on scores.
2. **Count scores above 60** — use `COUNTIFS` to count students whose scores exceed 60.
3. **Find product prices** — use `VLOOKUP` to retrieve prices from product codes or sales data.
4. **Retrieve employee salaries** — use `XLOOKUP` to return salaries based on a selected employee or ID.
5. **Analyze joining dates** — format employee joining dates and calculate date-related values.
6. **Filter high-performing students** — use `FILTER` to return students with scores above 80.


## 🎥 Video Demonstration

[![Watch Video](https://img.shields.io/badge/🎥-Watch_Video-red?style=for-the-badge)((https://drive.google.com/file/d/1FvaRePRLKLKudx00dU6BR1QMPsuATuBr/view?usp=sharing))

## 🗂️ Practice Modules

### 1. Cell References, Formatting & Data Input
- Enter and format sample data.
- Compare relative and absolute references.
- Use `$A$1` to lock a reference when copying formulas.

### 2. IF, Nested IF, AND & OR
- Assign student grades from scores.
- Calculate sales discounts based on price thresholds.
- Identify students scoring above 80 in both Math and Science using `IF(AND(...))`.
- Check product discount eligibility using `IF(OR(...))`.

### 3. COUNTIFS, SUMIFS & AVERAGEIFS
- Count students scoring above 50 in Math.
- Sum sales for a chosen region and product.
- Average scores for students meeting a score condition.

### 4. VLOOKUP, INDEX/MATCH, XLOOKUP & XMATCH
- Find student names using student IDs.
- Retrieve product prices using product codes.
- Find a salesperson's sales for a selected month.
- Retrieve employee details without relying on a fixed return-column index.
- Find a product's position in a list.

### 5. Text Functions
- Extract a first name from a full name using `LEFT` and `FIND`.
- Convert text to uppercase or lowercase using `UPPER` and `LOWER`.

### 6. INDIRECT & OFFSET
- Build a reference dynamically from a text address.
- Create a dynamic range for sales analysis.

### 7. Date, Time & Math
- Calculate age from date of birth.
- Find the number of days between two dates.
- Apply `ROUND`, `CEILING`, and `FLOOR` in calculations.

### 8. FILTER
- Return a dynamic list of students scoring above 80.
- Practice criteria-based extraction without manually filtering the source table.

## 💡 Example Formulas

> Replace the sample cell references with the ranges used in your workbook.

| Purpose | Example formula |
|---|---|
| Assign a grade | `=IF(B2>=80,"A",IF(B2>=60,"B",IF(B2>=40,"C","Fail")))` |
| Count scores above 60 | `=COUNTIFS(B2:B20,">60")` |
| Sum sales for a region and product | `=SUMIFS(D2:D100,B2:B100,"West",C2:C100,"Laptop")` |
| Average scores above 60 | `=AVERAGEIFS(B2:B20,B2:B20,">60")` |
| Look up a product price | `=VLOOKUP(G2,A2:D20,4,FALSE)` |
| Retrieve a salary | `=XLOOKUP(G2,A2:A20,C2:C20,"Not Found")` |
| Match a row and column | `=INDEX(B2:E10,MATCH(G2,A2:A10,0),MATCH(H1,B1:E1,0))` |
| Find a product's position | `=XMATCH(G2,A2:A20,0)` |
| Extract the first name | `=LEFT(A2,FIND(" ",A2&" ")-1)` |
| Convert to uppercase | `=UPPER(A2)` |
| Convert to lowercase | `=LOWER(A2)` |
| Build a dynamic reference | `=INDIRECT("B"&A2)` |
| Return a value two rows below | `=OFFSET(A2,2,0)` |
| Calculate age in completed years | `=DATEDIF(A2,TODAY(),"Y")` |
| Calculate days between dates | `=B2-A2` |
| Round to two decimals | `=ROUND(A2,2)` |
| Round up to a multiple of 10 | `=CEILING(A2,10)` |
| Round down to a multiple of 10 | `=FLOOR(A2,10)` |
| Filter scores above 80 | `=FILTER(A2:C20,B2:B20>80,"No students found")` |

## ▶️ How to Use

1. Open the Excel workbook in Microsoft Excel.
2. Review the sample tables and identify the input and output cells.
3. Enter each formula in the appropriate result cell.
4. Check formula ranges and criteria against your actual data.
5. Test formulas with different IDs, scores, products, and dates.
6. Save the completed workbook as `.xlsx`.

## ⚠️ Compatibility Notes

- `XLOOKUP`, `XMATCH`, and `FILTER` require a compatible modern Excel version, such as Microsoft 365 or Excel 2021/2024.
- `DATEDIF` with `"Y"` returns completed years. Make sure the date-of-birth cell contains a valid Excel date.
- `VLOOKUP` with `FALSE` requests an exact match.
- Formula separators can differ by regional settings. If Excel rejects commas, your installation may require semicolons.
- Check all cell references before submitting; the formulas above are examples, not fixed references for every workbook.

## 📦 Deliverable

A completed Excel workbook containing the required formulas, formatted data, and outputs for the practice modules.
## 👨‍💻 Author

**Kush Kumar**
## 📚 Learning Outcome

By completing this project, learners practice using Excel for **data entry, logical decision-making, data lookup, date calculations, and formula-driven analysis**.

---

<p align="center">
  <b>Built for learning, practice, and continuous improvement.</b>
</p>
