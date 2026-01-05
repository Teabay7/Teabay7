## Hi there 👋

# 📊 Automated WAEC Result Entry and Performance Analysis System (Excel)

---

## 🧾 Project Overview
This project is an Excel-based system designed to **automate the entry, computation, and analysis of WAEC student results**.  
It classifies student performance based on **credit attainment**, **compulsory subject combinations (Mathematics & English)**, **gender distribution**, and **subject-level outcomes**.

The workbook consists of **5 sheets**:

1. **Entry Sheet** 🖊️  [Data entry sheet](https://github.com/Teabay7/Automated-WAEC-Result-Data-Entry-Excel-Sheet/blob/main/Data%20entry%20sheet.jpg)
   - Input students’ **Exam Number**, **Gender**, and **subject grades** via dropdowns.  
   - Includes **School Name** and **Centre Number**.  

2. **Database Sheet** 💾  
   - Master table of all student results used by formulas.

3. **Result Analysis Sheet** 📊  [Analysis sheet](https://github.com/Teabay7/Automated-WAEC-Result-Data-Entry-Excel-Sheet/blob/main/Analysis%20sheet.jpg)
   - Shows calculated columns, aggregations and conditional formatting.  

4. **Result Summary Sheet** 🎯  
   - Calculates **5 & 4 credits**, including **Maths & English**, **English only**, and **Maths only**.  
   - Aggregates results **by gender**. 

5. **Subject Analysis Sheet** 📈  [Subject Analysis](https://github.com/Teabay7/Automated-WAEC-Result-Data-Entry-Excel-Sheet/blob/main/Subject%20Analysis.jpg)
   - Shows **grades breakdown per subject**, with **Pass/Fail counts**, **Registered**, **Present**, and **Absent**.  


## ❗ Problem Statement
Manual analysis of WAEC results is:
- ⏱️ Time-consuming  
- ❌ Prone to calculation errors  
- 📉 Difficult to aggregate across gender and compulsory subjects  
Making schools struggle to generate accurate summaries for decision-making.
---

## 💡 Solution
An **automated Excel workbook** that uses:
- Structured data entry forms
- Validation rules
- Formula-driven logic  
to compute student results and generate **instant performance analysis**.


## 🛠️ Tools & Techniques Used
- 📘 Microsoft Excel  
- 🔽 Data Validation (Dropdown Lists)  
- 🔢 Logical Functions (`IF`, `IFS`)  
- 🔍 Lookup Functions (`INDEX AND MATCH` / `XLOOKUP`/'LET')  
- ⚙️ Automated Calculations & Aggregation  
---

## ⚙️ Key Formulas & Logic Used

## 🔍 Lookup & Data Retrieval

### Auto-display Database Values
    =IF(Result_Database!A2="","",Result_Database!A2)

Ensures blank cells remain blank and prevents unnecessary zeros or errors from appearing.

### Dynamic Record Retrieval
    =LET(
    result, XLOOKUP($B6, Result_Database!$A$2:$A$500, Result_Database!$B$2:$B$500),
    IFNA(result, ""))
    
  -Fetches student records using Exam Number.
  
  -**LET** improves readability and performance, while IFNA suppresses lookup errors.

**`INDEX` + `MATCH`** 
Dynamically retrieves student grades row-by-row

    =LET(result_row, MATCH($B6, Result_Database!$A$2:$A$500, 0),
    col_offset, COLUMN() - COLUMN($D$1) + 1,
    value, INDEX(Result_Database!$C$2:$AL$500, result_row, col_offset),
    IFNA(value, ""))

## 🧮 Aggregation & Counting
### Count Subject Grades Per Student
    =LET(
    grade_to_count, Table5[[#Headers],[A1]],
    grade_count, COUNTIF(Table5[@[English Language]:[Electrical/ Electronics]], grade_to_count),
    IF(grade_count=0, "", grade_count))
    
-Counts how many times a specific grade (e.g., A1) appears across a student’s subjects, returning blank if none exist.

## ✅ Conditional Classification
### Determine Overall Pass Status
    =IF(
    AND(
     ISNA(MATCH([@[English Language]], {"D7","E8","F9","ABS","WITHHELD",""}, 0)),
     ISNA(MATCH([@Mathematics], {"D7","E8","F9","ABS","WITHHELD",""}, 0)),
     [@[NO OF CREDITS]] > 4), "PASS", "")
   
-Classifies a student as PASS only if:
- English and Mathematics are passed
- Total credits exceed four
- This ensures compliance with WAEC requirements.

## ➖ Calculations & Error Handling
    =IFERROR(A7-D7,"")
-Performs calculations while preventing Excel errors from displaying in reports.

## 📊 Attendance & Subject Analysis
### Count Students Present Per Subject
    =SUMPRODUCT(
    --(INDEX(Analysis!$D$6:$AL$500,,MATCH($B2,Analysis!$D$5:$AL$5,0))<>""),
    --(INDEX(Analysis!$D$6:$AL$500,,MATCH($B2,Analysis!$D$5:$AL$5,0))<>"ABS"))
-Counts students present for a selected subject by excluding blanks and absentees.
-The formula dynamically identifies the subject column, enabling scalable subject analysis.

## ✨ Key Features
- 🧩 Controlled data entry using dropdown lists  
- 🧮 Automatic computation of student credits  
- 🏫 Classification of results into WAEC performance categories:
  - ✅ 5 Credits including Mathematics & English  
  - ➕ 5 Credits with Mathematics or English only  
  - 📚 4 Credits including Mathematics & English  
  - ⚠️ Below 4 Credits  
- 🚻 Gender-based performance analysis  
- 🧍 Attendance tracking (Registered, Present, Absent)  
- 📑 Subject-level analysis showing:
  - Grade distribution  
  - Pass and fail counts  
  - Registration and attendance  

---

## 📈 Outcome & Insights
- ✔ Improved accuracy in result computation  
- ✔ Reduced manual workload  
- ✔ Faster generation of school-level performance reports  
- ✔ Clear insights for academic planning and evaluation  

---

## 🚀 Possible Improvements for later upgrade
- 📊 Interactive visual dashboards   
- 🏫 Multi-school comparative performance analysis  

---



