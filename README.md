# 🏥 NHS Patient Waiting Time Analysis

## 📊 Project Overview
An end-to-end data analytics project analysing patient waiting times
across 8 NHS Trusts in England. Built entirely in Microsoft Excel
using professional analyst tools and techniques.

**Scenario:** As a junior data analyst at NHS England, I was asked
to analyse patient waiting time data and produce an executive
dashboard for the Head of Operations.

---

## 🛠️ Tools & Techniques Used

| Tool | Purpose |
|------|---------|
| Power Query | Data cleaning and transformation |
| Excel Formulas | VLOOKUP, SUMIF, COUNTIF, AVERAGEIF, INDEX MATCH |
| Pivot Tables | Multi-dimensional data summarisation with Slicers |
| VBA Macros | Automated report formatting and breach highlighting |
| PowerPivot & DAX | Advanced KPI measures and data modelling |
| Dashboard | Executive-level visual reporting |

---

## 📁 Project Structure

| Sheet | Contents |
|-------|---------|
| Raw Data | 79 NHS patient records across 8 Trusts |
| PQ Output | Power Query cleaned data with Waiting Category |
| Analysis | 8 KPIs, regional averages, VLOOKUP lookups |
| Pivot | 3 Pivot Tables with Region Slicer |
| Dashboard | Professional NHS dashboard with KPI cards and charts |
| Trust Lookup | Reference table for VLOOKUP |
| DAX Pivot | PowerPivot measures — Breach Rate %, Avg Wait, Emergency Count |

---

## 📈 Key Findings

- **79 patients** analysed across 8 NHS Trusts in England
- **Average waiting time: 105 days** across all departments
- **39.2% breach rate** — over a third of patients waiting beyond the 18-week NHS target
- **Manchester University NHS** had the highest breach rate at **57.1%**
- **King's College Hospital London** performed best at **16.7%** breach rate
- **Cardiology** had the longest average wait at **125 days**
- **8 Emergency** patients | **23 Urgent** | **48 Routine**

---

## 🔄 VBA Macro Features

- **NHS_Format_Report** — Auto-highlights breached patients (>126 days) in red
and fast-tracked patients in green. Processes all rows with one button click.
- **Practice_Highlight_Priority** — Colour codes Priority column by
Emergency (red), Urgent (orange), Routine (green) with patient counts.

---

## 📊 DAX Measures (PowerPivot)

```dax
Breach Rate % := DIVIDE(
    CALCULATE(COUNTROWS('Table1'), 'Table1'[Waiting Days] > 126),
    COUNTROWS('Table1'))

Emergency Count := CALCULATE(
    COUNTROWS('Table1'),
    'Table1'[Priority] = "Emergency")
```

---

## 🎯 Skills Demonstrated

- Data cleaning and transformation with Power Query
- Complex Excel formula writing (VLOOKUP, SUMIF, COUNTIF, AVERAGEIF)
- Pivot Table analysis with Slicers
- VBA macro development for automation
- DAX formula writing in PowerPivot
- Professional dashboard design
- NHS domain knowledge (18-week waiting time targets)

---

## 👤 Author

**Ehtisham Nazir**
Aspiring Junior Data Analyst | UK
[GitHub](https://github.com/EhtishamNazir)

---

## 📌 Note
This project uses realistic synthetic NHS data generated for
portfolio purposes. No real patient data has been used.
