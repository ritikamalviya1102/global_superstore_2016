# 🛒 Global Superstore Sales Analysis & Dashboard

A comprehensive e-commerce data analysis and interactive Excel dashboard built using the **Global Superstore** dataset. This project focuses on end-to-end data processing—from auditing data quality to generating descriptive statistics, applying conditional formatting, and building dynamic Pivot Table visual reports.

---

## 📌 Project Overview

As an e-commerce data analyst, the goal of this project was to analyze historical sales data across international markets, evaluate performance metrics, and create an executive-ready dashboard to drive data-informed business decisions.

### Key Objectives:
1. **Data Cleaning & Quality Assurance:** Detect and resolve duplicate entries and handle missing values appropriately.
2. **Descriptive Statistics:** Calculate foundational statistical measures (Mean, Median, Mode) using native Excel formulas.
3. **Conditional Formatting:** Apply visual highlights to identify top sales performance.
4. **Data Visualization:** Build category-wise revenue charts and line charts to track multi-year revenue growth.
5. **Interactive Dashboard:** Construct a multi-faceted Pivot Table dashboard breaking down total revenue, category performance, annual trends, and regional/market sales.

---

## 📊 Key Findings & Metrics Summary

| Metric | Sales ($) | Profit ($) | Quantity Sold | Shipping Cost ($) |
| :--- | :--- | :--- | :--- | :--- |
| **Mean** | $246.49 | $28.61 | 3.48 | $26.48 |
| **Median** | $85.05 | $9.24 | 3.00 | $7.79 |
| **Mode** | $12.96 | $0.00 | 2.00 | $1.35 |

* **Total Revenue:** **$12,642,501.91** across 51,290 sales transactions.
* **Top Performing Category:** **Technology** ($4,744,557.50), followed by **Furniture** ($4,110,451.90) and **Office Supplies** ($3,787,492.51).
* **Top Market by Revenue:** **Asia Pacific** ($4,042,658.27), followed by **Europe** ($3,287,336.23) and **USCA** ($2,364,129.03).
* **Yearly Revenue Growth:**
  * **2012:** $2,259,450.90
  * **2013:** $2,677,438.69
  * **2014:** $3,405,746.45
  * **2015:** $4,299,865.87 *(Consistent year-over-year growth)*

---

## 🛠️ Excel Tools & Techniques Applied

* **Data Cleaning:** `Remove Duplicates`, `Go To Special (Blanks)`
* **Formulas:** `=AVERAGE()`, `=MEDIAN()`, `=MODE.SNGL()`, `=SUMIF()`
* **Formatting:** `Conditional Formatting (Top 10 Rules)`, Merge & Center, Header Styling
* **Data Summarization & Charts:** `PivotTables`, `PivotCharts (2-D Column, Line, Bar Charts)`, `Date Grouping (Years)`, `Slicers`

---

## 📂 Repository Contents

* `global_superstore_2016.xlsx` — Source Excel file containing the raw dataset, summary calculations, and interactive dashboard sheets.
* `README.md` — Project documentation and insights summary.

---

## 💻 How to View the Project

1. Download or clone this repository:
   ```bash
   git clone [https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git](https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git)
