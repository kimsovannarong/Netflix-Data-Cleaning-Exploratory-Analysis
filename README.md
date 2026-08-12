# 🎬 Netflix Data Analysis & Exploratory Analysis

## 📌 Project Overview
This project is an end-to-end data manipulation and exploratory analysis performed on the **Netflix Titles Dataset** using **Microsoft Excel**. 

The main objective of this project was to apply practical Excel capabilities—including data cleaning, text formatting, dynamic formulas (`SUMIFS`, `UNIQUE`, `XLOOKUP` ), and multidimensional **Pivot Tables**—to answer 5 specific analytical queries regarding Netflix's catalog.

---

## 🛠️ Excel Skills & Functions Demonstrated

* **Data Cleaning & Manipulation:** Formatted text strings, extracted numeric duration values from text, and parsed date fields.
* **Conditional Aggregation:** Used **`SUMIFS`** to calculate total movie durations within specified release year ranges.
* **Array Formulas:** Applied **`UNIQUE`** and **`SORT`** formulas to extract distinct values dynamically.
* **Pivot Table Analytics:** Configured custom Pivot Tables with custom row/column layouts, field sorting (Largest to Smallest), value aggregations (`SUM` vs `COUNT`), and date grouping (Months).

---

## ❓ Analytical Questions & Solutions

### 1. Count of Movies by Film Duration
* **Method:** Built a Pivot Table filtered by `type = Movie`.
* **Configuration:** Duration placed in **Rows**, Count of Titles in **Values**, sorted in descending order to identify the most common film lengths.

### 2. Total Duration of Movies Released (1990 – 2000)
* **Method:** Calculated using conditional multi-criteria aggregation (`SUMIFS`).
* **Formula:** `=SUMIFS(Duration_Col, Type_Col, "Movie", ReleaseYear_Col, ">=1990", ReleaseYear_Col, "<=2000")`

### 3. Monthly TV Show Additions (Regardless of Year)
* **Method:** Analyzed seasonality of TV show releases across months.
* **Configuration:** Pivot Table filtered for `TV Show`, with date fields grouped specifically by **Months** across the row layout.

### 4. Unique List of Release Years
* **Method:** Extracted distinct release years available in the catalog.
* **Formula / Feature:** Applied `=SORT(UNIQUE(Release_Year_Range))` for dynamic extraction.

### 5. Cross-Tabulation: Release Year vs. Rating vs. Duration
* **Method:** Multi-dimensional matrix constructed using Pivot Tables.
* **Configuration:** `Release Year` placed on **Rows**, `Rating` on **Columns**, and `Duration` set to **Sum** in the values area.

---
