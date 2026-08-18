# Retail Sales Data Analysis — Excel Project

An end-to-end Excel project covering data cleaning, formulas, PivotTables, Power Query, and dashboard design, built on a ~10,000-row raw retail sales dataset (Bangladesh regions, 2023–2025).

## 📊 Project Overview

This project simulates a real-world messy dataset — inconsistent date formats, mismatched text casing, missing values, and duplicate records — and works through it using a structured, staged approach: from basic cleaning to an interactive dashboard.

**Dataset:** 10,000+ retail transaction records including Order ID, Order Date, Customer Name, Region, Category, Product, Quantity, Unit Price, Discount %, Payment Method, Sales Rep, and Order Status.

## 🎯 Objectives

- Clean and standardize raw, inconsistent transactional data
- Build calculated fields and lookups to enrich the dataset
- Summarize and analyze sales performance across regions, categories, and time
- Solve real data-quality issues (ambiguous date formats) using the right tool for the job
- Present findings in a single interactive dashboard

## 🛠️ Skills & Tools Demonstrated

| Area | Techniques Used |
|---|---|
| Data Cleaning | `TRIM`, `PROPER`, Find & Replace, Remove Duplicates, Conditional Formatting |
| Formulas | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, nested `IF`, `INDEX/MATCH`, `XLOOKUP`, `IFERROR` |
| Data Analysis | PivotTables, PivotCharts, Slicers, Date Grouping |
| Data Transformation | Power Query (custom columns, locale-aware date parsing, repeatable query steps) |
| Data Validation | Dropdown lists for controlled inputs |
| Dashboard Design | KPI cards, linked charts, slicer-driven filtering, layout & formatting |

## 🧹 Data Cleaning Process

The raw data intentionally includes:
- **Inconsistent casing** in customer names and payment methods (e.g. `bkash`, `BKASH`, `Bkash`)
- **Mixed date formats** (`2024-04-14`, `16/08/2024`, `11/10/2024`, `16-Aug-2024`) — including genuinely ambiguous day/month vs month/day values
- **Missing values** scattered across Region, Sales Rep, Customer Name, and Payment Method
- **~1% duplicate rows**

**Key finding:** Simple formulas (`DATEVALUE`) could not reliably resolve all date ambiguities — some rows failed silently or couldn't be parsed at all. This was solved using **Power Query** with a fallback parsing strategy (attempt UK-style day/month parsing, then fall back to US-style month/day), which resolved 100% of the ambiguous date entries — a good example of recognizing the limits of formulas and reaching for the right tool.

## 📐 Key Formulas Used

```excel
Revenue           = Quantity * UnitPrice * (1 - DiscountPct)
Order Tier        = IF(Revenue>=20000,"High Value",IF(Revenue>=5000,"Mid Value","Low Value"))
Category Manager  = XLOOKUP(Product, Lookup!Product, Lookup!Manager, "Not found")
Regional Revenue  = SUMIFS(Revenue, Region, "Dhaka")
```

## 📈 Dashboard Features

- KPI summary cards: Total Revenue, Total Orders, Average Order Value, Delivered Orders
- Revenue by Region × Category (column chart)
- Revenue by Sales Rep (bar chart)
- Monthly revenue trend (line chart)
- Interactive Slicers for Category and Order Status filtering

## 📁 Repository Structure

```
├── data/
│   └── retail_sales_raw.csv          # Original raw dataset
├── Retail_Sales_Analysis.xlsx        # Full workbook (cleaning, formulas, pivots, dashboard)
├── screenshots/
│   └── dashboard.png                 # Dashboard preview
└── README.md
```

## 🚀 How to Use

1. Download `Retail_Sales_Analysis.xlsx`
2. Open the `RawData` sheet to see the original, unprocessed data
3. Review `CleanData` for the Power Query–cleaned output
4. Explore `Pivot_1`–`Pivot_3` for the underlying analysis
5. View `Dashboard` for the consolidated summary

## 🔑 Key Takeaways

- Real-world data is rarely clean — this project deliberately practices identifying *and* resolving common data quality issues rather than assuming clean input
- Formulas and PivotTables are powerful but have real limits (e.g. ambiguous text parsing); Power Query exists to solve exactly these limits
- A dashboard is only as good as the pipeline feeding it — cleaning and structuring the data properly made every downstream step (formulas, pivots, charts) far more reliable

## 👤 Author

**Rahat**

