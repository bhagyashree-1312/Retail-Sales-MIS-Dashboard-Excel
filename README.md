# Retail-Sales-MIS-Dashboard-Excel

# Retail Sales MIS Dashboard (Excel)

An end-to-end **Management Information System (MIS) dashboard** built entirely in Excel, covering Q1 (Jan–Mar) retail transactions across three branches. The workbook demonstrates a full BI pipeline — raw monthly data → Power Query consolidation → calculated fields → PivotTables → interactive dashboard — without any external tools.  
---

## 📌 Overview

| | |
|---|---|
| **Dataset** | 1,000 retail transactions (Jan–Mar) across 3 branches, 3 cities |
| **Source** | Three monthly worksheets (`January`, `February`, `March`) |
| **Consolidation** | Power Query `Table.Combine` append into `Sales_Master` |
| **Analysis layer** | 12 PivotTables on a shared PivotCache |
| **Presentation layer** | 8 KPI cards, 8 charts, 5 cross-filtering slicers |
| **Narrative layer** | Auto-derived Business Insights summary |

---

## 🗂️ Workbook Structure

```
Dashboard          → Executive view: KPI cards, charts, slicers
Business Insights  → Narrative takeaways (8 key findings)
Pivot              → 12 PivotTables (the analytical engine)
Sales_Master       → Consolidated 1,000-row fact table (24 columns)
January            → Raw source table (352 rows)
February           → Raw source table (303 rows)
March              → Raw source table (345 rows)
```

---

## ⚙️ Power Query (Get & Transform) Layer

Four queries, written in Power Query M, consolidate the three monthly tables into one master fact table.

```m
section Section1;

shared tbl_January = let
    Source = Excel.CurrentWorkbook(){[Name="tbl_January"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{
        {"Invoice ID", type text}, {"Branch", type text}, {"City", type text},
        {"Customer type", type text}, {"Gender", type text}, {"Product line", type text},
        {"Unit price", type number}, {"Quantity", Int64.Type}, {"Tax 5%", type number},
        {"Total", type number}, {"Date", type datetime}, {"Time", type number},
        {"Payment", type text}, {"cogs", type number}, {"gross income", type number},
        {"Rating", type number}})
in
    #"Changed Type";

// tbl_February / tbl_March follow the same pattern

shared Sales_Master = let
    Source = Table.Combine({tbl_January, tbl_February, tbl_March})
in
    Source;
```

**What this does:**
- Each monthly query pulls its Excel Table as a `Source`, then enforces strict data types (`Table.TransformColumnTypes`) — this is what makes downstream date/number functions reliable instead of silently failing on text-formatted values.
- `Sales_Master` uses **`Table.Combine`** (an append/union query) to stack all three months into a single 1,000-row table — the standard MIS pattern for consolidating monthly tabs without copy-paste.
- Loaded as a workbook Connection + Table, refreshable via `Data → Refresh All` whenever a new month's data is pasted into its sheet.

---

## 🧮 Calculated Columns (Sales_Master)

Seven columns are engineered on top of the raw fields using native Excel Table structured-reference formulas (auto-propagate to every new row):

| Column | Formula | Purpose |
|---|---|---|
| `gross margin percentage` | `=([@[gross income]]/[@Total])*100` | Per-transaction margin % |
| `year` | `=YEAR([@Date])` | Time-intelligence key |
| `Month Number` | `=MONTH([@Date])` | Chronological sort key (vs. alphabetical month names) |
| `Month Name` | `=TEXT([@Date],"mmmm")` | Human-readable label for pivots/slicers |
| `Day Name` | `=TEXT([@Date],"dddd")` | Weekday-pattern analysis |
| `Hour` | `=HOUR([@Time])` | Hourly footfall/sales pattern |
| `Week Type` | `=IF(WEEKDAY([@Date],2)<=5,"Weekday","Weekend")` | Weekday vs. weekend segmentation |
| `Average Order Value` | `=[@Total]` | Base measure — averaged inside PivotTables to derive true AOV |

---

## 📊 PivotTables (12, one shared PivotCache)

All 12 pivots run off a single cache sourced from `Sales_Master` (1,000 records), which keeps the workbook light and ensures every report refreshes in sync.

| Pivot | Rows | Value | Feeds |
|---|---|---|---|
| KPI Values | — | Sum Total, Sum Gross Income, Count Invoices, Sum Quantity, Avg Rating | KPI cards |
| Gender-wise Sales | Gender | Sum of Total | — |
| Monthly Sales Trend | Month Name | Sum of Total | Line chart |
| Best Selling Product Line | Product line | Sum of Total | KPI card |
| Sales by City | City | Sum of Total | Bar chart |
| Sales by Product Line | Product line | Sum of Total | Bar chart |
| Sales by Branch | Branch | Sum of Total | Pie chart |
| Customer Type | Customer type | Sum of Total | Bar chart |
| Payment Method | Payment | Sum of Total | Doughnut chart |
| Top Performing Branch | Branch | Sum of Total | KPI card |
| Hourly Sales | Hour | Sum of Total | Line chart |
| Average Rating by Product Line | Product line | Average of Rating | Bar chart |

---

## 📈 Charts (8, native Excel charts)

| Chart | Type | Bound Pivot |
|---|---|---|
| Monthly | Line | Monthly Sales Trend |
| Sales by Product Line | Bar | Sales by Product Line |
| Sales by Branch | Pie | Sales by Branch |
| Sales by City | Bar | Sales by City |
| Payment Method Analysis | Doughnut | Payment Method |
| Customer Type Analysis | Bar | Customer Type |
| Hourly Sales Analysis | Line | Hourly Sales |
| Average Rating by Product Line | Bar | Average Rating by Product Line |

---

## 🎯 KPI Cards (Dashboard sheet)

Implemented as **cell-linked textboxes** (`textlink="Pivot!A2"` etc.) rather than hardcoded values — each card renders live from the `KPI Values` pivot and updates automatically on refresh/filter.

| KPI | Value |
|---|---|
| Total Sales | ₹322,966.75 |
| Total Profit (Gross Income) | ₹15,379.37 |
| Total Orders | 1,000 |
| Total Quantity Sold | 5,510 |
| Average Rating | 7.0 |
| Average Transaction Value | ₹322.97 |
| Top Performing Branch | Branch C — ₹110,568.71 |
| Best Selling Product Line | Food & Beverages — ₹56,144.84 |

---

## 🎛️ Slicers (5, cross-filtering)

`Month` · `City` · `Payment` · `Product line` · `Customer type`

Each slicer has its own `slicerCache` wired to every relevant PivotTable, so one click filters all 12 pivots, all 8 charts, and all 8 KPI cards in sync — the core of the "interactive dashboard" experience.

---

## 💡 Business Insights (auto-generated narrative)

1. January generated the highest sales.
2. Branch C recorded the highest revenue.
3. Food & Beverages was the top-selling product line.
4. Cash was the most preferred payment method.
5. Members contributed more revenue than Normal customers.
6. Food & Beverages achieved the highest average customer rating.
7. Peak sales occurred around 7 PM.
8. Naypyitaw generated the highest city-wise sales.

---

## 🛠️ Skills Demonstrated

- Power Query (M language): source connections, type transformation, append/`Table.Combine` consolidation
- Excel Tables & structured references (dynamic calculated columns)
- Date/time engineering (`YEAR`, `MONTH`, `TEXT`, `HOUR`, `WEEKDAY`)
- PivotTables & shared PivotCache architecture
- Native Excel charting (line, bar, pie, doughnut)
- Slicers & cross-filter dashboard interactivity
- Cell-linked textboxes for live KPI cards
- MIS/executive dashboard design and data storytelling

---

## 🚀 How to Use

1. Open the workbook in Excel (2016+ recommended for slicer/chart support).
2. Use the slicers on the Dashboard sheet to filter by Month, City, Payment method, Product line, or Customer type.
3. To refresh with new data: paste new rows into the relevant monthly sheet's table, then `Data → Refresh All`.
4. Review `Business Insights` for the narrative summary of current filtered results.

---

## 📁 File

`Retail_Sales_MIS_Dashboard.xlsx`
