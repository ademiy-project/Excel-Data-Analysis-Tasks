# Excel: Data Analysis Tasks

A set of practical Excel tasks for a **Data Analyst** role: combining data sources, building **pivot tables**, calculating **KPIs**, preparing **management-ready reports** and **charts**, structuring raw data, and designing **A/B test groups**.

## Part 1: Sales by Territory (`Тестирование_Аналитик.xlsx`)

### Task
1. Combine the data from two sheets (`Лист1` and `Лист2`)
   - 1.1 Build a pivot table by **weeks** and **territories**
   - Find the **top-3 territories by share of total turnover** and the **top-3 by turnover per warehouse** for the last week
   - 1.2 Build a chart of weekly dynamics of turnover and turnover per warehouse across all territories

### Data
- `Лист1`: date, territory, turnover (units and ₽), number of orders, clients
- `Лист2`: date, territory, number of warehouses, orders, clients
- Period: 28.04.2020 – 01.06.2020, 18 territories

### Solution
- Combined the sheets with **`SUMIFS`** on two keys (date + territory)
- Calculated **turnover per warehouse** (`turnover / warehouses`)
- Added the week number with **`ISOWEEKNUM`**
- Built **pivot tables** by week × territory with turnover, turnover per warehouse and **% share**
- Built a **combo chart**: weekly turnover (columns, mln ₽) + turnover per warehouse (line)

### Results: Top-3 Territories, Week 22 (25.05–31.05.2020)
| Place | By share of total turnover | Share | By turnover per warehouse | Turnover per warehouse, ₽ |
|---|---|---|---|---|
| 1 | Санкт-Петербург Север | 27.7% | Москва Восток | 400,664 |
| 2 | Санкт-Петербург Юг | 20.8% | Москва Запад | 381,298 |
| 3 | Москва Запад | 15.0% | Санкт-Петербург Север | 336,584 |

\* Week 22 was taken as the last full week. Week 23 contains only one day (01.06).

**Weekly dynamics:** turnover grew from ≈837 mln ₽ (week 18) to ≈1,056 mln ₽ (week 22). Turnover per warehouse stayed stable at ≈237–254 thousand ₽.

## Part 2: Assignment (`Assignment.xlsx`)

### Sheet "Diagrams": Data Visualization
- Calculated total monthly sales with **`SUMPRODUCT`** (quantity × price)
- Built charts for management:
  - column chart: product sales in units (Jan–Apr 2021)
  - line chart: revenue in tenge (Jan–Apr 2021)
  - doughnut chart: sales structure by product type (base product 52.4%, additional products 34.5%, service 13.1%)

### Sheet "Structuring": Data Structuring & Segmentation
- **Split unstructured data** (`msisdn|City|market_id|...`) into a proper table
- Added `market_desc` (B2C / B2B / B2O / VIP) with a lookup by `market_id`
- **Traffic segmentation** (upload + download) in 5 GB steps up to 50+ GB, counted with **`COUNTIF`**
- **Pivot table** of subscribers by city, expandable to client type, with **% of all subscribers**
- **Interactive slicer** filter by modem
- **Conditional formatting** and a chart for the segments

| City | Subscribers | % of all |
|---|---|---|
| Caracas | 202 | 67.3% |
| Maracay | 98 | 32.7% |
| **Total** | **300** | **100%** |

The largest traffic segment is **50+ GB** (79 subscribers).

### Sheet "AB test": A/B Test Group Design
- Built **two balanced groups of 3,000 numbers** (test and control) using **random sampling** (`RAND`)
- Marked every number as `Test` / `Control` / `Out of test`
- Checked **group balance** with pivot tables by ARPU, smartphone share and region (stratification check)

| Group | Users | Avg ARPU | Smartphone share |
|---|---|---|---|
| Control | 3,000 | 1,831.6 | 71.3% |
| Test | 3,000 | 1,830.7 | 71.2% |
| Out of test | 15,644 | 1,828.0 | 69.9% |

The groups are balanced on ARPU, smartphone share and regional distribution (Beijing, Shanghai, Shenzhen, Guangzhou).

### Sheet "Hidden"
Found and unhidden the hidden sheet, as the task required.

## Part 3: Novartis Pharma Services AG Test Case
*Coming soon.*

## Topics Covered
**Excel Skills**
- Formulas: `SUMIFS`, `COUNTIF`, `SUMPRODUCT`, `ISOWEEKNUM`, `RAND`, lookups
- Pivot tables, calculated fields, % of total
- Slicers (interactive filters)
- Conditional formatting
- Text to columns (data structuring)
- Charts: column, line, doughnut, combo chart with secondary axis

**Analytics**
- Data merging from multiple sources
- KPI calculation: turnover, turnover per warehouse, share of total
- Territory ranking (top-N analysis)
- Weekly trend analysis
- Customer / subscriber segmentation
- A/B test design: random sampling, group balance check, ARPU
- Management reporting and data visualization

## Tools
Microsoft Excel

## Files
- `Тестирование_Аналитик.xlsx`: Part 1 (sales by territory)
- `Assignment.xlsx`: Part 2 (Diagrams, Structuring, AB test)
- `Excel test case Novartis Pharma Services AG.xlsx`: Part 3 *(coming soon)*

## Keywords
`excel` `microsoft-excel` `pivot-tables` `sumifs` `countif` `sumproduct` `slicers` `conditional-formatting` `data-visualization` `charts` `dashboard` `kpi` `data-analysis` `data-cleaning` `data-structuring` `segmentation` `ab-testing` `random-sampling` `sales-analysis` `business-analytics` `management-reporting` `data-analyst` `portfolio-project`
