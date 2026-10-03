# Excel: Data Analysis Tasks

![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=googlesheets&logoColor=white)
 
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

## Part 3: Novartis Pharma Services AG Test Case (`Excel_test_case_Novartis_Pharma_Services_AG_.xlsx`)

### Task
A self-checking Excel test with 24 sequential questions on a telecom call database. Each question unlocks only after the previous one is answered correctly.

**Status: test passed ("Тест пройден. Поздравляем!")**

### Data
`База данных` sheet: 623 calls, 292 unique subscribers, 6 tariff plans
| Column | Description |
|---|---|
| `Абонент` | Subscriber number |
| `Тарифный план` | Tariff plan (1–6) |
| `Направление звонка` | Call direction: in-network, other operators, landline, international, other |
| `Начисления, тенге` | Charges, KZT |
| `Продолжительность разговора, секунды` | Call duration, seconds |

### Questions Covered
- Totals: charges, call duration (seconds / minutes), number of calls
- Breakdowns by call direction and tariff plan
- Averages: cost per minute, duration per call, charges and minutes per subscriber
- Shares (%): in-network duration, tariff plan 1 charges, tariff plan 2 subscribers
- Unique counts: subscribers in total, by direction, by tariff plan
- Lookups and ranking: tariff plan of a given subscriber, top subscriber by charges, 7th place by charges
- Finding call directions not used on tariff plan 6

### Key Results
| Metric | Result |
|---|---|
| Total charges | 3,899.58 KZT |
| Total duration | 7,883 sec (131.4 min) |
| Total calls | 623 |
| Unique subscribers | 292 |
| In-network share of duration | 91.55% |
| Average cost per minute | 29.68 KZT |
| Average call duration | 13 sec |
| Average charges per subscriber | 13.35 KZT |
| Subscriber with the highest charge | 77010000109 |
| Call directions not used on tariff plan 6 | Landline, Other |

### Solution
Formulas: `SUM`, `SUMIF`, `SUMIFS`, `COUNTIFS`, `SUMPRODUCT` (unique counts), `ROUND`, `VLOOKUP`, `XLOOKUP`, `INDEX`, and dynamic arrays: `UNIQUE`, `FILTER`, `SORTBY`.

**Note:** the answer sheet is protected. The answer to question 20 (2705 seconds) was entered as text, so the checker marks it wrong. The correct formula-based answer is in cell **D27**, below the table.
## Keywords
`excel` `microsoft-excel` `pivot-tables` `sumifs` `countif` `sumproduct` `slicers` `conditional-formatting` `data-visualization` `charts` `dashboard` `kpi` `data-analysis` `data-cleaning` `data-structuring` `segmentation` `ab-testing` `random-sampling` `sales-analysis` `business-analytics` `management-reporting` `data-analyst` `portfolio-project`
