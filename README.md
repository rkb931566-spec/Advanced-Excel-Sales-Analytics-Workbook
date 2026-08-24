# Advanced Excel — Sales Analytics Workbook

> A **working Excel workbook** on 57,542 real order lines, built with the formula patterns that carry actual analytical weight — `SUMPRODUCT`, `INDEX`/`MATCH`, `LARGE`, `PERCENTILE.INC`, structured references, two-way lookups — plus a **Power Query (M)** loader.

[![Excel](https://img.shields.io/badge/Excel-2016%20compatible-217346?style=flat&logo=microsoftexcel&logoColor=white)](./workbook/Sales_Analytics_Excel.xlsx)
[![Power Query](https://img.shields.io/badge/Power%20Query-M-F2C811?style=flat&logo=powerbi&logoColor=black)](./power_query/load_sales.m)
[![Data](https://img.shields.io/badge/Data-included%20in%20repo-34A853?style=flat)](./data/thelook_sales_2024_2025.csv)

---

## Business Context

**Industry:** E-commerce / Retail
**Stakeholders:** Commercial and Finance — the teams that live in spreadsheets
**Business question:** *Where does revenue come from, at what margin, and is growth coming from more orders or bigger ones?*

Excel is where most business stakeholders actually read numbers. This project treats it as a first-class analytical tool rather than an export format: the workbook computes everything itself, from a flat extract, with formulas that a reviewer can open and audit line by line.

---

## Dataset

| Field | Details |
|---|---|
| **Source** | `bigquery-public-data.thelook_ecommerce` — public, free to query, never expires |
| **Extract** | [`data/thelook_sales_2024_2025.csv`](./data/thelook_sales_2024_2025.csv) · 57,542 order lines · 14 columns |
| **Window** | 2024-01-01 to 2025-12-31 (two full years, so YoY is meaningful) |
| **Scale** | US$3.43M revenue · 39,743 orders · 51.9% gross margin |

Cancelled and Returned items are excluded at the source query — leaving them in would inflate revenue with sales that never completed.

---

## Headline results

| Metric | Value |
|---|---:|
| Revenue | US$3,426,968.93 |
| Profit | US$1,777,544.97 |
| Gross margin | 51.9% |
| Orders | 39,743 |
| Average order value | US$86.23 |
| Items per order | 1.45 |

**Revenue grew +54.8% YoY — and it came from volume, not from basket size.** Orders rose +52.6% while AOV moved only +1.5%. Margin held flat at 51.9% across both years, so the growth was not bought with discounting.

**The catalogue has a long tail, not a Pareto curve.** The top 5 of 26 categories carry 43.3% of revenue — assortment decisions here cannot lean on a handful of lines.

**It is priced for volume.** Nearly 62% of items sell below US$50 (P50 = US$39.99, P99 = US$299.95).

---

## What's in this repo

| File | Purpose |
|---|---|
| [`workbook/Sales_Analytics_Excel.xlsx`](./workbook/Sales_Analytics_Excel.xlsx) | The workbook — 3 sheets, 57,542 rows, all formulas live |
| [`docs/formula_reference.md`](./docs/formula_reference.md) | Every formula in the workbook, verbatim, with *why that pattern* |
| [`power_query/load_sales.m`](./power_query/load_sales.m) | Power Query (M) loader — parameterised path, locale-proof typing |
| [`sql/01_export_sales.sql`](./sql/01_export_sales.sql) | The BigQuery query that produces the extract |
| [`sql/02_control_totals.sql`](./sql/02_control_totals.sql) | Control totals, to check the workbook against the source |

---

## Techniques demonstrated

**Counting distinct orders in O(n).** Excel has no `COUNTDISTINCT`, and the usual `SUMPRODUCT(1/COUNTIF(range,range))` workaround is O(n²) — on 57,542 rows that is ~3.3 billion comparisons and the file hangs. Because the extract is sorted by `order_id`, `=IF($A2<>$A1,1,0)` summed down the column gives the distinct count in one pass.

**Revenue-weighted margin, not average margin.** `SUMPRODUCT(revenue, margin_pct)/SUM(revenue)` — averaging `margin_pct` directly would weight a US$5 sale the same as a US$500 one.

**Top-N without sorting.** `INDEX`/`MATCH`/`LARGE` with an anchored running total (`SUM($C$52:C52)`) produces a self-ranking Pareto that updates when the data changes.

**Two-way interactive lookup.** Nested `MATCH` calls feeding one `INDEX`, driven by data-validation dropdowns — interactive without a macro.

**A month spine that is a real date dimension**, so a month with no sales shows as zero instead of disappearing from the chart.

**Locale-proof loading.** The CSV is written by BigQuery with dot decimals and ISO dates; on a pt-BR machine a plain open turns `79.95` into `7995`. Every conversion in the M query is pinned to `en-US`.

**Excel 2016 compatibility, deliberately.** No `XLOOKUP`, no `LET`, no dynamic arrays. A portfolio file that opens with `#NAME?` errors on the reviewer's machine has failed before anyone reads a formula.

---

## Data quality notes

Three things in this source would silently distort a report built on it:

1. **Country names are not canonical** — the source carries both `Germany` and `Deutschland`, and `Brasil` rather than `Brazil`. A raw `GROUP BY` splits Germany across two rows and under-reports it. The workbook resolves it with an `INDEX`/`MATCH` mapping table on the `Lookups` sheet.
2. **Cancelled and Returned items are present** in `order_items` and are excluded at the source query.
3. **No item in this window sells at a loss** — 0 of 57,542. That is a property of the generated dataset, not a business finding, and it is stated so nobody reads the 100% "sold at a profit" tile as a real insight.

---

## Reproducing it

1. Run [`sql/01_export_sales.sql`](./sql/01_export_sales.sql) in BigQuery — no billing account needed, the dataset is public — and export to CSV. Or just use the CSV already in `data/`.
2. Open [`workbook/Sales_Analytics_Excel.xlsx`](./workbook/Sales_Analytics_Excel.xlsx). Everything recalculates from the `Data` sheet.
3. To refresh from a new extract instead, paste [`power_query/load_sales.m`](./power_query/load_sales.m) into Power Query and point the `p_DataPath` named cell at your file.

---

## Related projects

| Project | Angle |
|---|---|
| [ecommerce-sales-analysis](https://github.com/ANAPBORGES/ecommerce-sales-analysis) | The same theLook data in **BigQuery SQL** — RFM, cohorts, window functions |
| [saas-financial-kpis](https://github.com/ANAPBORGES/saas-financial-kpis) | **Power BI** — DAX, star schema, time intelligence |
| [tableau-sales-profitability-dashboard](https://github.com/ANAPBORGES/tableau-sales-profitability-dashboard) | **Tableau** — LOD, maps, what-if parameters |

---

<div align="center">

**Ana Paula Borges** · Senior Data Analyst & Team Leader
[LinkedIn](https://linkedin.com/in/ana-paula-d-araújo-borges) · [ap.daraujo@gmail.com](mailto:ap.daraujo@gmail.com)

</div>
