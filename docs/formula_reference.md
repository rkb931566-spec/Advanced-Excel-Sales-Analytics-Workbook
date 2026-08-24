# Formula reference

Every formula below is copied verbatim out of
[`workbook/Sales_Analytics_Excel.xlsx`](../workbook/Sales_Analytics_Excel.xlsx).
They are not illustrations — they are what the file actually contains.

The workbook targets **Excel 2016**: no dynamic arrays, no `XLOOKUP`, no `LET`.
A file that opens with `#NAME?` errors on someone else's machine has failed
before anyone reads a formula, and the classic forms below run everywhere.

---

## Engineered columns — `Data` sheet

The source extract has 14 columns. Six more are computed inside the table, which
keeps the joins and the reporting logic in one visible place instead of buried in
a query.

| Column | Formula | What it is for |
|---|---|---|
| `year` | `=YEAR($B2)` | Grouping key for YoY |
| `month_start` | `=EOMONTH($B2,-1)+1` | First day of the month. `EOMONTH(date,-1)+1` is the classic way to get it without `DATE(YEAR(),MONTH(),1)` |
| `margin_pct` | `=IFERROR($N2/$L2,0)` | Line-level margin, divide-by-zero safe |
| `price_band` | `=IF($L2<25,"1. Up to 25",IF($L2<50,"2. 25 to 50",IF($L2<100,"3. 50 to 100","4. 100+")))` | Nested IF, deliberately in place of `IFS` — see note below |
| `is_profitable` | `=IF($N2>0,1,0)` | 1/0 flag, so it can be `SUM`med and `AVERAGE`d as a rate |
| `is_new_order` | `=IF($A2<>$A1,1,0)` | **Distinct order counter.** See below |

### Counting distinct orders without a pivot

Excel has no `COUNTDISTINCT`. The usual workaround,
`SUMPRODUCT(1/COUNTIF(range,range))`, is O(n²) — on 57,542 rows that is roughly
3.3 billion comparisons and the workbook would hang.

Because the extract is sorted by `order_id` (`ORDER BY oi.order_id` in the source
query), every line of an order is adjacent to the rest. So a single comparison
against the row above is enough:

```excel
=IF($A2<>$A1,1,0)
```

Summing that column gives 39,743 distinct orders. It is O(n), and it survives
filtering because it lives in the table.

The trick has one precondition — **the data must stay sorted by `order_id`** —
and that precondition is enforced upstream in the SQL, not assumed.

### Why nested IF instead of IFS

`IFS` does not exist in Excel 2016. Nested `IF` is longer to read but runs on
any version, which matters more for a file that gets shared.

---

## 1. Core KPIs — structured references

Structured references (`tbl_Sales[revenue]`) rather than `L2:L57543`, so the
formulas keep working when rows are added and read as English when reviewed.

```excel
Total revenue      =SUM(tbl_Sales[revenue])
Total profit       =SUM(tbl_Sales[profit])
Gross margin %     =IFERROR(B6/B5,0)
Line items         =ROWS(tbl_Sales)
Distinct orders    =SUM(tbl_Sales[is_new_order])
Average order value=IFERROR(B5/B9,0)
Average item price =AVERAGE(tbl_Sales[revenue])
Items per order    =IFERROR(B8/B9,0)
Sold at a profit % =AVERAGE(tbl_Sales[is_profitable])
```

`AVERAGE` over a 1/0 flag returns the rate directly — no `COUNTIF`/`COUNT` pair
needed.

### Distinct count on a small range

```excel
=SUMPRODUCT(1/COUNTIF(dim_category,dim_category))
```

The classic array-style distinct count, returning 26. Safe here because
`dim_category` is 26 cells. Deliberately **not** used on the 57k-row table, for
the reason above — knowing when a technique does not scale matters more than
knowing the technique.

---

## 2. Year over year

```excel
Revenue  =SUMIFS(tbl_Sales[revenue],tbl_Sales[year],$A18)
Profit   =SUMIFS(tbl_Sales[profit],tbl_Sales[year],$A18)
Orders   =SUMIFS(tbl_Sales[is_new_order],tbl_Sales[year],$A18)
Margin   =IFERROR(C18/B18,0)
AOV      =IFERROR(B18/E18,0)
YoY      =IFERROR(B19/B18-1,"")
```

The year sits in a cell (`$A18`), not inside the formula, so the block extends
to a new year by typing the year.

---

## 3. Category matrix

```excel
Category   =Lookups!A2
Revenue    =SUMIFS(tbl_Sales[revenue],tbl_Sales[category],$A23)
Profit     =SUMIFS(tbl_Sales[profit],tbl_Sales[category],$A23)
Margin %   =IFERROR(C23/B23,0)
% of total =IFERROR(B23/SUM($B$23:$B$48),0)
Rank       =RANK(B23,$B$23:$B$48)
Items      =COUNTIFS(tbl_Sales[category],$A23)
```

The category list is a dimension on the `Lookups` sheet, not a hand-typed list —
so a new category appears by extending one range, and the mixed
relative/absolute referencing (`$A23` vs `$B$23:$B$48`) lets the whole block fill
down from one cell.

Conditional formatting: data bars on the revenue column.

---

## 4. Top 10 — LARGE + INDEX/MATCH

The classic top-N pattern, no sorting and no helper column:

```excel
Category    =INDEX($A$23:$A$48,MATCH(LARGE($B$23:$B$48,$A52),$B$23:$B$48,0))
Revenue     =LARGE($B$23:$B$48,$A52)
Cumulative  =SUM($C$52:C52)/SUM($B$23:$B$48)
```

`LARGE(range,n)` pulls the nth largest value; `MATCH` finds where it sits;
`INDEX` returns the label beside it. The rank `n` comes from column A, so the
list re-ranks itself when the data changes.

`SUM($C$52:C52)` — anchored start, relative end — is the running-total idiom
that produces the Pareto curve as it fills down.

---

## 5. Two-way lookup with dropdowns

```excel
=INDEX($B$23:$G$48,
       MATCH($B$64,$A$23:$A$48,0),
       MATCH($B$65,$B$22:$G$22,0))
```

Two `MATCH` calls — one down the category column, one across the header row —
feed `INDEX` to pick any cell in the matrix. Both inputs are data-validation
dropdowns (`B64` bound to the `dim_category` named range, `B65` to a metric
list), so the block is interactive without a macro or a slicer.

---

## 6. SUMPRODUCT — conditions without helper columns

```excel
Revenue 2025, Women     =SUMPRODUCT((tbl_Sales[year]=2025)*(tbl_Sales[department]="Women")*tbl_Sales[revenue])
Revenue 2025, item >100 =SUMPRODUCT((tbl_Sales[year]=2025)*(tbl_Sales[revenue]>100)*tbl_Sales[revenue])
Revenue-weighted margin =SUMPRODUCT(tbl_Sales[revenue],tbl_Sales[margin_pct])/SUM(tbl_Sales[revenue])
Orders from Search 2025 =SUMPRODUCT((tbl_Sales[year]=2025)*(tbl_Sales[traffic_source]="Search")*tbl_Sales[is_new_order])
```

Multiplying boolean arrays gives AND logic; the products act as a mask. The
weighted-margin line is the one that matters analytically — averaging
`margin_pct` directly would weight a 5 sale the same as a 500 one and give the
wrong number.

---

## 7. Distribution

```excel
Median (P50) =PERCENTILE.INC(tbl_Sales[revenue],0.5)
P75          =PERCENTILE.INC(tbl_Sales[revenue],0.75)
P90          =PERCENTILE.INC(tbl_Sales[revenue],0.9)
P99          =PERCENTILE.INC(tbl_Sales[revenue],0.99)
IQR          =QUARTILE.INC(tbl_Sales[revenue],3)-QUARTILE.INC(tbl_Sales[revenue],1)
```

Banded histogram with `COUNTIFS` — first band, then the repeating band:

```excel
=COUNTIFS(Data!$L$2:$L$57543,"<="&$B83)
=COUNTIFS(Data!$L$2:$L$57543,">"&$B83,Data!$L$2:$L$57543,"<="&$B84)
```

The bounds live in cells and are concatenated into the criteria with `&`, so the
bands are editable without touching a formula. The counts sum to exactly 57,542.

`FREQUENCY` as a multi-cell array formula would be the textbook approach. It is
not used here because a formula that has to be re-entered with Ctrl+Shift+Enter
is one the next person will eventually break — `COUNTIFS` survives editing.

---

## 8. Country normalisation

```excel
Canonical  =INDEX(Lookups!$F$2:$F$16,MATCH($A95,Lookups!$E$2:$E$16,0))
Revenue    =SUMIFS(tbl_Sales[revenue],tbl_Sales[country],$A95)
Germany    =SUMIF($B$95:$B$109,"Germany",$C$95:$C$109)
```

The source carries both `Germany` and `Deutschland`. `INDEX`/`MATCH` against a
mapping table resolves them, and the final `SUMIF` rolls the canonical name back
up — **139,844.20** consolidated, against the split figure a naive `GROUP BY`
would report.

The mapping is a table on a sheet, not a nested `IF` chain, so a business user
can extend it.

---

## 9. Monthly trend

```excel
Revenue    =SUMIFS(tbl_Sales[revenue],tbl_Sales[month_start],$A114)
Orders     =SUMIFS(tbl_Sales[is_new_order],tbl_Sales[month_start],$A114)
AOV        =IFERROR(B114/C114,0)
MoM %      =IFERROR(B115/B114-1,"")
3-mo avg   =IF(ROW()-ROW($A$114)<2,"",AVERAGE(B112:B114))
Cumulative =SUM($B$114:B114)
```

The moving average guards its own window: `ROW()-ROW($A$114)<2` suppresses the
first two months rather than averaging over rows that are not there — the common
bug in hand-built moving averages.

The month spine is a real 24-row date dimension on `Lookups`, so a month with no
sales shows as a zero instead of vanishing from the chart.
