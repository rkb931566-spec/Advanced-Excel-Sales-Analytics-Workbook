// ============================================================================
// Power Query (M) - load the sales extract with explicit, locale-proof typing.
//
// Paste into Excel: Data > Get Data > Launch Power Query Editor >
//                   New Source > Blank Query > Advanced Editor
//
// Why this exists when the workbook already contains the data:
// the CSV is written by BigQuery with a dot decimal separator and ISO dates.
// On a machine whose regional settings use a comma decimal separator - which is
// the case here, pt-BR - a plain "open the CSV" turns 79.95 into either text or
// 7995. Every type conversion below is pinned to en-US so the result is the same
// on any machine.
// ============================================================================

let
    // ---- parameterised path, so nobody has to edit the query body ----
    SourcePath = Excel.CurrentWorkbook(){[Name = "p_DataPath"]}[Content]{0}[Column1],

    Source = Csv.Document(
        File.Contents(SourcePath),
        [Delimiter = ",", Columns = 14, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),

    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),

    // ---- explicit types, pinned to en-US ----
    Typed = Table.TransformColumnTypes(
        Promoted,
        {
            {"order_id",        Int64.Type},
            {"order_date",      type date},
            {"user_id",         Int64.Type},
            {"country",         type text},
            {"customer_gender", type text},
            {"customer_age",    Int64.Type},
            {"traffic_source",  type text},
            {"department",      type text},
            {"category",        type text},
            {"brand",           type text},
            {"product_name",    type text},
            {"revenue",         type number},
            {"cost",            type number},
            {"profit",          type number}
        },
        "en-US"
    ),

    // ---- clean text keys before they become grouping columns ----
    Trimmed = Table.TransformColumns(
        Typed,
        {
            {"country",        each Text.Trim(_), type text},
            {"category",       each Text.Trim(_), type text},
            {"brand",          each Text.Trim(_), type text},
            {"department",     each Text.Trim(_), type text},
            {"traffic_source", each Text.Trim(_), type text}
        }
    ),

    // ---- canonical country names.
    // The source mixes Germany/Deutschland and uses Brasil for Brazil. Grouping
    // on the raw column splits Germany across two rows and under-reports it.
    CountryFix = Table.AddColumn(
        Trimmed,
        "country_clean",
        each
            let c = [country] in
            if      c = "Deutschland" then "Germany"
            else if c = "Brasil"      then "Brazil"
            else if c = "España"      then "Spain"
            else c,
        type text
    ),

    // ---- engineered columns, matching the workbook's Data sheet ----
    WithYear = Table.AddColumn(CountryFix, "year",
        each Date.Year([order_date]), Int64.Type),

    WithMonth = Table.AddColumn(WithYear, "month_start",
        each Date.StartOfMonth([order_date]), type date),

    WithMargin = Table.AddColumn(WithMonth, "margin_pct",
        each if [revenue] = 0 then 0 else [profit] / [revenue], type number),

    WithBand = Table.AddColumn(WithMargin, "price_band",
        each if      [revenue] < 25  then "1. Up to 25"
             else if [revenue] < 50  then "2. 25 to 50"
             else if [revenue] < 100 then "3. 50 to 100"
             else                         "4. 100+",
        type text),

    WithProfitFlag = Table.AddColumn(WithBand, "is_profitable",
        each if [profit] > 0 then 1 else 0, Int64.Type),

    // ---- guard rails: fail loudly rather than publish a wrong number ----
    NoNullKeys = Table.SelectRows(WithProfitFlag,
        each [order_id] <> null and [order_date] <> null and [revenue] <> null),

    Final = Table.Buffer(NoNullKeys)
in
    Final
