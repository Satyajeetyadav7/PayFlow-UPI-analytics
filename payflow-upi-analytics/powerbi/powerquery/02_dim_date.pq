// Query name: dim_date
// One row per calendar day. After loading: Table tools > Mark as date table > choose [date].
let
    Source = Csv.Document(File.Contents(DataFolder & "dim_date.csv"),
                          [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"date_key", Int64.Type}, {"date", type date}, {"year", Int64.Type}, {"quarter", type text},
        {"month_num", Int64.Type}, {"month_name", type text}, {"month_key", type text},
        {"day_of_month", Int64.Type}, {"day_name", type text}, {"dow", Int64.Type},
        {"is_weekend", Int64.Type}, {"is_festive", Int64.Type}, {"festive_name", type text},
        {"is_salary_window", Int64.Type}
    }, "en-US"),
    MonthStart = Table.AddColumn(Typed, "month_start", each Date.StartOfMonth([date]), type date),
    DayType = Table.AddColumn(MonthStart, "day_type", each if [is_festive] = 1 then "Festive" else if [is_weekend] = 1 then "Weekend" else "Weekday", type text)
in
    DayType
