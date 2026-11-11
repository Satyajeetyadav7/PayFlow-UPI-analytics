// Query name: fact_transactions
// Loads the transaction fact table, sets types and adds three banding columns used for slicing.
let
    Source = Csv.Document(File.Contents(DataFolder & "fact_transactions.csv"),
                          [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"txn_id", Int64.Type}, {"txn_ts", type datetime}, {"date_key", Int64.Type}, {"txn_hour", Int64.Type},
        {"customer_id", Int64.Type}, {"merchant_id", Int64.Type}, {"bank_id", Int64.Type},
        {"txn_type", type text}, {"network", type text}, {"amount", Int64.Type}, {"status", type text},
        {"failure_type", type text}, {"failure_reason", type text}, {"fraud_flag", Int64.Type},
        {"fraud_loss", Int64.Type}, {"cashback", type number}, {"response_ms", Int64.Type}, {"tenure_days", Int64.Type}
    }, "en-US"),
    HourBand = Table.AddColumn(Typed, "hour_band", each
        if [txn_hour] <= 4 then "1. Night (0-4)"
        else if [txn_hour] <= 8 then "2. Early (5-8)"
        else if [txn_hour] <= 11 then "3. Morning (9-11)"
        else if [txn_hour] <= 16 then "4. Day (12-16)"
        else if [txn_hour] <= 21 then "5. Evening (17-21)"
        else "6. Late (22-23)", type text),
    AmountBand = Table.AddColumn(HourBand, "amount_band", each
        if [amount] <= 500 then "1. Up to 500"
        else if [amount] <= 2000 then "2. 501-2,000"
        else if [amount] <= 10000 then "3. 2,001-10,000"
        else if [amount] <= 25000 then "4. 10,001-25,000"
        else "5. Above 25,000", type text),
    TenureBand = Table.AddColumn(AmountBand, "tenure_band", each
        if [tenure_days] <= 30 then "1. 0-30 days"
        else if [tenure_days] <= 90 then "2. 31-90 days"
        else if [tenure_days] <= 180 then "3. 91-180 days"
        else "4. 180+ days", type text)
in
    TenureBand
