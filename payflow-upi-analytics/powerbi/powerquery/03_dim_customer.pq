// Query name: dim_customer
let
    Source = Csv.Document(File.Contents(DataFolder & "dim_customer.csv"),
                          [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"customer_id", Int64.Type}, {"signup_date", type date}, {"city_id", Int64.Type}, {"age_group", type text},
        {"device_os", type text}, {"primary_bank_id", Int64.Type}, {"kyc_level", type text}
    }, "en-US")
in
    Typed
