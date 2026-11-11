// Query name: dim_bank
let
    Source = Csv.Document(File.Contents(DataFolder & "dim_bank.csv"),
                          [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"bank_id", Int64.Type}, {"bank_name", type text}, {"bank_type", type text}}, "en-US")
in
    Typed
