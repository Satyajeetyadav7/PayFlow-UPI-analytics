// Query name: dim_merchant
// merchant_id 0 is the placeholder "P2P (Individual)" used for person-to-person transfers.
let
    Source = Csv.Document(File.Contents(DataFolder & "dim_merchant.csv"),
                          [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"merchant_id", Int64.Type}, {"merchant_name", type text}, {"category", type text},
        {"city_id", Int64.Type}, {"merchant_size", type text}
    }, "en-US")
in
    Typed
