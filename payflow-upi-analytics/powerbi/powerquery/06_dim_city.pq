// Query name: dim_city
// city_id 0 is the placeholder "N/A (P2P)" used by the P2P merchant row.
let
    Source = Csv.Document(File.Contents(DataFolder & "dim_city.csv"),
                          [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {{"city_id", Int64.Type}, {"city", type text}, {"tier", type text}, {"region", type text}}, "en-US")
in
    Typed
