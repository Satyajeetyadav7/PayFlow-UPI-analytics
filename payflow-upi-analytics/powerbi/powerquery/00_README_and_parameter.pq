// ============================================================================
// POWER QUERY SETUP  (Home > Transform data > New Source > Blank Query > Advanced Editor)
// ============================================================================
// 1. Create a Parameter first:  Manage Parameters > New Parameter
//      Name:          DataFolder
//      Type:          Text
//      Current value: the folder that holds the six CSV files, ENDING with a slash, e.g.
//                     C:\Users\you\payflow-upi-analytics\powerbi\data\
//
// 2. For each of the six files in this folder (01_... to 06_...) create a Blank Query,
//    open Advanced Editor, paste the file's code, and rename the query exactly as shown in
//    the first comment line of that file (fact_transactions, dim_date, ...).
//
// 3. Close & Apply. Then follow powerbi/BUILD_GUIDE.md for relationships and measures.
//
// OPTIONAL - load straight from GitHub so the report refreshes in the Power BI service without a
// gateway (works only if your repository is public). In each file, replace the Source line with:
//
//    Source = Csv.Document(Web.Contents("https://raw.githubusercontent.com/YOUR-USERNAME/payflow-upi-analytics/main/powerbi/data/fact_transactions.csv"),
//             [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
//
// and set the data source privacy level to Anonymous when prompted.
// ============================================================================
