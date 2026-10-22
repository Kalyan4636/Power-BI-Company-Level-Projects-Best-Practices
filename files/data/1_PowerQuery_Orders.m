// Power BI > Home > Transform data > New Source > Blank Query > Advanced Editor. Paste this, edit the file path, name the query "Orders".
let
    Source   = Csv.Document(File.Contents("C:\YOUR_PATH\Power_BI_Dataset_xlsx_-_List_of_Orders.csv"), [Delimiter=",", Columns=16, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    // Parse BOTH date styles (01-03-2017 and 1/13/2017), month-first, en-US
    Dated    = Table.TransformColumns(Promoted, {{"Order Date", each Date.FromText(Text.Replace(_, "-", "/"), [Culture="en-US"]), type date}}),
    // Fix "United Kingdom Excel Excel Excel" and trim all text
    Country  = Table.TransformColumns(Dated, {{"Country", each Text.Trim(Text.Replace(_, " Excel", "")), type text}}),
    Trimmed  = Table.TransformColumns(Country, List.Transform({"Customer Name","City","State","Region","Segment","Ship Mode","Category","Sub-Category","Product Name"}, each {_, Text.Trim, type text})),
    Typed    = Table.TransformColumnTypes(Trimmed, {{"Order ID", Int64.Type}, {"Cost", Int64.Type}, {"Sales", Int64.Type}, {"Profit", Int64.Type}, {"Quantity", Int64.Type}}),
    // Profit must equal Sales - Cost (fixes Order ID 1 sign error)
    ProfitFx = Table.ReplaceValue(Typed, each [Profit], each [Sales] - [Cost], Replacer.ReplaceValue, {"Profit"}),
    // Remove the 14 test rows (Sales 100,000 / 1,000,000 on 31-Dec-2020). Largest genuine sale is 6,517.
    NoOutlier = Table.SelectRows(ProfitFx, each [Sales] < 100000)
in
    NoOutlier
