$P = 'D:\Onedrive\MCI Excel\1_Sessions\B4_Power_Query\Practice\'
$q = [ordered]@{}
$q['Sales_T1'] = @"
let
    Source = Csv.Document(File.Contents("${P}Data_Sales\Sales_2025-01.csv"),[Delimiter=",", Columns=8, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type with Locale" = Table.TransformColumnTypes(#"Promoted Headers", {{"Order_Date", type date}}, "vi-VN"),
    #"Changed Type with Locale1" = Table.TransformColumnTypes(#"Changed Type with Locale", {{"Unit_Price", Int64.Type}}, "vi-VN"),
    #"Trimmed Text" = Table.TransformColumns(#"Changed Type with Locale1",{{"Store_Code", Text.Trim, type text}}),
    #"Removed Duplicates" = Table.Distinct(#"Trimmed Text")
in
    #"Removed Duplicates"
"@
$q['Product'] = @"
let
    Source = Excel.Workbook(File.Contents("${P}B4_Lookups.xlsx"), null, true),
    Navigation = Source{[Item="Product",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Navigation,{{"Product_Code", type text}, {"Product_Name", type text}, {"Category", type text}, {"Unit_Price", Int64.Type}, {"Unit_Cost", Int64.Type}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"Product_Code"}),
    #"Counted Rows" = Table.RowCount(#"Removed Duplicates")
in
    #"Counted Rows"
"@
$q['Store'] = @"
let
    Source = Excel.Workbook(File.Contents("${P}B4_Lookups.xlsx"), null, true),
    Navigation = Source{[Item="Store",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Navigation,{{"Store_Code", type text}, {"Store_Name", type text}, {"City", type text}, {"Region", type text}, {"Store_Type", type text}, {"Open_Date", type date}})
in
    #"Changed Type"
"@
$q['Customer'] = @"
let
    Source = Excel.Workbook(File.Contents("${P}B4_Lookups.xlsx"), null, true),
    Navigation = Source{[Item="Customer",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Navigation,{{"Customer_ID", type text}, {"Full_Name", type text}, {"Gender", type text}, {"Birth_Year", Int64.Type}, {"Segment", type text}, {"Join_Date", type date}, {"Home_Store_Code", type text}})
in
    #"Changed Type"
"@
$xlM = New-Object -ComObject Excel.Application
$xlM.Visible = $true
$wbM = $xlM.Workbooks.Open("$Base\B4_Chup_M3.xlsx")
foreach ($k in $q.Keys) { [void]$wbM.Queries.Add($k, $q[$k]) }
Start-Sleep -Seconds 1
$xlM.WindowState = -4137
"queries=" + $wbM.Queries.Count
