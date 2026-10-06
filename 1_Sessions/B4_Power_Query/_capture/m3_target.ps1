$P = 'D:\Onedrive\MCI Excel\1_Sessions\B4_Power_Query\Practice\'
$TBefore = @"
let
    Source = Excel.Workbook(File.Contents("${P}B4_Target_2025.xlsx"), null, true),
    Navigation = Source{[Item="Target_2025",Kind="Sheet"]}[Data]
in
    Navigation
"@
$TAfter = @"
let
    Source = Excel.Workbook(File.Contents("${P}B4_Target_2025.xlsx"), null, true),
    Navigation = Source{[Item="Target_2025",Kind="Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Navigation,2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Promoted Headers", {"Store_Code", "Store_Name"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Thang"}, {"Value", "Target"}}),
    #"Extracted Text After Delimiter" = Table.TransformColumns(#"Renamed Columns", {{"Thang", each Text.AfterDelimiter(_, " "), type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Extracted Text After Delimiter",{{"Store_Code", type text}, {"Thang", Int64.Type}, {"Target", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Store_Name"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "Ngay_Dau_Thang", each #date(2025, [Thang], 1), type date)
in
    #"Added Custom"
"@
$ProductNoCount = @"
let
    Source = Excel.Workbook(File.Contents("${P}B4_Lookups.xlsx"), null, true),
    Navigation = Source{[Item="Product",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Navigation,{{"Product_Code", type text}, {"Product_Name", type text}, {"Category", type text}, {"Unit_Price", Int64.Type}, {"Unit_Cost", Int64.Type}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"Product_Code"})
in
    #"Removed Duplicates"
"@
