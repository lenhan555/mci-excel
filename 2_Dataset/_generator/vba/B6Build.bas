Attribute VB_Name = "B6Build"
Option Explicit
' Dung Practice B6 trong B6_MCI_Mart_Dashboard_Answer.xlsx (da SaveAs tu dap an B5).
' Chay tu Immediate window: B6Build.RunAll   (file nay sinh boi gen_b6_vba.py; nhat ky o %TEMP%\b6_build_log.txt)
' Cach dung: mo B6_MCI_Mart_Dashboard_Answer.xlsx (da SaveAs tu dap an B5) + mot workbook trong de chua macro, Alt+F11, Ctrl+M nap file .bas, chay RunAll.

Private wb As Workbook
Private gF As Integer
Private Const CN As String = "ThisWorkbookDataModel"
Private Const FILE_NAME As String = "B6_MCI_Mart_Dashboard_Answer.xlsx"
Private Const XL_WATERFALL As Long = 119
Private Const XL_REGIONMAP As Long = 140
Private Const SL As String = ",""[Calendar].[Nam].&[2025]"",Slicer_Thang,Slicer_Region"

Private Function LogPath() As String
    LogPath = Environ("TEMP") & "\b6_build_log.txt"
End Function

Private Sub L(ByVal s As String)
    Print #gF, s
End Sub

Private Sub Chk(ByVal stepName As String)
    If Err.Number <> 0 Then
        L "ERR [" & stepName & "] " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
End Sub

Private Function CVF(ByVal meas As String, ByVal args As String) As String
    CVF = "=CUBEVALUE(""" & CN & """,""[Measures].[" & meas & "]""" & args & ")"
End Function

Private Function GetSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nm
    End If
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    Set GetSheet = ws
End Function

Private Sub DelSheet(ByVal nm As String)
    On Error Resume Next
    wb.Worksheets(nm).Delete
    Err.Clear
End Sub

Private Sub Title(ws As Worksheet, ByVal txt As String, ByVal note As String)
    ws.Range("A1").Value = txt
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Color = RGB(32, 56, 100)
    If Len(note) > 0 Then
        ws.Range("A2").Value = note
        ws.Range("A2").Font.Italic = True
        ws.Range("A2").Font.Color = RGB(88, 88, 88)
    End If
End Sub

Private Sub Hdr(rng As Range)
    rng.Font.Bold = True
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = RGB(32, 56, 100)
    rng.VerticalAlignment = xlCenter
End Sub

Public Sub RunAll()
    RunFrom 1
End Sub

Public Sub RunFrom(ByVal n As Long)
    gF = FreeFile
    If n = 1 Then
        Open LogPath() For Output As #gF
    Else
        Open LogPath() For Append As #gF
    End If
    L "START stage " & n & " " & Now
    Set wb = Workbooks(FILE_NAME)
    wb.Activate
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    If n <= 1 Then Stage1_Cleanup
    If n <= 2 Then Stage2_Settings
    If n <= 3 Then Stage3_Pivots
    If n <= 4 Then Stage4_DataDash
    If n <= 5 Then Stage5_Dashboard
    If n <= 6 Then Stage6_Brief
    If n <= 7 Then Stage7_DuBao
    If n <= 8 Then Stage8_KiemTra
    Stage9_Finish
    Application.DisplayAlerts = True
    L "DONE " & Now
    Close #gF
End Sub

Private Sub Stage1_Cleanup()
    Dim i As Long
    On Error Resume Next
    For i = wb.SlicerCaches.Count To 1 Step -1
        If wb.SlicerCaches(i).Name = "Slicer_Thang" Or wb.SlicerCaches(i).Name = "Slicer_Region" Then wb.SlicerCaches(i).Delete
        Chk "del slicer cache"
    Next
    DelSheet "Dashboard"
    DelSheet "Data_Dash"
    DelSheet "Cai_dat"
    DelSheet "Brief"
    DelSheet "Du_bao"
    DelSheet "Kiem_tra_B6"
    DelSheet "PV_Dash"
    L "stage1 ok"
End Sub

Private Sub Stage2_Settings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetSheet("PV_Dash")
    Set ws = GetSheet("Kiem_tra_B6")
    Set ws = GetSheet("Du_bao")
    Set ws = GetSheet("Brief")
    Set ws = GetSheet("Data_Dash")
    Set ws = GetSheet("Dashboard")
    Chk "sheets"
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:S").ColumnWidth = 9.5
    ws.Columns("T").ColumnWidth = 2
    ws.Columns("U:X").ColumnWidth = 9.5
    Chk "dash widths"
    Set ws = GetSheet("Cai_dat")
    Title ws, "C" & ChrW(224) & "i " & ChrW(273) & ChrW(7863) & "t dashboard: target b" & ChrW(7893) & " sung v" & ChrW(224) & " ng" & ChrW(432) & ChrW(7905) & "ng RAG", "M" & ChrW(7885) & "i con s" & ChrW(7889) & " dashboard d" & ChrW(249) & "ng l" & ChrW(224) & "m m" & ChrW(7889) & "c so s" & ChrW(225) & "nh n" & ChrW(7857) & "m " & ChrW(7903) & " " & ChrW(273) & ChrW(226) & "y. " & ChrW(272) & ChrW(7893) & "i m" & ChrW(7897) & "t n" & ChrW(417) & "i, m" & ChrW(7885) & "i th" & ChrW(7867) & " KPI " & ChrW(273) & ChrW(7893) & "i theo. " & ChrW(212) & " n" & ChrW(7873) & "n v" & ChrW(224) & "ng l" & ChrW(224) & " " & ChrW(244) & " nh" & ChrW(7853) & "p."
    ws.Range("A4:D4").Value = Array("KPI", ChrW(272) & ChrW(417) & "n v" & ChrW(7883), "Target", "Ghi ch" & ChrW(250))
    Hdr ws.Range("A4:D4")
    ws.Range("A5:D5").Value = Array("S" & ChrW(7889) & " " & ChrW(273) & ChrW(417) & "n m" & ChrW(7895) & "i th" & ChrW(225) & "ng", ChrW(273) & ChrW(417) & "n", 1000, "Target doanh thu l" & ChrW(7845) & "y t" & ChrW(7915) & " measure [Total Target] trong Data Model, kh" & ChrW(244) & "ng nh" & ChrW(7853) & "p " & ChrW(7903) & " " & ChrW(273) & ChrW(226) & "y")
    ws.Range("A6:D6").Value = Array("Gi" & ChrW(225) & " tr" & ChrW(7883) & " " & ChrW(273) & ChrW(417) & "n trung b" & ChrW(236) & "nh (AOV)", "VND", 300000, "")
    ws.Range("A7:D7").Value = Array("Bi" & ChrW(234) & "n l" & ChrW(7907) & "i nhu" & ChrW(7853) & "n g" & ChrW(7897) & "p (GM%)", "%", 0.33, "")
    ws.Range("A8:D8").Value = Array("T" & ChrW(259) & "ng tr" & ChrW(432) & ChrW(7903) & "ng so c" & ChrW(249) & "ng k" & ChrW(7923) & " (YoY)", "%", 0.1, "M" & ChrW(7913) & "c " & ChrW(273) & ChrW(7841) & "t = YoY " & ChrW(247) & " target (v" & ChrW(237) & " d" & ChrW(7909) & " " & ChrW(8722) & "0,9% " & ChrW(247) & " 10% = " & ChrW(8722) & "9%)")
    ws.Range("A9:D9").Value = Array("Kh" & ChrW(225) & "ch ho" & ChrW(7841) & "t " & ChrW(273) & ChrW(7897) & "ng m" & ChrW(7895) & "i th" & ChrW(225) & "ng", "kh" & ChrW(225) & "ch", 700, "Target theo th" & ChrW(225) & "ng; dashboard nh" & ChrW(236) & "n m" & ChrW(7897) & "t th" & ChrW(225) & "ng m" & ChrW(7897) & "t l" & ChrW(7847) & "n")
    ws.Range("C5").NumberFormat = "#,##0"
    ws.Range("C6").NumberFormat = "#,##0"
    ws.Range("C7:C8").NumberFormat = "0%"
    ws.Range("C9").NumberFormat = "#,##0"
    ws.Range("C5:C9").Interior.Color = RGB(255, 242, 204)
    ws.Range("A11").Value = "Ng" & ChrW(432) & ChrW(7905) & "ng RAG (m" & ChrW(7913) & "c " & ChrW(273) & ChrW(7841) & "t = gi" & ChrW(225) & " tr" & ChrW(7883) & " " & ChrW(247) & " target)"
    ws.Range("A11").Font.Bold = True
    ws.Range("A12:D12").Value = Array("Tr" & ChrW(7841) & "ng th" & ChrW(225) & "i", "Bi" & ChrW(7875) & "u t" & ChrW(432) & ChrW(7907) & "ng", "Ng" & ChrW(432) & ChrW(7905) & "ng t" & ChrW(7915), "Ghi ch" & ChrW(250))
    Hdr ws.Range("A12:D12")
    ws.Range("A13:D13").Value = Array(ChrW(272) & ChrW(7841) & "t", ChrW(10004), 1, "M" & ChrW(7913) & "c " & ChrW(273) & ChrW(7841) & "t t" & ChrW(7915) & " 100% tr" & ChrW(7903) & " l" & ChrW(234) & "n")
    ws.Range("A14:D14").Value = Array("C" & ChrW(7843) & "nh b" & ChrW(225) & "o", ChrW(9650), 0.8, "T" & ChrW(7915) & " 80% " & ChrW(273) & ChrW(7871) & "n d" & ChrW(432) & ChrW(7899) & "i 100%")
    ws.Range("A15:D15").Value = Array("Kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(7841) & "t", ChrW(10008), "", "D" & ChrW(432) & ChrW(7899) & "i ng" & ChrW(432) & ChrW(7905) & "ng C" & ChrW(7843) & "nh b" & ChrW(225) & "o")
    ws.Range("C13:C14").NumberFormat = "0%"
    ws.Range("C13:C14").Interior.Color = RGB(255, 242, 204)
    ws.Range("B13:B15").HorizontalAlignment = xlCenter
    ws.Range("B13").Font.Color = RGB(12, 163, 12)
    ws.Range("B14").Font.Color = RGB(250, 178, 25)
    ws.Range("B15").Font.Color = RGB(208, 59, 59)
    ws.Range("A17").Value = "Quy " & ChrW(432) & ChrW(7899) & "c: kh" & ChrW(244) & "ng d" & ChrW(249) & "ng m" & ChrW(224) & "u " & ChrW(273) & ChrW(417) & "n " & ChrW(273) & ChrW(7897) & "c. M" & ChrW(7895) & "i tr" & ChrW(7841) & "ng th" & ChrW(225) & "i c" & ChrW(243) & " bi" & ChrW(7875) & "u t" & ChrW(432) & ChrW(7907) & "ng kh" & ChrW(225) & "c h" & ChrW(236) & "nh (" & ChrW(10004) & " " & ChrW(9650) & " " & ChrW(10008) & ") v" & ChrW(224) & " nh" & ChrW(227) & "n ch" & ChrW(7919) & "."
    ws.Range("A18").Value = "M" & ChrW(224) & "u theo Design System c" & ChrW(7911) & "a kh" & ChrW(243) & "a: " & ChrW(272) & ChrW(7841) & "t #0CA30C " & ChrW(183) & " C" & ChrW(7843) & "nh b" & ChrW(225) & "o #FAB219 " & ChrW(183) & " Kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(7841) & "t #D03B3B."
    ws.Range("A19").Value = "Quy t" & ChrW(7855) & "c Icon Set tr" & ChrW(234) & "n Dashboard (Type = Formula) tr" & ChrW(7887) & " t" & ChrW(7899) & "i C13 v" & ChrW(224) & " C14 c" & ChrW(7911) & "a sheet n" & ChrW(224) & "y."
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 70
    Chk "cai_dat"
    L "stage2 ok"
End Sub

Private Function NewPT(ws As Worksheet, ByVal addr As String, ByVal nm As String, firstPT As PivotTable) As PivotTable
    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create(SourceType:=xlExternal, SourceData:=wb.Connections(CN), Version:=6)
    Set NewPT = pc.CreatePivotTable(TableDestination:=ws.Range(addr), TableName:=nm)
End Function

Private Sub PageNam(pt As PivotTable)
    pt.CubeFields("[Calendar].[Nam]").Orientation = xlPageField
    pt.PivotFields("[Calendar].[Nam].[Nam]").CurrentPageName = "[Calendar].[Nam].&[2025]"
End Sub

Private Sub Stage3_Pivots()
    Dim wsP As Worksheet, wsD As Worksheet
    Dim ptLine As PivotTable, ptCombo As PivotTable, ptTop As PivotTable
    Dim scT As SlicerCache, scR As SlicerCache, sl As Slicer
    On Error Resume Next
    Set wsP = GetSheet("PV_Dash")
    Set wsD = GetSheet("Dashboard")
    Title wsP, "PivotTable ngu" & ChrW(7891) & "n cho PivotChart v" & ChrW(224) & " Slicer c" & ChrW(7911) & "a Dashboard", "Kh" & ChrW(244) & "ng ch" & ChrW(7881) & "nh tay. M" & ChrW(7895) & "i PivotTable l" & ChrW(7845) & "y t" & ChrW(7915) & " Data Model (ThisWorkbookDataModel), l" & ChrW(7885) & "c N" & ChrW(259) & "m = 2025."
    ' --- pvtNgay_Dash: doanh thu theo ngay (line) ---
    Set ptLine = NewPT(wsP, "A8", "pvtNgay_Dash", Nothing)
    Chk "pt line create"
    PageNam ptLine
    Chk "pt line nam"
    ptLine.CubeFields("[Calendar].[Date]").Orientation = xlRowField
    ptLine.CubeFields("[Measures].[Total Revenue]").Orientation = xlDataField
    Chk "pt line fields"
    L "ptLine cacheIndex=" & ptLine.CacheIndex
    ' --- pvtCombo_Dash: thuc te vs target theo thang (combo) ---
    Set ptCombo = NewPT(wsP, "F8", "pvtCombo_Dash", ptLine)
    Chk "pt combo create"
    PageNam ptCombo
    ptCombo.CubeFields("[Calendar].[Thang]").Orientation = xlRowField
    ptCombo.CubeFields("[Measures].[Total Revenue]").Orientation = xlDataField
    ptCombo.CubeFields("[Measures].[Total Target]").Orientation = xlDataField
    Chk "pt combo fields"
    ptCombo.PivotFields("[Calendar].[Thang].[Thang]").VisibleItemsList = Array("[Calendar].[Thang].&[1]", "[Calendar].[Thang].&[2]", "[Calendar].[Thang].&[3]", "[Calendar].[Thang].&[4]")
    Chk "pt combo 4 thang"
    L "ptCombo cacheIndex=" & ptCombo.CacheIndex
    ' --- pvtTop5_Dash: top 5 san pham ---
    Set ptTop = NewPT(wsP, "L8", "pvtTop5_Dash", ptLine)
    Chk "pt top create"
    PageNam ptTop
    ptTop.CubeFields("[Product].[Product_Name]").Orientation = xlRowField
    ptTop.CubeFields("[Measures].[Total Revenue]").Orientation = xlDataField
    Chk "pt top fields"
    L "ptTop cacheIndex=" & ptTop.CacheIndex
    ptTop.PivotFields("[Product].[Product_Name].[Product_Name]").PivotFilters.Add2 Type:=xlTopCount, DataField:=ptTop.PivotFields("[Measures].[Total Revenue]"), Value1:=5
    Chk "pt top filter"
    ptTop.PivotFields("[Product].[Product_Name].[Product_Name]").AutoSort xlDescending, "[Measures].[Total Revenue]"
    Chk "pt top sort"
    ' --- Slicer ---
    Set scT = wb.SlicerCaches.Add2(ptLine, "[Calendar].[Thang]")
    Chk "slicer cache thang"
    scT.Name = "Slicer_Thang"
    Chk "slicer name thang"
    Set scR = wb.SlicerCaches.Add2(ptLine, "[Store].[Region]")
    Chk "slicer cache region"
    scR.Name = "Slicer_Region"
    Chk "slicer name region"
    scT.PivotTables.AddPivotTable ptTop
    Chk "conn thang-top"
    scR.PivotTables.AddPivotTable ptTop
    Chk "conn region-top"
    scR.PivotTables.AddPivotTable ptCombo
    Chk "conn region-combo"
    Set sl = scT.Slicers.Add(wsD, , "Thang", "Th" & ChrW(225) & "ng", wsD.Range("U5").Top, wsD.Range("U5").Left, wsD.Range("U5:X5").Width, 175)
    Chk "slicer thang add"
    sl.NumberOfColumns = 4
    sl.Style = "SlicerStyleLight1"
    Chk "slicer thang fmt"
    Set sl = scR.Slicers.Add(wsD, , "Region", "Mi" & ChrW(7873) & "n", wsD.Range("U17").Top, wsD.Range("U17").Left, wsD.Range("U17:X17").Width, 120)
    Chk "slicer region add"
    sl.Style = "SlicerStyleLight1"
    scT.VisibleSlicerItemsList = Array("[Calendar].[Thang].&[4]")
    Chk "slicer thang select 4"
    L "slicer caches: " & wb.SlicerCaches.Count
    Dim i As Long
    For i = 1 To wb.SlicerCaches.Count
        L "  cache " & wb.SlicerCaches(i).Name & " pivots=" & wb.SlicerCaches(i).PivotTables.Count
    Next
    L "stage3 ok"
End Sub

Private Sub Stage4_DataDash()
    Dim ws As Worksheet, k As Long, m As Long, r As Long
    Dim kpi, meas, tgt, fmt
    On Error Resume Next
    Set ws = GetSheet("Data_Dash")
    Title ws, "Data_Dash: s" & ChrW(7889) & " li" & ChrW(7879) & "u th" & ChrW(244) & " cho Dashboard (CUBEVALUE t" & ChrW(7915) & " Data Model)", "Kh" & ChrW(7889) & "i A3:E9 nghe Slicer Th" & ChrW(225) & "ng v" & ChrW(224) & " Mi" & ChrW(7873) & "n qua " & ChrW(273) & ChrW(7889) & "i s" & ChrW(7889) & " c" & ChrW(7911) & "a h" & ChrW(224) & "m. Kh" & ChrW(7889) & "i T1" & ChrW(8211) & "T4 theo Mi" & ChrW(7873) & "n " & ChrW(273) & "ang ch" & ChrW(7885) & "n. Kh" & ChrW(244) & "ng t" & ChrW(237) & "nh l" & ChrW(7841) & "i KPI " & ChrW(7903) & " " & ChrW(273) & ChrW(226) & "y, ch" & ChrW(7881) & " so s" & ChrW(225) & "nh v" & ChrW(7899) & "i target."
    ws.Range("A3:E3").Value = Array("KPI", "Gi" & ChrW(225) & " tr" & ChrW(7883), "Target", "M" & ChrW(7913) & "c " & ChrW(273) & ChrW(7841) & "t", "Tr" & ChrW(7841) & "ng th" & ChrW(225) & "i")
    Hdr ws.Range("A3:E3")
    kpi = Array("Doanh thu", "S" & ChrW(7889) & " " & ChrW(273) & ChrW(417) & "n", "AOV", "GM%", "YoY", "Kh" & ChrW(225) & "ch ho" & ChrW(7841) & "t " & ChrW(273) & ChrW(7897) & "ng")
    meas = Array("Total Revenue", "Total Orders", "Avg Order Value", "Gross Margin %", "Revenue YoY %", "Active Customers")
    fmt = Array("#,##0", "#,##0", "#,##0", "0.0%", "0.0%", "#,##0")
    For k = 0 To 5
        r = 4 + k
        ws.Cells(r, 1).Value = kpi(k)
        ws.Cells(r, 2).Formula = CVF(CStr(meas(k)), SL)
        If k = 0 Then
            ws.Cells(r, 3).Formula = CVF("Total Target", SL)
        Else
            ws.Cells(r, 3).Formula = "=Cai_dat!$C$" & (4 + k)
        End If
        ws.Cells(r, 4).Formula = "=IFERROR(B" & r & "/C" & r & ","""")"
        ws.Cells(r, 5).Formula = "=IF(D" & r & "="""","""",IF(D" & r & ">=Cai_dat!$C$13,""" & ChrW(10004) & " " & ChrW(272) & ChrW(7841) & "t"",IF(D" & r & ">=Cai_dat!$C$14,""" & ChrW(9650) & " C" & ChrW(7843) & "nh b" & ChrW(225) & "o"",""" & ChrW(10008) & " Kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(7841) & "t"")))"
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).NumberFormat = fmt(k)
        ws.Cells(r, 4).NumberFormat = "0.0%"
    Next
    Chk "kpi block"
    ws.Range("A11").Value = "Theo th" & ChrW(225) & "ng (Mi" & ChrW(7873) & "n " & ChrW(273) & "ang ch" & ChrW(7885) & "n): ngu" & ChrW(7891) & "n cho sparkline T1" & ChrW(8211) & "T4"
    ws.Range("A11").Font.Bold = True
    ws.Range("A12:I12").Value = Array("Th" & ChrW(225) & "ng", "Doanh thu", "Target", "M" & ChrW(7913) & "c " & ChrW(273) & ChrW(7841) & "t", "S" & ChrW(7889) & " " & ChrW(273) & ChrW(417) & "n", "AOV", "GM%", "YoY", "Kh" & ChrW(225) & "ch ho" & ChrW(7841) & "t " & ChrW(273) & ChrW(7897) & "ng")
    Hdr ws.Range("A12:I12")
    For m = 1 To 4
        r = 12 + m
        Dim a As String
        a = ",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[" & m & "]"",Slicer_Region"
        ws.Cells(r, 1).Value = m
        ws.Cells(r, 2).Formula = CVF("Total Revenue", a)
        ws.Cells(r, 3).Formula = CVF("Total Target", a)
        ws.Cells(r, 4).Formula = "=IFERROR(B" & r & "/C" & r & ","""")"
        ws.Cells(r, 5).Formula = CVF("Total Orders", a)
        ws.Cells(r, 6).Formula = CVF("Avg Order Value", a)
        ws.Cells(r, 7).Formula = CVF("Gross Margin %", a)
        ws.Cells(r, 8).Formula = CVF("Revenue YoY %", a)
        ws.Cells(r, 9).Formula = CVF("Active Customers", a)
    Next
    ws.Range("B13:C16").NumberFormat = "#,##0"
    ws.Range("D13:D16").NumberFormat = "0.0%"
    ws.Range("E13:F16").NumberFormat = "#,##0"
    ws.Range("G13:H16").NumberFormat = "0.0%"
    ws.Range("I13:I16").NumberFormat = "#,##0"
    Chk "month block"
    ws.Range("A19").Value = "Waterfall: ch" & ChrW(234) & "nh l" & ChrW(7879) & "ch doanh thu so v" & ChrW(7899) & "i target theo c" & ChrW(7917) & "a h" & ChrW(224) & "ng (Th" & ChrW(225) & "ng v" & ChrW(224) & " Mi" & ChrW(7873) & "n " & ChrW(273) & "ang ch" & ChrW(7885) & "n)"
    ws.Range("A19").Font.Bold = True
    ws.Range("A20:C20").Value = Array("M" & ChrW(227), "C" & ChrW(7917) & "a h" & ChrW(224) & "ng", "Ch" & ChrW(234) & "nh l" & ChrW(7879) & "ch vs target")
    Hdr ws.Range("A20:C20")
    Dim wsS As Worksheet
    Set wsS = wb.Worksheets("Store")
    For k = 1 To 12
        r = 20 + k
        ws.Cells(r, 1).Formula = "=Store!A" & (1 + k)
        ws.Cells(r, 2).Formula = "=Store!B" & (1 + k)
        ws.Cells(r, 3).Formula = "=IFERROR(CUBEVALUE(""" & CN & """,""[Measures].[Variance]"",""[Store].[Store_Code].&["" & A" & r & " & ""]"",""[Calendar].[Nam].&[2025]"",Slicer_Thang),0)"
    Next
    ws.Cells(33, 2).Value = "T" & ChrW(7893) & "ng"
    ws.Cells(33, 3).Formula = "=SUM(C21:C32)"
    ws.Range("A33:C33").Font.Bold = True
    ws.Range("C21:C33").NumberFormat = "#,##0;-#,##0"
    Chk "waterfall block"
    ws.Range("A36").Value = "Doanh thu theo th" & ChrW(224) & "nh ph" & ChrW(7889) & " (Th" & ChrW(225) & "ng v" & ChrW(224) & " Mi" & ChrW(7873) & "n " & ChrW(273) & "ang ch" & ChrW(7885) & "n)"
    ws.Range("A36").Font.Bold = True
    ws.Range("A37:B37").Value = Array("Th" & ChrW(224) & "nh ph" & ChrW(7889), "Doanh thu")
    Hdr ws.Range("A37:B37")
    Dim seen As String, city As String, n As Long
    seen = "|"
    n = 0
    For k = 2 To 13
        city = CStr(wsS.Cells(k, 3).Value)
        If InStr(seen, "|" & city & "|") = 0 Then
            seen = seen & city & "|"
            n = n + 1
            ws.Cells(37 + n, 1).Value = city
            ws.Cells(37 + n, 2).Formula = "=IFERROR(CUBEVALUE(""" & CN & """,""[Measures].[Total Revenue]"",""[Store].[City].&["" & A" & (37 + n) & " & ""]"",""[Calendar].[Nam].&[2025]"",Slicer_Thang,Slicer_Region),0)"
        End If
    Next
    ws.Range("B38:B50").NumberFormat = "#,##0"
    L "cities=" & n
    Chk "city block"
    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B:I").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 18
    ws.Range("E4:E9").Font.Name = "Segoe UI Symbol"
    L "stage4 ok"
End Sub

Private Sub CardBorder(rng As Range)
    Dim b As Variant
    rng.Interior.Color = RGB(255, 255, 255)
    For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        rng.Borders(b).LineStyle = xlContinuous
        rng.Borders(b).Color = RGB(200, 205, 215)
        rng.Borders(b).Weight = xlThin
    Next
End Sub

Private Sub Section(ws As Worksheet, ByVal r As Long, ByVal txt As String)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 19)).Merge
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Size = 11
    ws.Cells(r, 2).Font.Color = RGB(32, 56, 100)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 19)).Borders(xlEdgeBottom).Color = RGB(32, 56, 100)
End Sub

Private Function PlaceChart(ws As Worksheet, ByVal tl As String, ByVal br As String, ByVal ctype As Long, ByVal nm As String) As Shape
    Dim rg As Range
    Set rg = ws.Range(tl & ":" & br)
    Set PlaceChart = ws.Shapes.AddChart2(-1, ctype, rg.Left, rg.Top, rg.Width, rg.Height)
    PlaceChart.Name = nm
End Function

Private Sub ChartTitle(ch As Chart, ByVal txt As String)
    On Error Resume Next
    ch.HasTitle = True
    Chk "ct hastitle"
    ch.ChartTitle.Text = txt
    Chk "ct text"
    ch.ChartTitle.Font.Size = 11
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = RGB(32, 56, 100)
    Chk "ct font"
End Sub

Private Sub Stage5_Dashboard()
    Dim ws As Worksheet, wsP As Worksheet, wsD As Worksheet
    Dim k As Long, c0 As Long, r As Long
    Dim lab, fmtV, tgtF, shp As Shape, ch As Chart
    On Error Resume Next
    Set ws = GetSheet("Dashboard")
    Set wsP = GetSheet("PV_Dash")
    Set wsD = GetSheet("Data_Dash")
    ws.Cells.Interior.Color = RGB(243, 245, 249)
    ws.Range("A1:X60").Interior.Color = RGB(243, 245, 249)
    ws.Rows(1).RowHeight = 8
    ws.Range("B2:S2").Merge
    ws.Range("B2").Value = "MCI Mart " & ChrW(183) & " Dashboard b" & ChrW(225) & "n h" & ChrW(224) & "ng 2025"
    ws.Range("B2").Font.Size = 20
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(255, 255, 255)
    ws.Range("B2:S2").Interior.Color = RGB(32, 56, 100)
    ws.Rows(2).RowHeight = 34
    ws.Range("B3:S3").Merge
    ws.Range("B3").Formula = "=""Ng" & ChrW(432) & ChrW(7901) & "i xem: Gi" & ChrW(225) & "m " & ChrW(273) & ChrW(7889) & "c b" & ChrW(225) & "n h" & ChrW(224) & "ng " & ChrW(183) & " Th" & ChrW(225) & "ng " & ChrW(273) & "ang ch" & ChrW(7885) & "n: ""&IFERROR(CUBERANKEDMEMBER(""" & CN & """,Slicer_Thang,1),""?"")&""/2025 " & ChrW(183) & " D" & ChrW(7919) & " li" & ChrW(7879) & "u: Data Model (Sales, Target) " & ChrW(183) & " L" & ChrW(224) & "m m" & ChrW(7899) & "i h" & ChrW(7857) & "ng th" & ChrW(225) & "ng"""
    ws.Range("B3").Font.Size = 9
    ws.Range("B3").Font.Color = RGB(88, 88, 88)
    ws.Rows(4).RowHeight = 6
    Chk "title"
    ' --- 6 the KPI ---
    lab = Array("DOANH THU", "S" & ChrW(7888) & " " & ChrW(272) & ChrW(416) & "N", "GI" & ChrW(193) & " TR" & ChrW(7882) & " " & ChrW(272) & ChrW(416) & "N TB (AOV)", "BI" & ChrW(202) & "N L" & ChrW(7906) & "I NHU" & ChrW(7852) & "N G" & ChrW(7896) & "P", "T" & ChrW(258) & "NG TR" & ChrW(431) & ChrW(7902) & "NG SO C" & ChrW(217) & "NG K" & ChrW(7922), "KH" & ChrW(193) & "CH HO" & ChrW(7840) & "T " & ChrW(272) & ChrW(7896) & "NG")
    fmtV = Array("#,##0.0,"" tr""", "#,##0", "#,##0", "0.0%", "0.0%", "#,##0")
    tgtF = Array("=""Target ""&TEXT(Data_Dash!C4/1000000,""#,##0.0"")&"" tr""", "=""Target ""&TEXT(Data_Dash!C5,""#,##0"")", "=""Target ""&TEXT(Data_Dash!C6,""#,##0"")", "=""Target ""&TEXT(Data_Dash!C7,""0.0%"")", "=""Target ""&TEXT(Data_Dash!C8,""0.0%"")", "=""Target ""&TEXT(Data_Dash!C9,""#,##0"")")
    ws.Rows(5).RowHeight = 18
    ws.Rows(6).RowHeight = 22
    ws.Rows(7).RowHeight = 22
    ws.Rows(8).RowHeight = 16
    ws.Rows(9).RowHeight = 20
    ws.Rows(10).RowHeight = 18
    ws.Rows(11).RowHeight = 32
    For k = 0 To 5
        c0 = 2 + 3 * k
        r = 4 + k
        CardBorder ws.Range(ws.Cells(5, c0), ws.Cells(11, c0 + 2))
        ws.Range(ws.Cells(5, c0), ws.Cells(5, c0 + 2)).Merge
        ws.Cells(5, c0).Value = lab(k)
        ws.Cells(5, c0).Font.Size = 8
        ws.Cells(5, c0).Font.Bold = True
        ws.Cells(5, c0).Font.Color = RGB(88, 88, 88)
        ws.Range(ws.Cells(6, c0), ws.Cells(7, c0 + 2)).Merge
        ws.Cells(6, c0).Formula = "=Data_Dash!B" & r
        ws.Cells(6, c0).NumberFormat = fmtV(k)
        ws.Cells(6, c0).Font.Size = 22
        ws.Cells(6, c0).Font.Bold = True
        ws.Cells(6, c0).Font.Color = RGB(32, 56, 100)
        ws.Range(ws.Cells(8, c0), ws.Cells(8, c0 + 2)).Merge
        ws.Cells(8, c0).Formula = tgtF(k)
        ws.Cells(8, c0).Font.Size = 9
        ws.Cells(8, c0).Font.Color = RGB(88, 88, 88)
        ws.Range(ws.Cells(9, c0), ws.Cells(9, c0 + 2)).Merge
        ws.Cells(9, c0).Formula = "=Data_Dash!D" & r
        ws.Cells(9, c0).NumberFormat = """M" & ChrW(7913) & "c " & ChrW(273) & ChrW(7841) & "t ""0.0%"
        ws.Cells(9, c0).Font.Size = 10
        ws.Cells(9, c0).Font.Bold = True
        ws.Range(ws.Cells(10, c0), ws.Cells(10, c0 + 2)).Merge
        ws.Cells(10, c0).Formula = "=Data_Dash!E" & r
        ws.Cells(10, c0).Font.Name = "Segoe UI Symbol"
        ws.Cells(10, c0).Font.Size = 10
        ws.Cells(10, c0).Font.Bold = True
        ws.Range(ws.Cells(5, c0), ws.Cells(11, c0 + 2)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(11, c0), ws.Cells(11, c0 + 2)).Merge
        Chk "card " & k
        ' RAG: icon set (Type = Formula tro Cai_dat) + chu trang thai
        Dim fc As Object, rg As Range
        Set rg = ws.Cells(9, c0)
        rg.FormatConditions.Delete
        Set fc = rg.FormatConditions.AddIconSetCondition
        fc.IconSet = wb.IconSets(xl3Symbols)
        fc.ShowIconOnly = False
        fc.IconCriteria(2).Type = xlConditionValueFormula
        fc.IconCriteria(2).Value = "=Cai_dat!$C$14"
        fc.IconCriteria(2).Operator = xlGreaterEqual
        fc.IconCriteria(3).Type = xlConditionValueFormula
        fc.IconCriteria(3).Value = "=Cai_dat!$C$13"
        fc.IconCriteria(3).Operator = xlGreaterEqual
        Chk "icon set " & k
        Set rg = ws.Cells(10, c0)
        rg.FormatConditions.Delete
        Set fc = rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=Data_Dash!$D$" & r & ">=Cai_dat!$C$13")
        fc.Font.Color = RGB(12, 163, 12)
        Set fc = rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(Data_Dash!$D$" & r & "<Cai_dat!$C$13,Data_Dash!$D$" & r & ">=Cai_dat!$C$14)")
        fc.Font.Color = RGB(190, 125, 0)
        Set fc = rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(Data_Dash!$D$" & r & "),Data_Dash!$D$" & r & "<Cai_dat!$C$14)")
        fc.Font.Color = RGB(208, 59, 59)
        Chk "status cf " & k
    Next
    ' --- sparkline T1-T4 (Data_Dash B13:B16 ...) ---
    Dim srcCol: srcCol = Array("B", "E", "F", "G", "H", "I")
    For k = 0 To 5
        c0 = 2 + 3 * k
        ws.Cells(11, c0).SparklineGroups.Add Type:=xlSparkLine, SourceData:="Data_Dash!" & srcCol(k) & "13:" & srcCol(k) & "16"
        Chk "sparkline " & k
        With ws.Cells(11, c0).SparklineGroups(1)
            .SeriesColor.Color = RGB(32, 56, 100)
            .Points.Markers.Visible = True
            .Points.Lowpoint.Visible = False
        End With
        Chk "sparkline fmt " & k
    Next
    ws.Rows(12).RowHeight = 8
    Section ws, 13, "XU H" & ChrW(431) & ChrW(7898) & "NG"
    Section ws, 31, "CHI TI" & ChrW(7870) & "T"
    ' --- hang Xu huong ---
    Set shp = PlaceChart(ws, "B14", "J29", xlLine, "chtDoanhThuNgay")
    Set ch = shp.Chart
    ch.SetSourceData Source:=wsP.Range("A8").CurrentRegion
    Chk "line setsource"
    Dim pt As PivotTable
    Set pt = wsP.PivotTables("pvtNgay_Dash")
    ch.SetSourceData Source:=pt.TableRange1
    Chk "line setsource pt"
    ChartTitle ch, "Doanh thu theo ng" & ChrW(224) & "y (theo Slicer Th" & ChrW(225) & "ng v" & ChrW(224) & " Mi" & ChrW(7873) & "n)"
    ch.HasLegend = False
    ch.ShowAllFieldButtons = False
    shp.AlternativeText = "Bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(7891) & " " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng: doanh thu t" & ChrW(7915) & "ng ng" & ChrW(224) & "y trong th" & ChrW(225) & "ng " & ChrW(273) & "ang ch" & ChrW(7885) & "n."
    Chk "line fmt"
    Set shp = PlaceChart(ws, "K14", "S29", xlColumnClustered, "chtCombo")
    Set ch = shp.Chart
    Set pt = wsP.PivotTables("pvtCombo_Dash")
    ch.SetSourceData Source:=pt.TableRange1
    Chk "combo setsource"
    ch.SeriesCollection(2).ChartType = xlLineMarkers
    Chk "combo type"
    ChartTitle ch, "Th" & ChrW(7921) & "c t" & ChrW(7871) & " v" & ChrW(224) & " target theo th" & ChrW(225) & "ng " & ChrW(183) & " c" & ChrW(7843) & " 4 th" & ChrW(225) & "ng, kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(7893) & "i theo Slicer Th" & ChrW(225) & "ng"
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.ShowAllFieldButtons = False
    shp.AlternativeText = "Bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(7891) & " c" & ChrW(7897) & "t v" & ChrW(224) & " " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng: doanh thu th" & ChrW(7921) & "c t" & ChrW(7871) & " so v" & ChrW(7899) & "i target, theo th" & ChrW(225) & "ng 2025. Ch" & ChrW(7881) & " " & ChrW(273) & ChrW(7893) & "i theo Slicer Mi" & ChrW(7873) & "n."
    Chk "combo fmt"
    ' --- hang Chi tiet ---
    wsD.Activate
    wsD.Range("B20:C33").Select
    Set shp = wsD.Shapes.AddChart2(-1, XL_WATERFALL, 700, 20, 420, 260)
    Chk "waterfall add"
    shp.Name = "chtWaterfall"
    Set ch = shp.Chart
    ch.HasTitle = True
    Chk "waterfall hastitle"
    ch.ChartTitle.Text = "Ch" & ChrW(234) & "nh l" & ChrW(7879) & "ch so v" & ChrW(7899) & "i target theo c" & ChrW(7917) & "a h" & ChrW(224) & "ng"
    Chk "waterfall title text"
    ch.HasLegend = False
    Chk "waterfall legend"
    ch.FullSeriesCollection(1).Points(13).IsTotal = True
    Chk "waterfall total"
    shp.AlternativeText = "Bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(7891) & " waterfall: ch" & ChrW(234) & "nh l" & ChrW(7879) & "ch doanh thu so v" & ChrW(7899) & "i target c" & ChrW(7911) & "a t" & ChrW(7915) & "ng c" & ChrW(7917) & "a h" & ChrW(224) & "ng, c" & ChrW(7897) & "ng d" & ChrW(7891) & "n th" & ChrW(224) & "nh t" & ChrW(7893) & "ng."
    Chk "waterfall alt"
    shp.Cut
    Chk "waterfall cut"
    ws.Activate
    ws.Range("B32").Select
    ws.Paste
    Chk "waterfall paste"
    Set shp = ws.Shapes(ws.Shapes.Count)
    shp.Left = ws.Range("B32").Left
    shp.Top = ws.Range("B32").Top
    shp.Width = ws.Range("B32:G32").Width
    shp.Height = ws.Range("B32:B49").Height
    Chk "waterfall place " & shp.Name
    Set shp = PlaceChart(ws, "H32", "M49", xlBarClustered, "chtTop5")
    Set ch = shp.Chart
    Set pt = wsP.PivotTables("pvtTop5_Dash")
    ch.SetSourceData Source:=pt.TableRange1
    Chk "top5 setsource"
    ch.Axes(xlCategory).ReversePlotOrder = True
    ChartTitle ch, "Top 5 s" & ChrW(7843) & "n ph" & ChrW(7849) & "m theo doanh thu"
    ch.HasLegend = False
    ch.ShowAllFieldButtons = False
    shp.AlternativeText = "Bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(7891) & " thanh ngang: 5 s" & ChrW(7843) & "n ph" & ChrW(7849) & "m c" & ChrW(243) & " doanh thu cao nh" & ChrW(7845) & "t trong th" & ChrW(225) & "ng " & ChrW(273) & "ang ch" & ChrW(7885) & "n."
    Chk "top5 fmt"
    Set shp = PlaceChart(ws, "N32", "S49", XL_REGIONMAP, "chtMap")
    Set ch = shp.Chart
    ch.SetSourceData Source:=wsD.Range("A37:B46")
    Chk "map setsource"
    ChartTitle ch, "Doanh thu theo th" & ChrW(224) & "nh ph" & ChrW(7889)
    Chk "map title"
    shp.AlternativeText = "B" & ChrW(7843) & "n " & ChrW(273) & ChrW(7891) & " doanh thu theo th" & ChrW(224) & "nh ph" & ChrW(7889) & " trong th" & ChrW(225) & "ng " & ChrW(273) & "ang ch" & ChrW(7885) & "n."
    Chk "map alt"
    ws.Range("B51").Value = "Quy t" & ChrW(7855) & "c RAG: " & ChrW(10004) & " " & ChrW(272) & ChrW(7841) & "t khi m" & ChrW(7913) & "c " & ChrW(273) & ChrW(7841) & "t " & ChrW(8805) & " 100% " & ChrW(183) & " " & ChrW(9650) & " C" & ChrW(7843) & "nh b" & ChrW(225) & "o t" & ChrW(7915) & " 80% " & ChrW(273) & ChrW(7871) & "n d" & ChrW(432) & ChrW(7899) & "i 100% " & ChrW(183) & " " & ChrW(10008) & " Kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(7841) & "t khi d" & ChrW(432) & ChrW(7899) & "i 80% (YoY: m" & ChrW(7913) & "c " & ChrW(273) & ChrW(7841) & "t = YoY " & ChrW(247) & " 10%). Ng" & ChrW(432) & ChrW(7905) & "ng n" & ChrW(7857) & "m " & ChrW(7903) & " sheet Cai_dat."
    ws.Range("B52").Value = "Slicer Th" & ChrW(225) & "ng v" & ChrW(224) & " Mi" & ChrW(7873) & "n " & ChrW(273) & "i" & ChrW(7873) & "u khi" & ChrW(7875) & "n th" & ChrW(7867) & " KPI, " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng theo ng" & ChrW(224) & "y, waterfall, Top 5 v" & ChrW(224) & " th" & ChrW(224) & "nh ph" & ChrW(7889) & ". Combo Th" & ChrW(7921) & "c t" & ChrW(7871) & ChrW(8211) & "Target lu" & ChrW(244) & "n hi" & ChrW(7879) & "n c" & ChrW(7843) & " 4 th" & ChrW(225) & "ng, ch" & ChrW(7881) & " " & ChrW(273) & ChrW(7893) & "i theo Mi" & ChrW(7873) & "n."
    ws.Range("B51:B52").Font.Size = 8
    ws.Range("B51:B52").Font.Color = RGB(88, 88, 88)
    L "stage5 ok"
End Sub

Private Sub Stage6_Brief()
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = GetSheet("Brief")
    Title ws, "Brief dashboard MCI Mart (" & ChrW(273) & ChrW(225) & "p " & ChrW(225) & "n m" & ChrW(7851) & "u)", "B" & ChrW(7889) & "n m" & ChrW(7909) & "c c" & ChrW(7911) & "a brief. M" & ChrW(7895) & "i visual tr" & ChrW(234) & "n Dashboard ph" & ChrW(7843) & "i tr" & ChrW(7843) & " l" & ChrW(7901) & "i " & ChrW(273) & ChrW(250) & "ng m" & ChrW(7897) & "t c" & ChrW(226) & "u h" & ChrW(7887) & "i b" & ChrW(234) & "n d" & ChrW(432) & ChrW(7899) & "i."
    ws.Range("A4:C4").Value = Array("M" & ChrW(7909) & "c", "C" & ChrW(226) & "u h" & ChrW(7887) & "i", "Dashboard MCI Mart")
    Hdr ws.Range("A4:C4")
    ws.Range("A5:C5").Value = Array("Ng" & ChrW(432) & ChrW(7901) & "i xem", "Executive, operational hay self-service?", "Sales Director: executive, xem trong h" & ChrW(7885) & "p ban " & ChrW(273) & "i" & ChrW(7873) & "u h" & ChrW(224) & "nh.")
    ws.Range("A6:C6").Value = Array("C" & ChrW(226) & "u h" & ChrW(7887) & "i (3" & ChrW(8211) & "5, " & ChrW(273) & "o l" & ChrW(432) & ChrW(7901) & "ng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c)", "Ch" & ChrW(7881) & " s" & ChrW(7889) & " " & ChrW(183) & " chi" & ChrW(7873) & "u " & ChrW(183) & " th" & ChrW(7901) & "i gian " & ChrW(183) & " m" & ChrW(7889) & "c so s" & ChrW(225) & "nh", "Q1. Doanh thu th" & ChrW(225) & "ng n" & ChrW(224) & "y " & ChrW(273) & ChrW(7841) & "t bao nhi" & ChrW(234) & "u % target, so v" & ChrW(7899) & "i th" & ChrW(225) & "ng tr" & ChrW(432) & ChrW(7899) & "c v" & ChrW(224) & " c" & ChrW(249) & "ng k" & ChrW(7923) & "?")
    ws.Range("A7:C7").Value = Array("", "", "Q2. C" & ChrW(7917) & "a h" & ChrW(224) & "ng n" & ChrW(224) & "o d" & ChrW(432) & ChrW(7899) & "i 80% target?")
    ws.Range("A8:C8").Value = Array("", "", "Q3. Ng" & ChrW(224) & "y n" & ChrW(224) & "o trong th" & ChrW(225) & "ng b" & ChrW(7845) & "t th" & ChrW(432) & ChrW(7901) & "ng?")
    ws.Range("A9:C9").Value = Array("", "", "Q4. 5 s" & ChrW(7843) & "n ph" & ChrW(7849) & "m " & ChrW(273) & ChrW(7847) & "u chi" & ChrW(7871) & "m bao nhi" & ChrW(234) & "u % doanh thu?")
    ws.Range("A10:C10").Value = Array("", "", "Q5. Th" & ChrW(224) & "nh ph" & ChrW(7889) & " n" & ChrW(224) & "o " & ChrW(273) & ChrW(243) & "ng g" & ChrW(243) & "p nhi" & ChrW(7873) & "u nh" & ChrW(7845) & "t?")
    ws.Range("A11:C11").Value = Array("T" & ChrW(7847) & "n su" & ChrW(7845) & "t refresh", "Bao l" & ChrW(226) & "u m" & ChrW(7897) & "t l" & ChrW(7847) & "n?", "H" & ChrW(7857) & "ng th" & ChrW(225) & "ng, khi c" & ChrW(243) & " file th" & ChrW(225) & "ng m" & ChrW(7899) & "i.")
    ws.Range("A12:C12").Value = Array("Quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh", "Dashboard gi" & ChrW(250) & "p quy" & ChrW(7871) & "t " & ChrW(273) & ChrW(7883) & "nh g" & ChrW(236) & "?", "H" & ChrW(7895) & " tr" & ChrW(7907) & " c" & ChrW(7917) & "a h" & ChrW(224) & "ng n" & ChrW(224) & "o tr" & ChrW(432) & ChrW(7899) & "c; " & ChrW(273) & "i" & ChrW(7873) & "u ch" & ChrW(7881) & "nh target.")
    ws.Range("A14").Value = "Visual n" & ChrW(224) & "o tr" & ChrW(7843) & " l" & ChrW(7901) & "i c" & ChrW(226) & "u h" & ChrW(7887) & "i n" & ChrW(224) & "o"
    ws.Range("A14").Font.Bold = True
    ws.Range("A15:C15").Value = Array("Visual", "C" & ChrW(226) & "u h" & ChrW(7887) & "i", "Ghi ch" & ChrW(250))
    Hdr ws.Range("A15:C15")
    ws.Range("A16:C16").Value = Array("6 th" & ChrW(7867) & " KPI + combo Th" & ChrW(7921) & "c t" & ChrW(7871) & ChrW(8211) & "Target", "Q1", "M" & ChrW(7913) & "c " & ChrW(273) & ChrW(7841) & "t k" & ChrW(232) & "m bi" & ChrW(7875) & "u t" & ChrW(432) & ChrW(7907) & "ng " & ChrW(10004) & " " & ChrW(9650) & " " & ChrW(10008) & "; combo cho m" & ChrW(7889) & "c so s" & ChrW(225) & "nh gi" & ChrW(7919) & "a c" & ChrW(225) & "c th" & ChrW(225) & "ng")
    ws.Range("A17:C17").Value = Array("Waterfall theo c" & ChrW(7917) & "a h" & ChrW(224) & "ng", "Q2", "C" & ChrW(7917) & "a h" & ChrW(224) & "ng d" & ChrW(432) & ChrW(7899) & "i 80% target nh" & ChrW(236) & "n th" & ChrW(7845) & "y ngay " & ChrW(7903) & " c" & ChrW(225) & "c c" & ChrW(7897) & "t " & ChrW(226) & "m l" & ChrW(7899) & "n")
    ws.Range("A18:C18").Value = Array(ChrW(272) & ChrW(432) & ChrW(7901) & "ng doanh thu theo ng" & ChrW(224) & "y", "Q3", "Ng" & ChrW(224) & "y v" & ChrW(432) & ChrW(7907) & "t kh" & ChrW(7887) & "i d" & ChrW(7843) & "i th" & ChrW(432) & ChrW(7901) & "ng l" & ChrW(7879))
    ws.Range("A19:C19").Value = Array("Top 5 s" & ChrW(7843) & "n ph" & ChrW(7849) & "m", "Q4", "T" & ChrW(7893) & "ng 5 s" & ChrW(7843) & "n ph" & ChrW(7849) & "m chia cho doanh thu th" & ChrW(225) & "ng")
    ws.Range("A20:C20").Value = Array("Doanh thu theo th" & ChrW(224) & "nh ph" & ChrW(7889), "Q5", "B" & ChrW(7843) & "n " & ChrW(273) & ChrW(7891) & "; n" & ChrW(7871) & "u kh" & ChrW(244) & "ng c" & ChrW(243) & " Internet th" & ChrW(236) & " d" & ChrW(249) & "ng bar theo th" & ChrW(224) & "nh ph" & ChrW(7889))
    ws.Range("A22").Value = "C" & ChrW(226) & "u h" & ChrW(7887) & "i y" & ChrW(7871) & "u: ""Xem doanh thu"", ""Ph" & ChrW(226) & "n t" & ChrW(237) & "ch kh" & ChrW(225) & "ch h" & ChrW(224) & "ng"" (ch" & ChrW(7911) & " " & ChrW(273) & ChrW(7873) & ", thi" & ChrW(7871) & "u ch" & ChrW(7881) & " s" & ChrW(7889) & ", chi" & ChrW(7873) & "u, th" & ChrW(7901) & "i gian ho" & ChrW(7863) & "c m" & ChrW(7889) & "c so s" & ChrW(225) & "nh)."
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 78
    ws.Range("A5:C20").WrapText = True
    ws.Range("A5:C20").VerticalAlignment = xlTop
    ws.Range("A5:A12").Font.Bold = True
    Chk "brief"
    L "stage6 ok"
End Sub

Private Sub Stage7_DuBao()
    Dim ws As Worksheet, wsS As Worksheet, colD As Variant, colR As Variant, r As Long
    Dim sc As Object, shp As Shape, ch As Chart
    On Error Resume Next
    Set ws = GetSheet("Du_bao")
    Set wsS = wb.Worksheets("Sales_PQ")
    Title ws, "D" & ChrW(7921) & " b" & ChrW(225) & "o doanh thu: Trendline v" & ChrW(224) & " ETS tr" & ChrW(234) & "n b" & ChrW(7843) & "ng Ng" & ChrW(224) & "y" & ChrW(8211) & "Doanh thu", "486 ng" & ChrW(224) & "y li" & ChrW(234) & "n t" & ChrW(7909) & "c 01/01/2024" & ChrW(8211) & "30/04/2025 (SUMIFS t" & ChrW(7915) & " Sales_PQ). D" & ChrW(7921) & " b" & ChrW(225) & "o th" & ChrW(225) & "ng 5 b" & ChrW(7857) & "ng FORECAST.ETS v" & ChrW(7899) & "i chu k" & ChrW(7923) & " " & ChrW(273) & ChrW(7863) & "t tay = 7."
    colD = Application.Match("Order_Date", wsS.Rows(1), 0)
    colR = Application.Match("Revenue", wsS.Rows(1), 0)
    L "Sales_PQ cols: Order_Date=" & colD & " Revenue=" & colR
    ws.Range("A3:B3").Value = Array("Ng" & ChrW(224) & "y", "Doanh thu")
    Hdr ws.Range("A3:B3")
    ws.Range("A4").Formula = "=DATE(2024,1,1)"
    ws.Range("A5:A489").Formula = "=A4+1"
    ws.Range("A4:A489").NumberFormat = "dd/mm/yyyy"
    Dim sD As String, sR As String
    sD = Split(wsS.Cells(1, colD).Address(True, False), "$")(0)
    sR = Split(wsS.Cells(1, colR).Address(True, False), "$")(0)
    ws.Range("B4:B489").Formula = "=SUMIFS(Sales_PQ!$" & sR & ":$" & sR & ",Sales_PQ!$" & sD & ":$" & sD & ",A4)"
    ws.Range("B4:B489").NumberFormat = "#,##0"
    Chk "du_bao data"
    L "Du_bao B4=" & ws.Range("B4").Value & " tong=" & Application.Sum(ws.Range("B4:B489"))
    ws.Range("D3:G3").Value = Array("Ng" & ChrW(224) & "y d" & ChrW(7921) & " b" & ChrW(225) & "o", "D" & ChrW(7921) & " b" & ChrW(225) & "o (ETS)", "C" & ChrW(7853) & "n d" & ChrW(432) & ChrW(7899) & "i 95%", "C" & ChrW(7853) & "n tr" & ChrW(234) & "n 95%")
    Hdr ws.Range("D3:G3")
    ws.Range("D4").Formula = "=DATE(2025,5,1)"
    ws.Range("D5:D33").Formula = "=D4+1"
    ws.Range("D4:D33").NumberFormat = "dd/mm/yyyy"
    ws.Range("E4:E33").Formula = "=_xlfn.FORECAST.ETS(D4,$B$4:$B$489,$A$4:$A$489,7,1,1)"
    ws.Range("F4:F33").Formula = "=E4-_xlfn.FORECAST.ETS.CONFINT(D4,$B$4:$B$489,$A$4:$A$489,0.95,7,1,1)"
    ws.Range("G4:G33").Formula = "=E4+_xlfn.FORECAST.ETS.CONFINT(D4,$B$4:$B$489,$A$4:$A$489,0.95,7,1,1)"
    ws.Range("E4:G33").NumberFormat = "#,##0"
    Chk "du_bao forecast"
    ws.Range("D35").Value = "T" & ChrW(7893) & "ng d" & ChrW(7921) & " b" & ChrW(225) & "o th" & ChrW(225) & "ng 5"
    ws.Range("E35").Formula = "=SUM(E4:E33)"
    ws.Range("E35").NumberFormat = "#,##0"
    ws.Range("D36").Value = "Chu k" & ChrW(7923) & " t" & ChrW(7921) & " ph" & ChrW(225) & "t hi" & ChrW(7879) & "n (Detect)"
    ws.Range("E36").Formula = "=_xlfn.FORECAST.ETS.SEASONALITY($B$4:$B$489,$A$4:$A$489,1,1)"
    ws.Range("D35:D36").Font.Bold = True
    ws.Range("I3:L3").Value = Array("Ng" & ChrW(224) & "y (th" & ChrW(225) & "ng 4)", "Th" & ChrW(7921) & "c t" & ChrW(7871), "D" & ChrW(7921) & " b" & ChrW(225) & "o t" & ChrW(7915) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u " & ChrW(273) & ChrW(7871) & "n 31/03", "Ch" & ChrW(234) & "nh l" & ChrW(7879) & "ch")
    Hdr ws.Range("I3:L3")
    ws.Range("I4").Formula = "=DATE(2025,4,1)"
    ws.Range("I5:I33").Formula = "=I4+1"
    ws.Range("I4:I33").NumberFormat = "dd/mm/yyyy"
    ws.Range("J4:J33").Formula = "=SUMIFS($B$4:$B$489,$A$4:$A$489,I4)"
    ws.Range("K4:K33").Formula = "=_xlfn.FORECAST.ETS(I4,$B$4:$B$459,$A$4:$A$459,7,1,1)"
    ws.Range("L4:L33").Formula = "=K4-J4"
    ws.Range("J4:L33").NumberFormat = "#,##0"
    ws.Range("I35").Value = "T" & ChrW(7893) & "ng"
    ws.Range("J35").Formula = "=SUM(J4:J33)"
    ws.Range("K35").Formula = "=SUM(K4:K33)"
    ws.Range("L35").Formula = "=K35-J35"
    ws.Range("I36").Value = "L" & ChrW(7879) & "ch so v" & ChrW(7899) & "i th" & ChrW(7921) & "c t" & ChrW(7871)
    ws.Range("J36").Formula = "=K35/J35-1"
    ws.Range("J35:L35").NumberFormat = "#,##0"
    ws.Range("J36").NumberFormat = "0.0%"
    ws.Range("I35:I36").Font.Bold = True
    Chk "du_bao backtest"
    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 28
    ws.Columns("E:G").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 22
    ws.Columns("J:L").ColumnWidth = 30
    Set shp = ws.Shapes.AddChart2(-1, xlXYScatterLinesNoMarkers, ws.Range("N3").Left, ws.Range("N3").Top, 620, 330)
    shp.Name = "chtDuBao"
    Set ch = shp.Chart
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    Set sc = ch.SeriesCollection.NewSeries
    sc.Name = "Th" & ChrW(7921) & "c t" & ChrW(7871)
    sc.XValues = ws.Range("A4:A489")
    sc.Values = ws.Range("B4:B489")
    Set sc = ch.SeriesCollection.NewSeries
    sc.Name = "D" & ChrW(7921) & " b" & ChrW(225) & "o th" & ChrW(225) & "ng 5"
    sc.XValues = ws.Range("D4:D33")
    sc.Values = ws.Range("E4:E33")
    Set sc = ch.SeriesCollection.NewSeries
    sc.Name = "C" & ChrW(7853) & "n d" & ChrW(432) & ChrW(7899) & "i 95%"
    sc.XValues = ws.Range("D4:D33")
    sc.Values = ws.Range("F4:F33")
    Set sc = ch.SeriesCollection.NewSeries
    sc.Name = "C" & ChrW(7853) & "n tr" & ChrW(234) & "n 95%"
    sc.XValues = ws.Range("D4:D33")
    sc.Values = ws.Range("G4:G33")
    ChartTitle ch, "Doanh thu theo ng" & ChrW(224) & "y v" & ChrW(224) & " d" & ChrW(7921) & " b" & ChrW(225) & "o ETS th" & ChrW(225) & "ng 5/2025"
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Axes(xlCategory).TickLabels.NumberFormat = "mm/yyyy"
    shp.AlternativeText = "Bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(7891) & " " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng: doanh thu t" & ChrW(7915) & "ng ng" & ChrW(224) & "y 2024 " & ChrW(273) & ChrW(7871) & "n th" & ChrW(225) & "ng 4/2025 v" & ChrW(224) & " d" & ChrW(7921) & " b" & ChrW(225) & "o ETS th" & ChrW(225) & "ng 5/2025 k" & ChrW(232) & "m d" & ChrW(7843) & "i tin c" & ChrW(7853) & "y 95%."
    Chk "du_bao chart"
    L "stage7 ok"
End Sub

Private Sub Stage8_KiemTra()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetSheet("Kiem_tra_B6")
    Title ws, "KI" & ChrW(7874) & "M TRA / " & ChrW(272) & ChrW(7888) & "I SO" & ChrW(193) & "T " & ChrW(8212) & " Brick output B6 (Dashboard)", "C" & ChrW(7897) & "t C: CUBEVALUE v" & ChrW(7899) & "i th" & ChrW(225) & "ng vi" & ChrW(7871) & "t th" & ChrW(7859) & "ng trong c" & ChrW(244) & "ng th" & ChrW(7913) & "c (kh" & ChrW(244) & "ng ph" & ChrW(7909) & " thu" & ChrW(7897) & "c Slicer). C" & ChrW(7897) & "t D: " & ChrW(273) & ChrW(225) & "p " & ChrW(225) & "n t" & ChrW(237) & "nh t" & ChrW(7915) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u (build_b6_meta.json)."
    ws.Range("A4:E4").Value = Array("Ch" & ChrW(7881) & " ti" & ChrW(234) & "u", "", "K" & ChrW(7871) & "t qu" & ChrW(7843), ChrW(272) & ChrW(225) & "p " & ChrW(225) & "n", "Kh" & ChrW(7899) & "p?")
    Hdr ws.Range("A4:E4")
    ws.Cells(5, 1).Value = "Doanh thu th" & ChrW(225) & "ng 2/2025"
    ws.Cells(5, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[2]"")"
    ws.Cells(5, 4).Value = 245996000.0
    ws.Cells(5, 5).Formula = "=IF(ISNUMBER(C5),IF(ABS(C5-D5)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C5:D5").NumberFormat = "#,##0"
    ws.Cells(6, 1).Value = "S" & ChrW(7889) & " " & ChrW(273) & ChrW(417) & "n th" & ChrW(225) & "ng 2/2025"
    ws.Cells(6, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Orders]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[2]"")"
    ws.Cells(6, 4).Value = 823.0
    ws.Cells(6, 5).Formula = "=IF(ISNUMBER(C6),IF(ABS(C6-D6)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C6:D6").NumberFormat = "#,##0"
    ws.Cells(7, 1).Value = "AOV th" & ChrW(225) & "ng 2/2025"
    ws.Cells(7, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Avg Order Value]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[2]"")"
    ws.Cells(7, 4).Value = 298901.57958687725
    ws.Cells(7, 5).Formula = "=IF(ISNUMBER(C7),IF(ABS(C7-D7)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C7:D7").NumberFormat = "#,##0"
    ws.Cells(8, 1).Value = "GM% th" & ChrW(225) & "ng 2/2025"
    ws.Cells(8, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Gross Margin %]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[2]"")"
    ws.Cells(8, 4).Value = 0.33702783785102197
    ws.Cells(8, 5).Formula = "=IF(ISNUMBER(C8),IF(ABS(C8-D8)<=0.0005,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C8:D8").NumberFormat = "0.0%"
    ws.Cells(9, 1).Value = "YoY th" & ChrW(225) & "ng 2/2025"
    ws.Cells(9, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Revenue YoY %]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[2]"")"
    ws.Cells(9, 4).Value = -0.009295057671241704
    ws.Cells(9, 5).Formula = "=IF(ISNUMBER(C9),IF(ABS(C9-D9)<=0.0005,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C9:D9").NumberFormat = "0.0%"
    ws.Cells(10, 1).Value = "Kh" & ChrW(225) & "ch ho" & ChrW(7841) & "t " & ChrW(273) & ChrW(7897) & "ng th" & ChrW(225) & "ng 2/2025"
    ws.Cells(10, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Active Customers]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[2]"")"
    ws.Cells(10, 4).Value = 605.0
    ws.Cells(10, 5).Formula = "=IF(ISNUMBER(C10),IF(ABS(C10-D10)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C10:D10").NumberFormat = "#,##0"
    ws.Cells(11, 1).Value = "Target th" & ChrW(225) & "ng 2/2025"
    ws.Cells(11, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Target]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[2]"")"
    ws.Cells(11, 4).Value = 298000000.0
    ws.Cells(11, 5).Formula = "=IF(ISNUMBER(C11),IF(ABS(C11-D11)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C11:D11").NumberFormat = "#,##0"
    ws.Cells(12, 1).Value = "Doanh thu th" & ChrW(225) & "ng 3/2025"
    ws.Cells(12, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(12, 4).Value = 269608000.0
    ws.Cells(12, 5).Formula = "=IF(ISNUMBER(C12),IF(ABS(C12-D12)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C12:D12").NumberFormat = "#,##0"
    ws.Cells(13, 1).Value = "S" & ChrW(7889) & " " & ChrW(273) & ChrW(417) & "n th" & ChrW(225) & "ng 3/2025"
    ws.Cells(13, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Orders]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(13, 4).Value = 923.0
    ws.Cells(13, 5).Formula = "=IF(ISNUMBER(C13),IF(ABS(C13-D13)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C13:D13").NumberFormat = "#,##0"
    ws.Cells(14, 1).Value = "AOV th" & ChrW(225) & "ng 3/2025"
    ws.Cells(14, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Avg Order Value]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(14, 4).Value = 292099.6749729144
    ws.Cells(14, 5).Formula = "=IF(ISNUMBER(C14),IF(ABS(C14-D14)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C14:D14").NumberFormat = "#,##0"
    ws.Cells(15, 1).Value = "GM% th" & ChrW(225) & "ng 3/2025"
    ws.Cells(15, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Gross Margin %]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(15, 4).Value = 0.33073017121153675
    ws.Cells(15, 5).Formula = "=IF(ISNUMBER(C15),IF(ABS(C15-D15)<=0.0005,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C15:D15").NumberFormat = "0.0%"
    ws.Cells(16, 1).Value = "YoY th" & ChrW(225) & "ng 3/2025"
    ws.Cells(16, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Revenue YoY %]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(16, 4).Value = 0.1366656548281561
    ws.Cells(16, 5).Formula = "=IF(ISNUMBER(C16),IF(ABS(C16-D16)<=0.0005,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C16:D16").NumberFormat = "0.0%"
    ws.Cells(17, 1).Value = "Kh" & ChrW(225) & "ch ho" & ChrW(7841) & "t " & ChrW(273) & ChrW(7897) & "ng th" & ChrW(225) & "ng 3/2025"
    ws.Cells(17, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Active Customers]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(17, 4).Value = 637.0
    ws.Cells(17, 5).Formula = "=IF(ISNUMBER(C17),IF(ABS(C17-D17)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C17:D17").NumberFormat = "#,##0"
    ws.Cells(18, 1).Value = "Target th" & ChrW(225) & "ng 3/2025"
    ws.Cells(18, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Target]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(18, 4).Value = 285000000.0
    ws.Cells(18, 5).Formula = "=IF(ISNUMBER(C18),IF(ABS(C18-D18)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C18:D18").NumberFormat = "#,##0"
    ws.Cells(19, 1).Value = "Doanh thu th" & ChrW(225) & "ng 4/2025"
    ws.Cells(19, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"")"
    ws.Cells(19, 4).Value = 296965000.0
    ws.Cells(19, 5).Formula = "=IF(ISNUMBER(C19),IF(ABS(C19-D19)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C19:D19").NumberFormat = "#,##0"
    ws.Cells(20, 1).Value = "S" & ChrW(7889) & " " & ChrW(273) & ChrW(417) & "n th" & ChrW(225) & "ng 4/2025"
    ws.Cells(20, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Orders]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"")"
    ws.Cells(20, 4).Value = 966.0
    ws.Cells(20, 5).Formula = "=IF(ISNUMBER(C20),IF(ABS(C20-D20)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C20:D20").NumberFormat = "#,##0"
    ws.Cells(21, 1).Value = "AOV th" & ChrW(225) & "ng 4/2025"
    ws.Cells(21, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Avg Order Value]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"")"
    ws.Cells(21, 4).Value = 307417.18426501035
    ws.Cells(21, 5).Formula = "=IF(ISNUMBER(C21),IF(ABS(C21-D21)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C21:D21").NumberFormat = "#,##0"
    ws.Cells(22, 1).Value = "GM% th" & ChrW(225) & "ng 4/2025"
    ws.Cells(22, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Gross Margin %]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"")"
    ws.Cells(22, 4).Value = 0.33798595794117153
    ws.Cells(22, 5).Formula = "=IF(ISNUMBER(C22),IF(ABS(C22-D22)<=0.0005,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C22:D22").NumberFormat = "0.0%"
    ws.Cells(23, 1).Value = "YoY th" & ChrW(225) & "ng 4/2025"
    ws.Cells(23, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Revenue YoY %]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"")"
    ws.Cells(23, 4).Value = 0.24011341902399505
    ws.Cells(23, 5).Formula = "=IF(ISNUMBER(C23),IF(ABS(C23-D23)<=0.0005,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C23:D23").NumberFormat = "0.0%"
    ws.Cells(24, 1).Value = "Kh" & ChrW(225) & "ch ho" & ChrW(7841) & "t " & ChrW(273) & ChrW(7897) & "ng th" & ChrW(225) & "ng 4/2025"
    ws.Cells(24, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Active Customers]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"")"
    ws.Cells(24, 4).Value = 670.0
    ws.Cells(24, 5).Formula = "=IF(ISNUMBER(C24),IF(ABS(C24-D24)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C24:D24").NumberFormat = "#,##0"
    ws.Cells(25, 1).Value = "Target th" & ChrW(225) & "ng 4/2025"
    ws.Cells(25, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Target]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"")"
    ws.Cells(25, 4).Value = 288000000.0
    ws.Cells(25, 5).Formula = "=IF(ISNUMBER(C25),IF(ABS(C25-D25)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C25:D25").NumberFormat = "#,##0"
    ws.Cells(26, 1).Value = "Waterfall th" & ChrW(225) & "ng 3: t" & ChrW(7893) & "ng Variance (th" & ChrW(7921) & "c t" & ChrW(7871) & " " & ChrW(8722) & " target)"
    ws.Cells(26, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Variance]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(26, 4).Value = -15392000.0
    ws.Cells(26, 5).Formula = "=IF(ISNUMBER(C26),IF(ABS(C26-D26)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C26:D26").NumberFormat = "#,##0"
    ws.Cells(27, 1).Value = "Top 5 th" & ChrW(225) & "ng 4: H" & ChrW(7897) & "p qu" & ChrW(224) & " T" & ChrW(7871) & "t cao c" & ChrW(7845) & "p (P027)"
    ws.Cells(27, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"",""[Product].[Product_Code].&[P027]"")"
    ws.Cells(27, 4).Value = 35100000.0
    ws.Cells(27, 5).Formula = "=IF(ISNUMBER(C27),IF(ABS(C27-D27)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C27:D27").NumberFormat = "#,##0"
    ws.Cells(28, 1).Value = "Top 5 th" & ChrW(225) & "ng 4: Ch" & ChrW(7843) & "o ch" & ChrW(7889) & "ng d" & ChrW(237) & "nh 26cm (P039)"
    ws.Cells(28, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"",""[Product].[Product_Code].&[P039]"")"
    ws.Cells(28, 4).Value = 29155000.0
    ws.Cells(28, 5).Formula = "=IF(ISNUMBER(C28),IF(ABS(C28-D28)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C28:D28").NumberFormat = "#,##0"
    ws.Cells(29, 1).Value = "Top 5 th" & ChrW(225) & "ng 4: D" & ChrW(7847) & "u g" & ChrW(7897) & "i 650ml (P028)"
    ws.Cells(29, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"",""[Product].[Product_Code].&[P028]"")"
    ws.Cells(29, 4).Value = 26565000.0
    ws.Cells(29, 5).Formula = "=IF(ISNUMBER(C29),IF(ABS(C29-D29)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C29:D29").NumberFormat = "#,##0"
    ws.Cells(30, 1).Value = "Top 5 th" & ChrW(225) & "ng 4: M" & ChrW(236) & " " & ChrW(259) & "n li" & ChrW(7873) & "n t" & ChrW(244) & "m chua cay (th" & ChrW(249) & "ng 30 g" & ChrW(243) & "i) (P013)"
    ws.Cells(30, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"",""[Product].[Product_Code].&[P013]"")"
    ws.Cells(30, 4).Value = 19712000.0
    ws.Cells(30, 5).Formula = "=IF(ISNUMBER(C30),IF(ABS(C30-D30)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C30:D30").NumberFormat = "#,##0"
    ws.Cells(31, 1).Value = "Top 5 th" & ChrW(225) & "ng 4: B" & ChrW(7897) & "t gi" & ChrW(7863) & "t 3kg (P034)"
    ws.Cells(31, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"",""[Product].[Product_Code].&[P034]"")"
    ws.Cells(31, 4).Value = 19276000.0
    ws.Cells(31, 5).Formula = "=IF(ISNUMBER(C31),IF(ABS(C31-D31)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C31:D31").NumberFormat = "#,##0"
    ws.Cells(32, 1).Value = "Combo: doanh thu th" & ChrW(225) & "ng 1"
    ws.Cells(32, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[1]"")"
    ws.Cells(32, 4).Value = 425850000.0
    ws.Cells(32, 5).Formula = "=IF(ISNUMBER(C32),IF(ABS(C32-D32)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C32:D32").NumberFormat = "#,##0"
    ws.Cells(33, 1).Value = "Combo: target th" & ChrW(225) & "ng 1"
    ws.Cells(33, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Target]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[1]"")"
    ws.Cells(33, 4).Value = 353000000.0
    ws.Cells(33, 5).Formula = "=IF(ISNUMBER(C33),IF(ABS(C33-D33)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C33:D33").NumberFormat = "#,##0"
    ws.Cells(34, 1).Value = "Combo: doanh thu th" & ChrW(225) & "ng 2"
    ws.Cells(34, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[2]"")"
    ws.Cells(34, 4).Value = 245996000.0
    ws.Cells(34, 5).Formula = "=IF(ISNUMBER(C34),IF(ABS(C34-D34)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C34:D34").NumberFormat = "#,##0"
    ws.Cells(35, 1).Value = "Combo: target th" & ChrW(225) & "ng 2"
    ws.Cells(35, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Target]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[2]"")"
    ws.Cells(35, 4).Value = 298000000.0
    ws.Cells(35, 5).Formula = "=IF(ISNUMBER(C35),IF(ABS(C35-D35)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C35:D35").NumberFormat = "#,##0"
    ws.Cells(36, 1).Value = "Combo: doanh thu th" & ChrW(225) & "ng 3"
    ws.Cells(36, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(36, 4).Value = 269608000.0
    ws.Cells(36, 5).Formula = "=IF(ISNUMBER(C36),IF(ABS(C36-D36)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C36:D36").NumberFormat = "#,##0"
    ws.Cells(37, 1).Value = "Combo: target th" & ChrW(225) & "ng 3"
    ws.Cells(37, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Target]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[3]"")"
    ws.Cells(37, 4).Value = 285000000.0
    ws.Cells(37, 5).Formula = "=IF(ISNUMBER(C37),IF(ABS(C37-D37)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C37:D37").NumberFormat = "#,##0"
    ws.Cells(38, 1).Value = "Combo: doanh thu th" & ChrW(225) & "ng 4"
    ws.Cells(38, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Revenue]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"")"
    ws.Cells(38, 4).Value = 296965000.0
    ws.Cells(38, 5).Formula = "=IF(ISNUMBER(C38),IF(ABS(C38-D38)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C38:D38").NumberFormat = "#,##0"
    ws.Cells(39, 1).Value = "Combo: target th" & ChrW(225) & "ng 4"
    ws.Cells(39, 3).Formula = "=CUBEVALUE(""ThisWorkbookDataModel"",""[Measures].[Total Target]"",""[Calendar].[Nam].&[2025]"",""[Calendar].[Thang].&[4]"")"
    ws.Cells(39, 4).Value = 288000000.0
    ws.Cells(39, 5).Formula = "=IF(ISNUMBER(C39),IF(ABS(C39-D39)<=0.5,""" & ChrW(10004) & """,""" & ChrW(10008) & """),""" & ChrW(10008) & """)"
    ws.Range("C39:D39").NumberFormat = "#,##0"
    ws.Range("A3").Formula = "=""S" & ChrW(7889) & " " & ChrW(244) & " kh" & ChrW(7899) & "p: ""&COUNTIF(E5:E39,""" & ChrW(10004) & """)&""/""&COUNTA(A5:A39)"
    ws.Range("A3").Font.Bold = True
    ws.Columns("A").ColumnWidth = 62
    ws.Columns("B").ColumnWidth = 2
    ws.Columns("C:D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 9
    ws.Range("E5:E39").HorizontalAlignment = xlCenter
    ws.Range("E5:E39").Font.Name = "Segoe UI Symbol"
    Chk "kiem tra"
    L "stage8 ok"
End Sub

Private Sub Stage9_Finish()
    Dim nm, i As Long, ws As Worksheet
    On Error Resume Next
    nm = Array("PV_Dash", "Kiem_tra_B6", "Du_bao", "Brief", "Cai_dat", "Data_Dash", "Dashboard")
    For i = 0 To UBound(nm)
        wb.Worksheets(nm(i)).Move Before:=wb.Worksheets(1)
        Chk "move " & nm(i)
    Next
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateUntilAsyncQueriesDone
    Application.CalculateFull
    Application.CalculateUntilAsyncQueriesDone
    Chk "calc"
    For Each ws In wb.Worksheets
        If ws.Name = "Dashboard" Or ws.Name = "Data_Dash" Or ws.Name = "Cai_dat" Or ws.Name = "Brief" Or ws.Name = "Du_bao" Or ws.Name = "Kiem_tra_B6" Then
            ws.Activate
            ActiveWindow.DisplayGridlines = False
            If ws.Name = "Dashboard" Then ActiveWindow.Zoom = 80
        End If
    Next
    wb.Worksheets("Dashboard").Activate
    wb.Worksheets("Dashboard").Range("A1").Select
    Chk "finish"
    L "Data_Dash B4=" & wb.Worksheets("Data_Dash").Range("B4").Text & " C4=" & wb.Worksheets("Data_Dash").Range("C4").Text & " D4=" & wb.Worksheets("Data_Dash").Range("D4").Text & " E4=" & wb.Worksheets("Data_Dash").Range("E4").Text
    L "Kiem_tra A3=" & wb.Worksheets("Kiem_tra_B6").Range("A3").Text
    L "stage9 ok"
End Sub
