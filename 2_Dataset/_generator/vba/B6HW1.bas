Attribute VB_Name = "B6HW1"
Option Explicit
' Dung bai tap ve nha B6 (de + _Answer) tren Excel Windows. Sinh boi gen_b6_hw_vba.py.
' Chay tu Immediate window: B6HW1.BuildDe  hoac  B6HW1.BuildAnswer   (nhat ky: %TEMP%\b6_hw_log.txt)
' Cach dung: SaveAs dap an B5 thanh Homework\B6_BTVN_MCI_Mart_Dashboard_Marketing.xlsx (de) hoac ..._Answer.xlsx, mo kem mot workbook trong chua macro.

Private wb As Workbook
Private gF As Integer
Private isAns As Boolean
Private Const CN As String = "ThisWorkbookDataModel"
Private Const NAM As String = ",""[Calendar].[Nam].&[2025]"""

Private Function U(ByVal s As String) As String
    Dim p As Long, r As String
    p = InStr(s, "\u")
    Do While p > 0
        r = r & Left$(s, p - 1) & ChrW(CLng("&H" & Mid$(s, p + 2, 4)))
        s = Mid$(s, p + 6)
        p = InStr(s, "\u")
    Loop
    U = r & s
End Function

Private Function LogPath() As String
    LogPath = Environ("TEMP") & "\b6_hw_log.txt"
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
    rng.WrapText = True
End Sub

Private Function NewPT(ws As Worksheet, ByVal addr As String, ByVal nm As String) As PivotTable
    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create(SourceType:=xlExternal, SourceData:=wb.Connections(CN), Version:=6)
    Set NewPT = pc.CreatePivotTable(TableDestination:=ws.Range(addr), TableName:=nm)
End Function

Private Sub PageField(pt As PivotTable, ByVal cube As String, ByVal key As String)
    pt.CubeFields(cube).Orientation = xlPageField
    pt.PivotFields(cube & "." & Mid(cube, InStrRev(cube, "["))).CurrentPageName = cube & ".&[" & key & "]"
End Sub

Public Sub BuildDe()
    Run "B6_BTVN_MCI_Mart_Dashboard_Marketing.xlsx", False
End Sub

Public Sub BuildAnswer()
    Run "B6_BTVN_MCI_Mart_Dashboard_Marketing_Answer.xlsx", True
End Sub

Private Sub Run(ByVal fileName As String, ByVal ans As Boolean)
    gF = FreeFile
    Open LogPath() For Output As #gF
    L "START " & fileName & " " & Now
    Set wb = Workbooks(fileName)
    isAns = ans
    wb.Activate
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Cleanup
    MakePivotsCu
    MakeDashboardCu
    MakeDeBai
    MakeBrief
    MakeNhatKy
    MakeInsight
    If isAns Then
        MakeCaiDat
        MakeNewPivots
        MakeDataDash
        MakeDashboardNew
        MakeKiemTra
    End If
    Finish
    Application.DisplayAlerts = True
    L "DONE " & Now
    Close #gF
End Sub

Private Sub Cleanup()
    Dim i As Long, nm
    On Error Resume Next
    For i = wb.SlicerCaches.Count To 1 Step -1
        If Left(wb.SlicerCaches(i).Name, 14) = "Slicer_Thang_C" Or wb.SlicerCaches(i).Name = "Slicer_Thang" Then wb.SlicerCaches(i).Delete
        Chk "del slicer cache"
    Next
    For Each nm In Array("De_bai", "Brief", "Dashboard_Cu", "Nhat_ky_doi_soat", "Insight", "Insight_mau", "PV_Cu", "PV_New", "Dashboard", "Data_Dash", "Cai_dat", "Kiem_tra")
        DelSheet CStr(nm)
    Next
    L "cleanup ok"
End Sub

' ---------- Pivot cho Dashboard_Cu ----------
Private Sub MakePivotsCu()
    Dim ws As Worksheet, ptSeg As PivotTable, ptReg As PivotTable
    Dim wsD As Worksheet, sc As SlicerCache, sl As Slicer
    On Error Resume Next
    Set ws = GetSheet("PV_Cu")
    Set wsD = GetSheet("Dashboard_Cu")
    Title ws, U("PivotTable c\u1EE7a dashboard c\u0169 (Dashboard_Cu)"), U("Ngu\u1ED3n cho pie v\u00E0 bar c\u1EE7a dashboard c\u0169. Kh\u00F4ng s\u1EEDa \u1EDF \u0111\u00E2y: vi\u1EC7c c\u1EE7a b\u1EA1n l\u00E0 ki\u1EC3m tra Dashboard_Cu.")
    Set ptSeg = NewPT(ws, "A8", "pvtSegment_Cu")
    Chk "pt seg create"
    PageField ptSeg, "[Calendar].[Nam]", "2025"
    ptSeg.CubeFields("[Customer].[Segment]").Orientation = xlRowField
    ptSeg.CubeFields("[Measures].[Total Revenue]").Orientation = xlDataField
    Chk "pt seg fields"
    Set ptReg = NewPT(ws, "E8", "pvtMien_Cu")
    Chk "pt reg create"
    PageField ptReg, "[Calendar].[Nam]", "2025"
    PageField ptReg, "[Calendar].[Thang]", "4"
    ptReg.CubeFields("[Store].[Region]").Orientation = xlRowField
    ptReg.CubeFields("[Measures].[Active Customers]").Orientation = xlDataField
    Chk "pt reg fields"
    Set sc = wb.SlicerCaches.Add2(ptSeg, "[Calendar].[Thang]")
    Chk "slicer cu cache"
    sc.Name = "Slicer_Thang_Cu"
    Set sl = sc.Slicers.Add(wsD, , "Thang_Cu", U("Th\u00E1ng"), 560, 20, 330, 60)
    Chk "slicer cu add"
    sl.NumberOfColumns = 6
    sc.VisibleSlicerItemsList = Array("[Calendar].[Thang].&[4]")
    Chk "slicer cu select"
    L "slicer cu pivots=" & sc.PivotTables.Count
End Sub

Private Sub MakeDashboardCu()
    Dim ws As Worksheet, wsP As Worksheet, shp As Shape, ch As Chart, k As Long, c0 As Long
    Dim lab, f
    On Error Resume Next
    Set ws = GetSheet("Dashboard_Cu")
    Set wsP = GetSheet("PV_Cu")
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:M").ColumnWidth = 10
    ws.Range("B2:M2").Merge
    ws.Range("B2").Value = U("Dashboard kh\u00E1ch h\u00E0ng th\u00E1ng 4/2025 (b\u1EA3n c\u0169 c\u1EE7a \u0111\u1ED3ng nghi\u1EC7p)")
    ws.Range("B2").Font.Size = 16
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(255, 255, 255)
    ws.Range("B2:M2").Interior.Color = RGB(89, 89, 89)
    ws.Rows(2).RowHeight = 30
    ws.Range("B3:M3").Merge
    ws.Range("B3").Value = U("C\u00E2u h\u1ECFi c\u1EE7a Marketing Director: Xem kh\u00E1ch h\u00E0ng")
    ws.Range("B3").Font.Italic = True
    ws.Range("B3").Font.Color = RGB(88, 88, 88)
    ' 3 the KPI: khong target, khong RAG, Active Customers go cung
    lab = Array("Doanh thu", U("S\u1ED1 \u0111\u01A1n"), "Active Customers")
    For k = 0 To 2
        c0 = 2 + 4 * k
        ws.Range(ws.Cells(5, c0), ws.Cells(5, c0 + 3)).Merge
        ws.Cells(5, c0).Value = lab(k)
        ws.Cells(5, c0).Font.Bold = True
        ws.Range(ws.Cells(6, c0), ws.Cells(7, c0 + 3)).Merge
        ws.Range(ws.Cells(5, c0), ws.Cells(7, c0 + 3)).HorizontalAlignment = xlCenter
        ws.Cells(6, c0).Font.Size = 22
        ws.Cells(6, c0).Font.Bold = True
        ws.Range(ws.Cells(5, c0), ws.Cells(7, c0 + 3)).BorderAround xlContinuous, xlThin
    Next
    ws.Range("B6").Formula = CVF("Total Revenue", NAM & ",""[Calendar].[Thang].&[4]""")
    ws.Range("B6").NumberFormat = "#,##0"
    ws.Range("F6").Formula = CVF("Total Orders", NAM & ",""[Calendar].[Thang].&[4]""")
    ws.Range("F6").NumberFormat = "#,##0"
    ws.Range("J6").Value = 713
    ws.Range("J6").NumberFormat = "#,##0"
    Chk "cu cards"
    ' pie 3D theo phan khuc
    Set shp = ws.Shapes.AddChart2(-1, xl3DPie, ws.Range("B9").Left, ws.Range("B9").Top, ws.Range("B9:G9").Width, 250)
    shp.Name = "chtPie3D_Cu"
    Set ch = shp.Chart
    ch.SetSourceData Source:=wsP.PivotTables("pvtSegment_Cu").TableRange1
    Chk "cu pie setsource"
    ch.HasTitle = True
    ch.ChartTitle.Text = U("Doanh thu theo ph\u00E2n kh\u00FAc")
    ch.ShowAllFieldButtons = False
    ch.SeriesCollection(1).HasDataLabels = True
    Chk "cu pie fmt"
    ' bar theo mien, truc cat khuc
    Set shp = ws.Shapes.AddChart2(-1, xlColumnClustered, ws.Range("H9").Left, ws.Range("B9").Top, ws.Range("H9:M9").Width, 250)
    shp.Name = "chtBar_Cu"
    Set ch = shp.Chart
    ch.SetSourceData Source:=wsP.PivotTables("pvtMien_Cu").TableRange1
    Chk "cu bar setsource"
    ch.HasTitle = True
    ch.ChartTitle.Text = U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng theo mi\u1EC1n")
    ch.HasLegend = False
    ch.ShowAllFieldButtons = False
    ch.Axes(xlValue).MinimumScale = 100
    ch.SeriesCollection(1).HasDataLabels = True
    Chk "cu bar fmt"
    ws.Range("B31").Value = U("Ghi ch\u00FA c\u1EE7a \u0111\u1ED3ng nghi\u1EC7p: s\u1ED1 li\u1EC7u \u0111\u00E3 ki\u1EC3m tra, g\u1EEDi Marketing Director.")
    ws.Range("B31").Font.Italic = True
    ws.Range("B31").Font.Color = RGB(88, 88, 88)
    L "dashboard_cu ok"
End Sub

' ---------- De bai ----------
Private Sub MakeDeBai()
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = GetSheet("De_bai")
    ws.Range("A1").Value = U("B\u00C0I T\u1EACP V\u1EC0 NH\u00C0 BU\u1ED4I 6 (B\u1EAET BU\u1ED8C) \u2014 KI\u1EC2M TRA V\u00C0 L\u00C0M L\u1EA0I DASHBOARD KH\u00C1CH H\u00C0NG")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Color = RGB(32, 56, 100)
    ws.Range("A2").Value = U("N\u1ED9p tr\u01B0\u1EDBc Bu\u1ED5i 7. \u0110\u1EB7t t\u00EAn file: B6_BTVN_<HoTen>.xlsx (l\u00E0m ngay tr\u00EAn file n\u00E0y). T\u00ECnh hu\u1ED1ng: Marketing Director xin m\u1ED9t dashboard theo d\u00F5i kh\u00E1ch h\u00E0ng th\u00E1ng 4/2025. \u0110\u1ED3ng nghi\u1EC7p \u0111\u00E3 l\u00E0m m\u1ED9t b\u1EA3n (sheet Dashboard_Cu). B\u1EA1n l\u00E0 BI analyst: ki\u1EC3m tra b\u1EA3n c\u0169 c\u00F3 \u0111\u00E1ng tin kh\u00F4ng, r\u1ED3i l\u00E0m l\u1EA1i. B\u00E0i t\u1EADp d\u01B0\u1EE3t cho b\u00E0i capstone: ti\u00EAu ch\u00ED C1 (Brief), C5 (Dashboard) v\u00E0 C6 (Insight).")
    ws.Range("A3").Value = U("N\u1ED9p b\u00E0i t\u1EA1i c\u1ED5ng h\u1ECDc vi\u00EAn MCI (hocvien.mcivietnam.com). Kh\u00F4ng b\u1EA5m Refresh All trong file n\u00E0y: d\u1EEF li\u1EC7u \u0111\u00E3 n\u1EB1m s\u1EB5n trong Data Model. Ch\u1EC9 d\u00F9ng 6 b\u1EA3ng c\u1EE7a Data Model; m\u1ECDi con s\u1ED1 tr\u00EAn dashboard m\u1EDBi l\u1EA5y t\u1EEB measure c\u00F3 s\u1EB5n (kh\u00F4ng g\u00F5 c\u1EE9ng).")
    ws.Range("A5:C5").Value = Array("#", U("Vi\u1EC7c c\u1EA7n l\u00E0m"), U("Y\u00EAu c\u1EA7u"))
    Hdr ws.Range("A5:C5")
    ws.Range("A6:C6").Value = Array(1, U("\u0110\u1ED1i so\u00E1t"), U("\u0110\u1ED1i so\u00E1t t\u1EEBng con s\u1ED1 v\u00E0 t\u1EEBng bi\u1EC3u \u0111\u1ED3 c\u1EE7a Dashboard_Cu v\u1EDBi measure trong Data Model. Ghi v\u00E0o Nhat_ky_doi_soat m\u1ECDi \u0111i\u1EC3m sai: l\u1ED7i \u00B7 s\u1ED1 tr\u00EAn dashboard c\u0169 \u00B7 s\u1ED1 \u0111\u00FAng \u00B7 c\u00E1ch s\u1EEDa. C\u00F3 5 l\u1ED7i; m\u1ED7i l\u1ED7i m\u1ED9t d\u00F2ng. (D2: s\u1ED1 tr\u00F4ng h\u1EE3p l\u00FD ch\u01B0a ch\u1EAFc \u0111\u00FAng.)"))
    ws.Range("A7:C7").Value = Array(2, U("Vi\u1EBFt brief"), U("Sheet Brief: vi\u1EBFt brief cho Marketing Director g\u1ED3m ng\u01B0\u1EDDi xem, 3\u20135 c\u00E2u h\u1ECFi \u0111o l\u01B0\u1EDDng \u0111\u01B0\u1EE3c (ch\u1EC9 s\u1ED1 \u00B7 chi\u1EC1u \u00B7 th\u1EDDi gian \u00B7 m\u1ED1c so s\u00E1nh), t\u1EA7n su\u1EA5t refresh, quy\u1EBFt \u0111\u1ECBnh m\u00E0 dashboard h\u1ED7 tr\u1EE3. Kh\u00F4ng d\u00F9ng c\u00E2u h\u1ECFi d\u1EA1ng ch\u1EE7 \u0111\u1EC1 nh\u01B0 ""Xem kh\u00E1ch h\u00E0ng"". (C1 Brief)"))
    ws.Range("A8:C8").Value = Array(3, U("D\u1EF1ng dashboard 1 trang"), U("Sheet Dashboard (t\u1EA1o m\u1EDBi): t\u1ED1i \u0111a 8 visual; th\u1EBB KPI l\u1EA5y b\u1EB1ng CUBEVALUE, c\u00F3 target v\u00E0 RAG k\u00E8m bi\u1EC3u t\u01B0\u1EE3ng v\u00E0 nh\u00E3n ch\u1EEF; Slicer Th\u00E1ng n\u1ED1i \u0111\u1EE7 m\u1ECDi Pivot; m\u1ECDi chart c\u00F3 alt text; ti\u00EAu \u0111\u1EC1 chart l\u00E0 k\u1EBFt lu\u1EADn. Target kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng: 700 m\u1ED7i th\u00E1ng. (C5 Dashboard)"))
    ws.Range("A9:C9").Value = Array(4, U("Vi\u1EBFt insight"), U("Sheet Insight: 3 c\u00E2u theo khung SCR (Situation \u00B7 Complication \u00B7 Resolution), k\u00E8m m\u1ED9t h\u00E0nh \u0111\u1ED9ng c\u1EE5 th\u1EC3 cho Marketing Director. Con s\u1ED1 ph\u1EA3i c\u00F3 m\u1ED1c so s\u00E1nh. (C6 Insight)"))
    ws.Range("A11").Value = U("T\u1EF1 ki\u1EC3m tra tr\u01B0\u1EDBc khi n\u1ED9p")
    ws.Range("A11").Font.Bold = True
    ws.Range("A12").Value = U("\u2610 Th\u00E1ng 4/2025: kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng 670 \u00B7 target 700 \u00B7 m\u1EE9c \u0111\u1EA1t 95,7% (\u25B2 C\u1EA3nh b\u00E1o)")
    ws.Range("A13").Value = U("\u2610 Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng theo ph\u00E2n kh\u00FAc: Th\u01B0\u1EDDng 360 \u00B7 Th\u00E2n thi\u1EBFt 251 \u00B7 VIP 59 (c\u1ED9ng l\u1EA1i b\u1EB1ng 670)")
    ws.Range("A14").Value = U("\u2610 Doanh thu theo ph\u00E2n kh\u00FAc: Th\u01B0\u1EDDng 138.190.000 \u00B7 Th\u00E2n thi\u1EBFt 119.096.000 \u00B7 VIP 39.679.000 (t\u1ED5ng 296.965.000)")
    ws.Range("A15").Value = U("\u2610 C\u1ED9ng kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng theo ba mi\u1EC1n cho 713, l\u1EDBn h\u01A1n 670: m\u1ED9t kh\u00E1ch mua \u1EDF nhi\u1EC1u mi\u1EC1n \u0111\u01B0\u1EE3c \u0111\u1EBFm \u1EDF m\u1ED7i mi\u1EC1n; v\u00EC v\u1EADy s\u1ED1 n\u00E0y kh\u00F4ng \u0111\u01B0\u1EE3c c\u1ED9ng tay")
    ws.Range("A16").Value = U("\u2610 B\u1EA5m Slicer sang th\u00E1ng 3: m\u1ECDi bi\u1EC3u \u0111\u1ED3 v\u00E0 th\u1EBB KPI tr\u00EAn dashboard m\u1EDBi \u0111\u1EC1u \u0111\u1ED5i")
    ws.Range("A17").Value = U("G\u1EE3i \u00FD: xem handout B6 (CUBEVALUE, Report Connections, RAG) v\u00E0 KPI Definition Card c\u1EE7a Bu\u1ED5i 2.")
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 110
    ws.Range("A2:C2").Merge
    ws.Range("A3:C3").Merge
    ws.Range("A2:C3").WrapText = True
    ws.Range("A2:C3").VerticalAlignment = xlTop
    ws.Rows(2).RowHeight = 48
    ws.Rows(3).RowHeight = 32
    ws.Range("C6:C9").WrapText = True
    ws.Range("A6:C9").VerticalAlignment = xlTop
    Chk "de_bai"
    L "de_bai ok"
End Sub

Private Sub MakeBrief()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetSheet("Brief")
    Title ws, U("BRIEF \u2014 DASHBOARD KH\u00C1CH H\u00C0NG CHO MARKETING DIRECTOR"), IIf(isAns, U("B\u1EA3n m\u1EABu (\u0111\u00E1p \u00E1n tham kh\u1EA3o). C\u00E2u h\u1ECFi c\u00F3 \u0111\u1EE7 ch\u1EC9 s\u1ED1 \u00B7 chi\u1EC1u \u00B7 th\u1EDDi gian \u00B7 m\u1ED1c so s\u00E1nh."), U("\u0110i\u1EC1n v\u00E0o c\u1ED9t N\u1ED9i dung. M\u1ED7i c\u00E2u h\u1ECFi ph\u1EA3i \u0111o l\u01B0\u1EDDng \u0111\u01B0\u1EE3c."))
    ws.Range("A4:C4").Value = Array(U("M\u1EE5c"), U("G\u1EE3i \u00FD"), U("N\u1ED9i dung"))
    Hdr ws.Range("A4:C4")
    ws.Range("A5:B5").Value = Array(U("Ng\u01B0\u1EDDi xem"), "Executive, operational hay self-service?")
    ws.Range("A6:B6").Value = Array(U("C\u00E2u h\u1ECFi 1"), U("Ch\u1EC9 s\u1ED1 \u00B7 chi\u1EC1u \u00B7 th\u1EDDi gian \u00B7 m\u1ED1c so s\u00E1nh"))
    ws.Range("A7:B7").Value = Array(U("C\u00E2u h\u1ECFi 2"), "")
    ws.Range("A8:B8").Value = Array(U("C\u00E2u h\u1ECFi 3"), "")
    ws.Range("A9:B9").Value = Array(U("C\u00E2u h\u1ECFi 4 (t\u00F9y ch\u1ECDn)"), "")
    ws.Range("A10:B10").Value = Array(U("T\u1EA7n su\u1EA5t refresh"), U("Bao l\u00E2u m\u1ED9t l\u1EA7n?"))
    ws.Range("A11:B11").Value = Array(U("Quy\u1EBFt \u0111\u1ECBnh"), U("Dashboard gi\u00FAp quy\u1EBFt \u0111\u1ECBnh g\u00EC?"))
    If isAns Then
        ws.Range("C5").Value = U("Marketing Director: executive, xem trong h\u1ECDp h\u1EB1ng th\u00E1ng.")
        ws.Range("C6").Value = U("Q1. Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng th\u00E1ng 4/2025 \u0111\u1EA1t bao nhi\u00EAu % target 700, so v\u1EDBi c\u00E1c th\u00E1ng tr\u01B0\u1EDBc?")
        ws.Range("C7").Value = U("Q2. Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng th\u00E1ng 4 chia theo ph\u00E2n kh\u00FAc (Th\u01B0\u1EDDng, Th\u00E2n thi\u1EBFt, VIP) th\u1EBF n\u00E0o, v\u00E0 ph\u00E2n kh\u00FAc n\u00E0o \u0111\u00F3ng g\u00F3p nhi\u1EC1u doanh thu nh\u1EA5t?")
        ws.Range("C8").Value = U("Q3. Doanh thu b\u00ECnh qu\u00E2n tr\u00EAn m\u1ED7i kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng c\u1EE7a t\u1EEBng ph\u00E2n kh\u00FAc trong th\u00E1ng 4 l\u00E0 bao nhi\u00EAu?")
        ws.Range("C9").Value = U("Q4. Mi\u1EC1n n\u00E0o c\u00F3 nhi\u1EC1u kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng nh\u1EA5t trong th\u00E1ng 4? (kh\u00F4ng c\u1ED9ng ba mi\u1EC1n th\u00E0nh t\u1ED5ng kh\u00E1ch)")
        ws.Range("C10").Value = U("H\u1EB1ng th\u00E1ng, khi c\u00F3 file d\u1EEF li\u1EC7u th\u00E1ng m\u1EDBi.")
        ws.Range("C11").Value = U("Ch\u1ECDn ph\u00E2n kh\u00FAc \u01B0u ti\u00EAn cho chi\u1EBFn d\u1ECBch gi\u1EEF ch\u00E2n v\u00E0 k\u00EDch ho\u1EA1t l\u1EA1i kh\u00E1ch \u0111\u1EC3 \u0111\u1EA1t target kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng.")
    End If
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 110
    ws.Range("A5:C11").WrapText = True
    ws.Range("A5:C11").VerticalAlignment = xlTop
    ws.Range("A5:A11").Font.Bold = True
    If Not isAns Then ws.Range("C5:C11").Interior.Color = RGB(255, 242, 204)
    Chk "brief"
    L "brief ok"
End Sub

Private Sub MakeNhatKy()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetSheet("Nhat_ky_doi_soat")
    Title ws, U("NH\u1EACT K\u00DD \u0110\u1ED0I SO\u00C1T \u2014 DASHBOARD_CU \u2194 DATA MODEL"), U("M\u1ED7i l\u1ED7i m\u1ED9t d\u00F2ng: l\u1ED7i g\u00EC, \u1EDF \u0111\u00E2u, s\u1ED1 c\u0169, s\u1ED1 \u0111\u00FAng, c\u00E1ch s\u1EEDa.")
    ws.Range("A4:F4").Value = Array("#", U("L\u1ED7i (\u1EDF \u0111\u00E2u)"), U("S\u1ED1 tr\u00EAn dashboard c\u0169"), U("S\u1ED1 \u0111\u00FAng"), U("Nguy\u00EAn nh\u00E2n"), U("C\u00E1ch s\u1EEDa"))
    Hdr ws.Range("A4:F4")
    If isAns Then
        ws.Range("A5:F5").Value = Array(1, U("Th\u1EBB Active Customers (g\u00F5 c\u1EE9ng)"), "713", "670", U("G\u00F5 tay t\u1ED5ng ba mi\u1EC1n (234 + 151 + 328). M\u1ED9t kh\u00E1ch mua \u1EDF nhi\u1EC1u mi\u1EC1n b\u1ECB \u0111\u1EBFm nhi\u1EC1u l\u1EA7n: DISTINCTCOUNT kh\u00F4ng c\u1ED9ng d\u1ED3n."), U("Thay b\u1EB1ng CUBEVALUE g\u1ECDi measure [Active Customers], l\u1ECDc b\u1EB1ng Slicer Th\u00E1ng."))
        ws.Range("A6:F6").Value = Array(2, U("Th\u1EBB KPI kh\u00F4ng c\u00F3 target v\u00E0 RAG"), U("Ch\u1EC9 c\u00F3 s\u1ED1, kh\u00F4ng c\u00F3 m\u1ED1c so s\u00E1nh"), U("Target 700 \u00B7 m\u1EE9c \u0111\u1EA1t 95,7% \u25B2 C\u1EA3nh b\u00E1o"), U("Ng\u01B0\u1EDDi xem kh\u00F4ng bi\u1EBFt \u0111\u1EA1t hay ch\u01B0a; thi\u1EBFu m\u1ED1c so s\u00E1nh. RAG ch\u1EC9 b\u1EB1ng m\u00E0u th\u00EC ng\u01B0\u1EDDi m\u00F9 m\u00E0u kh\u00F4ng \u0111\u1ECDc \u0111\u01B0\u1EE3c."), U("Th\u00EAm target (Cai_dat), m\u1EE9c \u0111\u1EA1t = gi\u00E1 tr\u1ECB \u00F7 target, bi\u1EC3u t\u01B0\u1EE3ng \u2714 \u25B2 \u2718 v\u00E0 nh\u00E3n ch\u1EEF; ng\u01B0\u1EE1ng \u1EDF m\u1ED9t \u00F4 tham s\u1ED1."))
        ws.Range("A7:F7").Value = Array(3, U("Pie 3D theo ph\u00E2n kh\u00FAc; tr\u1EE5c bar c\u1EAFt kh\u00FAc"), U("Pie 3D; tr\u1EE5c mi\u1EC1n b\u1EAFt \u0111\u1EA7u t\u1EEB 100"), U("Th\u01B0\u1EDDng 46,5% \u00B7 Th\u00E2n thi\u1EBFt 40,1% \u00B7 VIP 13,4%; tr\u1EE5c b\u1EAFt \u0111\u1EA7u t\u1EEB 0"), U("Hi\u1EC7u \u1EE9ng 3D l\u00E0m l\u1EC7ch t\u1EF7 l\u1EC7; tr\u1EE5c c\u1EAFt kh\u00FAc l\u00E0m c\u1ED9t mi\u1EC1n l\u1EDBn nh\u1EA5t tr\u00F4ng g\u1EA5p nhi\u1EC1u l\u1EA7n so v\u1EDBi th\u1EF1c t\u1EBF (328 so v\u1EDBi 151)."), U("\u0110\u1ED5i sang bar 2D x\u1EBFp h\u1EA1ng (ho\u1EB7c pie 2D khi ch\u1EC9 3 nh\u00F3m), tr\u1EE5c b\u1EAFt \u0111\u1EA7u t\u1EEB 0."))
        ws.Range("A8:F8").Value = Array(4, U("Slicer Th\u00E1ng ch\u1EC9 n\u1ED1i m\u1ED9t Pivot"), U("B\u1EA5m th\u00E1ng 3, bi\u1EC3u \u0111\u1ED3 theo mi\u1EC1n kh\u00F4ng \u0111\u1ED5i"), U("M\u1ECDi bi\u1EC3u \u0111\u1ED3 c\u00F9ng \u0111\u1ED5i"), U("Slicer sinh ra t\u1EEB m\u1ED9t Pivot n\u00EAn m\u1EB7c \u0111\u1ECBnh ch\u1EC9 \u0111i\u1EC1u khi\u1EC3n Pivot \u0111\u00F3."), U("Chu\u1ED9t ph\u1EA3i Slicer, Report Connections, tick m\u1ECDi Pivot c\u1EE7a dashboard."))
        ws.Range("A9:F9").Value = Array(5, U("C\u00E2u h\u1ECFi d\u1EA1ng ch\u1EE7 \u0111\u1EC1"), U("""Xem kh\u00E1ch h\u00E0ng"""), U("C\u00E2u h\u1ECFi c\u00F3 ch\u1EC9 s\u1ED1 \u00B7 chi\u1EC1u \u00B7 th\u1EDDi gian \u00B7 m\u1ED1c so s\u00E1nh"), U("Ch\u1EE7 \u0111\u1EC1 kh\u00F4ng cho bi\u1EBFt c\u1EA7n measure v\u00E0 chart n\u00E0o."), U("Vi\u1EBFt l\u1EA1i, v\u00ED d\u1EE5: Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng th\u00E1ng 4/2025 \u0111\u1EA1t bao nhi\u00EAu % target 700, so v\u1EDBi th\u00E1ng 3?"))
    End If
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E").ColumnWidth = 55
    ws.Columns("F").ColumnWidth = 55
    ws.Range("A5:F10").WrapText = True
    ws.Range("A5:F10").VerticalAlignment = xlTop
    ws.Range("C5:C9").NumberFormat = "@"
    Chk "nhat ky"
    L "nhat_ky ok"
End Sub

Private Sub MakeInsight()
    Dim ws As Worksheet
    On Error Resume Next
    If isAns Then
        Set ws = GetSheet("Insight_mau")
        Title ws, U("INSIGHT M\u1EAAU (SCR) \u2014 KH\u00C1CH H\u00C0NG TH\u00C1NG 4/2025"), U("\u0110\u00E1p \u00E1n tham kh\u1EA3o. M\u1ED7i c\u00E2u c\u00F3 con s\u1ED1 k\u00E8m m\u1ED1c so s\u00E1nh.")
        ws.Range("A4:B4").Value = Array(U("Ph\u1EA7n"), U("N\u1ED9i dung"))
        Hdr ws.Range("A4:B4")
        ws.Range("A5:B5").Value = Array(U("S \u2014 Situation"), U("Th\u00E1ng 4/2025 c\u00F3 670 kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng, \u0111\u1EA1t 95,7% target 700 (thi\u1EBFu 30 kh\u00E1ch), \u1EDF m\u1EE9c C\u1EA3nh b\u00E1o."))
        ws.Range("A6:B6").Value = Array(U("C \u2014 Complication"), U("Doanh thu b\u00ECnh qu\u00E2n m\u1ED7i kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng kh\u00E1c nhau r\u00F5 gi\u1EEFa c\u00E1c ph\u00E2n kh\u00FAc: VIP 672.5 ngh\u00ECn \u0111\u1ED3ng, Th\u00E2n thi\u1EBFt 474.5 ngh\u00ECn, Th\u01B0\u1EDDng 383.9 ngh\u00ECn; nh\u01B0ng VIP ch\u1EC9 c\u00F3 59 kh\u00E1ch (8,8% s\u1ED1 kh\u00E1ch) v\u00E0 nh\u00F3m Th\u01B0\u1EDDng 360 kh\u00E1ch (53,7%)."))
        ws.Range("A7:B7").Value = Array(U("R \u2014 Resolution"), U("Marketing \u01B0u ti\u00EAn k\u00EDch ho\u1EA1t l\u1EA1i kho\u1EA3ng 30 kh\u00E1ch c\u00F2n thi\u1EBFu, b\u1EAFt \u0111\u1EA7u t\u1EEB nh\u00F3m Th\u00E2n thi\u1EBFt (251 kh\u00E1ch, gi\u00E1 tr\u1ECB b\u00ECnh qu\u00E2n cao h\u01A1n nh\u00F3m Th\u01B0\u1EDDng 24%), v\u00E0 theo d\u00F5i m\u1EE9c \u0111\u1EA1t h\u1EB1ng th\u00E1ng tr\u00EAn dashboard."))
        ws.Columns("A").ColumnWidth = 22
        ws.Columns("B").ColumnWidth = 130
        ws.Range("A5:B7").WrapText = True
        ws.Range("A5:B7").VerticalAlignment = xlTop
        ws.Range("A5:A7").Font.Bold = True
    Else
        Set ws = GetSheet("Insight")
        Title ws, U("INSIGHT (SCR) \u2014 KH\u00C1CH H\u00C0NG TH\u00C1NG 4/2025"), U("3 c\u00E2u theo khung Situation \u00B7 Complication \u00B7 Resolution k\u00E8m m\u1ED9t h\u00E0nh \u0111\u1ED9ng cho Marketing Director. Con s\u1ED1 ph\u1EA3i c\u00F3 m\u1ED1c so s\u00E1nh.")
        ws.Range("A4:B4").Value = Array(U("Ph\u1EA7n"), U("N\u1ED9i dung c\u1EE7a b\u1EA1n"))
        Hdr ws.Range("A4:B4")
        ws.Range("A5").Value = U("S \u2014 Situation")
        ws.Range("A6").Value = U("C \u2014 Complication")
        ws.Range("A7").Value = U("R \u2014 Resolution")
        ws.Columns("A").ColumnWidth = 22
        ws.Columns("B").ColumnWidth = 130
        ws.Range("B5:B7").Interior.Color = RGB(255, 242, 204)
        ws.Range("A5:B7").WrapText = True
        ws.Rows("5:7").RowHeight = 45
    End If
    Chk "insight"
    L "insight ok"
End Sub

' ---------- Chi _Answer ----------
Private Sub MakeCaiDat()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetSheet("Cai_dat")
    Title ws, U("C\u00E0i \u0111\u1EB7t dashboard kh\u00E1ch h\u00E0ng"), U("M\u1ECDi m\u1ED1c so s\u00E1nh n\u1EB1m \u1EDF \u0111\u00E2y. \u00D4 n\u1EC1n v\u00E0ng l\u00E0 \u00F4 nh\u1EADp.")
    ws.Range("A4:C4").Value = Array("KPI", U("\u0110\u01A1n v\u1ECB"), "Target")
    Hdr ws.Range("A4:C4")
    ws.Range("A5:C5").Value = Array(U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng m\u1ED7i th\u00E1ng"), U("kh\u00E1ch"), 700)
    ws.Range("A6:C6").Value = Array(U("Gi\u00E1 tr\u1ECB \u0111\u01A1n trung b\u00ECnh (AOV)"), "VND", 300000)
    ws.Range("C5:C6").Interior.Color = RGB(255, 242, 204)
    ws.Range("C5:C6").NumberFormat = "#,##0"
    ws.Range("A8").Value = U("Ng\u01B0\u1EE1ng RAG (m\u1EE9c \u0111\u1EA1t = gi\u00E1 tr\u1ECB \u00F7 target)")
    ws.Range("A8").Font.Bold = True
    ws.Range("A9:C9").Value = Array(U("Tr\u1EA1ng th\u00E1i"), U("Bi\u1EC3u t\u01B0\u1EE3ng"), U("Ng\u01B0\u1EE1ng t\u1EEB"))
    Hdr ws.Range("A9:C9")
    ws.Range("A10:C10").Value = Array(U("\u0110\u1EA1t"), ChrW(10004), 1)
    ws.Range("A11:C11").Value = Array(U("C\u1EA3nh b\u00E1o"), ChrW(9650), 0.8)
    ws.Range("A12:C12").Value = Array(U("Kh\u00F4ng \u0111\u1EA1t"), ChrW(10008), "")
    ws.Range("C10:C11").NumberFormat = "0%"
    ws.Range("C10:C11").Interior.Color = RGB(255, 242, 204)
    ws.Range("B10:B12").HorizontalAlignment = xlCenter
    ws.Range("B10").Font.Color = RGB(12, 163, 12)
    ws.Range("B11").Font.Color = RGB(250, 178, 25)
    ws.Range("B12").Font.Color = RGB(208, 59, 59)
    ws.Columns("A").ColumnWidth = 36
    ws.Columns("B:C").ColumnWidth = 14
    Chk "cai_dat"
    L "cai_dat ok"
End Sub

Private Sub MakeNewPivots()
    Dim ws As Worksheet, ptA As PivotTable, ptB As PivotTable, ptC As PivotTable
    Dim sc As SlicerCache, sl As Slicer, wsD As Worksheet
    On Error Resume Next
    Set ws = GetSheet("PV_New")
    Set wsD = GetSheet("Dashboard")
    Title ws, U("PivotTable c\u1EE7a Dashboard m\u1EDBi (kh\u00F4ng s\u1EEDa tay)"), U("Slicer Th\u00E1ng n\u1ED1i c\u1EA3 ba Pivot.")
    Set ptA = NewPT(ws, "A8", "pvtKhachPK")
    PageField ptA, "[Calendar].[Nam]", "2025"
    ptA.CubeFields("[Customer].[Segment]").Orientation = xlRowField
    ptA.CubeFields("[Measures].[Active Customers]").Orientation = xlDataField
    Chk "pt A"
    Set ptB = NewPT(ws, "E8", "pvtDoanhThuPK")
    PageField ptB, "[Calendar].[Nam]", "2025"
    ptB.CubeFields("[Customer].[Segment]").Orientation = xlRowField
    ptB.CubeFields("[Measures].[Total Revenue]").Orientation = xlDataField
    Chk "pt B"
    Set ptC = NewPT(ws, "I8", "pvtKhachMien")
    PageField ptC, "[Calendar].[Nam]", "2025"
    ptC.CubeFields("[Store].[Region]").Orientation = xlRowField
    ptC.CubeFields("[Measures].[Active Customers]").Orientation = xlDataField
    Chk "pt C"
    Set sc = wb.SlicerCaches.Add2(ptA, "[Calendar].[Thang]")
    Chk "slicer cache"
    sc.Name = "Slicer_Thang"
    sc.PivotTables.AddPivotTable ptB
    Chk "conn B"
    sc.PivotTables.AddPivotTable ptC
    Chk "conn C"
    Set sl = sc.Slicers.Add(wsD, , "Thang", U("Th\u00E1ng"), wsD.Range("B4").Top, wsD.Range("U4").Left, 250, 110)
    Chk "slicer add"
    sl.NumberOfColumns = 4
    sc.VisibleSlicerItemsList = Array("[Calendar].[Thang].&[4]")
    Chk "slicer select"
    L "slicer new pivots=" & sc.PivotTables.Count
End Sub

Private Sub MakeDataDash()
    Dim ws As Worksheet, k As Long, r As Long
    Dim kpi, meas, fmt
    On Error Resume Next
    Set ws = GetSheet("Data_Dash")
    Title ws, U("Data_Dash: s\u1ED1 li\u1EC7u th\u00F4 cho Dashboard (CUBEVALUE, nghe Slicer Th\u00E1ng)"), ""
    ws.Range("A3:E3").Value = Array("KPI", U("Gi\u00E1 tr\u1ECB"), "Target", U("M\u1EE9c \u0111\u1EA1t"), U("Tr\u1EA1ng th\u00E1i"))
    Hdr ws.Range("A3:E3")
    kpi = Array(U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng"), "Doanh thu", "AOV")
    meas = Array("Active Customers", "Total Revenue", "Avg Order Value")
    For k = 0 To 2
        r = 4 + k
        ws.Cells(r, 1).Value = kpi(k)
        ws.Cells(r, 2).Formula = CVF(CStr(meas(k)), NAM & ",Slicer_Thang")
    Next
    ws.Range("C4").Formula = "=Cai_dat!$C$5"
    ws.Range("C5").Formula = CVF("Total Target", NAM & ",Slicer_Thang")
    ws.Range("C6").Formula = "=Cai_dat!$C$6"
    For r = 4 To 6
        ws.Cells(r, 4).Formula = "=IFERROR(B" & r & "/C" & r & ","""")"
        ws.Cells(r, 5).Formula = "=IF(D" & r & "="""","""",IF(D" & r & U(">=Cai_dat!$C$10,""\u2714 \u0110\u1EA1t"",IF(D") & r & U(">=Cai_dat!$C$11,""\u25B2 C\u1EA3nh b\u00E1o"",""\u2718 Kh\u00F4ng \u0111\u1EA1t"")))")
    Next
    ws.Range("B4:C6").NumberFormat = "#,##0"
    ws.Range("D4:D6").NumberFormat = "0.0%"
    ws.Range("E4:E6").Font.Name = "Segoe UI Symbol"
    ws.Range("A8").Value = U("S\u1ED1 kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng theo ph\u00E2n kh\u00FAc v\u00E0 doanh thu (Slicer Th\u00E1ng)")
    ws.Range("A8").Font.Bold = True
    ws.Range("A9:D9").Value = Array(U("Ph\u00E2n kh\u00FAc"), U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng"), "Doanh thu", U("Doanh thu b\u00ECnh qu\u00E2n/kh\u00E1ch"))
    Hdr ws.Range("A9:D9")
    ws.Range("A10").Value = U("Th\u01B0\u1EDDng")
    ws.Range("A11").Value = U("Th\u00E2n thi\u1EBFt")
    ws.Range("A12").Value = "VIP"
    For r = 10 To 12
        ws.Cells(r, 2).Formula = "=IFERROR(CUBEVALUE(""" & CN & """,""[Measures].[Active Customers]""" & NAM & ",Slicer_Thang,""[Customer].[Segment].&["" & A" & r & " & ""]""),0)"
        ws.Cells(r, 3).Formula = "=IFERROR(CUBEVALUE(""" & CN & """,""[Measures].[Total Revenue]""" & NAM & ",Slicer_Thang,""[Customer].[Segment].&["" & A" & r & " & ""]""),0)"
        ws.Cells(r, 4).Formula = "=IFERROR(C" & r & "/B" & r & ",0)"
    Next
    ws.Range("A13").Value = U("T\u1ED5ng (\u0111\u1ED1i chi\u1EBFu th\u1EBB KPI)")
    ws.Range("B13").Formula = "=SUM(B10:B12)"
    ws.Range("C13").Formula = "=SUM(C10:C12)"
    ws.Range("A13:D13").Font.Bold = True
    ws.Range("B10:D13").NumberFormat = "#,##0"
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B:E").ColumnWidth = 22
    Chk "data_dash"
    L "data_dash ok"
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

Private Sub MakeDashboardNew()
    Dim ws As Worksheet, wsP As Worksheet, k As Long, c0 As Long, r As Long
    Dim lab, fmtV, shp As Shape, ch As Chart, fc As Object, rg As Range
    On Error Resume Next
    Set ws = GetSheet("Dashboard")
    Set wsP = GetSheet("PV_New")
    ws.Cells.Interior.Color = RGB(243, 245, 249)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:M").ColumnWidth = 10
    ws.Columns("N").ColumnWidth = 2
    ws.Columns("U").ColumnWidth = 2
    ws.Range("B2:M2").Merge
    ws.Range("B2").Value = U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng \u0111ang d\u01B0\u1EDBi target th\u00E1ng 4/2025: c\u1EA7n k\u00EDch ho\u1EA1t th\u00EAm kh\u00E1ch")
    ws.Range("B2").Font.Size = 15
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(255, 255, 255)
    ws.Range("B2:M2").Interior.Color = RGB(32, 56, 100)
    ws.Rows(2).RowHeight = 32
    ws.Range("B3:M3").Merge
    ws.Range("B3").Formula = U("=""Ng\u01B0\u1EDDi xem: Marketing Director \u00B7 Th\u00E1ng \u0111ang ch\u1ECDn: ""&IFERROR(CUBERANKEDMEMBER(""") & CN & U(""",Slicer_Thang,1),""?"")&""/2025 \u00B7 D\u1EEF li\u1EC7u: Data Model""")
    ws.Range("B3").Font.Size = 9
    ws.Range("B3").Font.Color = RGB(88, 88, 88)
    lab = Array(U("KH\u00C1CH HO\u1EA0T \u0110\u1ED8NG"), "DOANH THU", U("GI\u00C1 TR\u1ECA \u0110\u01A0N TB (AOV)"))
    fmtV = Array("#,##0", "#,##0.0,"" tr""", "#,##0")
    For k = 0 To 2
        c0 = 2 + 4 * k
        r = 4 + k
        CardBorder ws.Range(ws.Cells(5, c0), ws.Cells(10, c0 + 3))
        ws.Range(ws.Cells(5, c0), ws.Cells(5, c0 + 3)).Merge
        ws.Cells(5, c0).Value = lab(k)
        ws.Cells(5, c0).Font.Size = 8
        ws.Cells(5, c0).Font.Bold = True
        ws.Cells(5, c0).Font.Color = RGB(88, 88, 88)
        ws.Range(ws.Cells(6, c0), ws.Cells(7, c0 + 3)).Merge
        ws.Cells(6, c0).Formula = "=Data_Dash!B" & r
        ws.Cells(6, c0).NumberFormat = fmtV(k)
        ws.Cells(6, c0).Font.Size = 22
        ws.Cells(6, c0).Font.Bold = True
        ws.Cells(6, c0).Font.Color = RGB(32, 56, 100)
        ws.Range(ws.Cells(8, c0), ws.Cells(8, c0 + 3)).Merge
        If k = 1 Then
            ws.Cells(8, c0).Formula = "=""Target ""&TEXT(Data_Dash!C" & r & "/1000000,""#,##0.0"")&"" tr"""
        Else
            ws.Cells(8, c0).Formula = "=""Target ""&TEXT(Data_Dash!C" & r & ",""#,##0"")"
        End If
        ws.Cells(8, c0).Font.Size = 9
        ws.Cells(8, c0).Font.Color = RGB(88, 88, 88)
        ws.Range(ws.Cells(9, c0), ws.Cells(9, c0 + 3)).Merge
        ws.Cells(9, c0).Formula = "=Data_Dash!D" & r
        ws.Cells(9, c0).NumberFormat = U("""M\u1EE9c \u0111\u1EA1t ""0.0%")
        ws.Cells(9, c0).Font.Bold = True
        ws.Range(ws.Cells(10, c0), ws.Cells(10, c0 + 3)).Merge
        ws.Cells(10, c0).Formula = "=Data_Dash!E" & r
        ws.Cells(10, c0).Font.Name = "Segoe UI Symbol"
        ws.Cells(10, c0).Font.Bold = True
        ws.Range(ws.Cells(5, c0), ws.Cells(10, c0 + 3)).HorizontalAlignment = xlCenter
        Set rg = ws.Cells(9, c0)
        rg.FormatConditions.Delete
        Set fc = rg.FormatConditions.AddIconSetCondition
        fc.IconSet = wb.IconSets(xl3Symbols)
        fc.ShowIconOnly = False
        fc.IconCriteria(2).Type = xlConditionValueFormula
        fc.IconCriteria(2).Value = "=Cai_dat!$C$11"
        fc.IconCriteria(2).Operator = xlGreaterEqual
        fc.IconCriteria(3).Type = xlConditionValueFormula
        fc.IconCriteria(3).Value = "=Cai_dat!$C$10"
        fc.IconCriteria(3).Operator = xlGreaterEqual
        Chk "icon " & k
        Set rg = ws.Cells(10, c0)
        rg.FormatConditions.Delete
        Set fc = rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=Data_Dash!$D$" & r & ">=Cai_dat!$C$10")
        fc.Font.Color = RGB(12, 163, 12)
        Set fc = rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(Data_Dash!$D$" & r & "<Cai_dat!$C$10,Data_Dash!$D$" & r & ">=Cai_dat!$C$11)")
        fc.Font.Color = RGB(190, 125, 0)
        Set fc = rg.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(Data_Dash!$D$" & r & "),Data_Dash!$D$" & r & "<Cai_dat!$C$11)")
        fc.Font.Color = RGB(208, 59, 59)
        Chk "status cf " & k
    Next
    ' ---- 3 chart ----
    Set shp = ws.Shapes.AddChart2(-1, xlBarClustered, ws.Range("B12").Left, ws.Range("B12").Top, ws.Range("B12:E12").Width, 230)
    shp.Name = "chtKhachPK"
    Set ch = shp.Chart
    ch.SetSourceData Source:=wsP.PivotTables("pvtKhachPK").TableRange1
    Chk "chart A"
    ch.HasTitle = True
    ch.ChartTitle.Text = U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng theo ph\u00E2n kh\u00FAc")
    ch.HasLegend = False
    ch.ShowAllFieldButtons = False
    ch.Axes(xlValue).MinimumScale = 0
    ch.SeriesCollection(1).HasDataLabels = True
    shp.AlternativeText = U("Bi\u1EC3u \u0111\u1ED3 thanh ngang: s\u1ED1 kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng theo ph\u00E2n kh\u00FAc Th\u01B0\u1EDDng, Th\u00E2n thi\u1EBFt, VIP trong th\u00E1ng \u0111ang ch\u1ECDn.")
    Chk "chart A fmt"
    Set shp = ws.Shapes.AddChart2(-1, xlBarClustered, ws.Range("F12").Left, ws.Range("B12").Top, ws.Range("F12:I12").Width, 230)
    shp.Name = "chtDoanhThuPK"
    Set ch = shp.Chart
    ch.SetSourceData Source:=wsP.PivotTables("pvtDoanhThuPK").TableRange1
    Chk "chart B"
    ch.HasTitle = True
    ch.ChartTitle.Text = U("Doanh thu theo ph\u00E2n kh\u00FAc")
    ch.HasLegend = False
    ch.ShowAllFieldButtons = False
    ch.Axes(xlValue).MinimumScale = 0
    ch.SeriesCollection(1).HasDataLabels = True
    shp.AlternativeText = U("Bi\u1EC3u \u0111\u1ED3 thanh ngang: doanh thu theo ph\u00E2n kh\u00FAc kh\u00E1ch h\u00E0ng trong th\u00E1ng \u0111ang ch\u1ECDn.")
    Chk "chart B fmt"
    Set shp = ws.Shapes.AddChart2(-1, xlColumnClustered, ws.Range("J12").Left, ws.Range("B12").Top, ws.Range("J12:M12").Width, 230)
    shp.Name = "chtKhachMien"
    Set ch = shp.Chart
    ch.SetSourceData Source:=wsP.PivotTables("pvtKhachMien").TableRange1
    Chk "chart C"
    ch.HasTitle = True
    ch.ChartTitle.Text = U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng theo mi\u1EC1n (kh\u00F4ng c\u1ED9ng l\u1EA1i th\u00E0nh t\u1ED5ng)")
    ch.HasLegend = False
    ch.ShowAllFieldButtons = False
    ch.Axes(xlValue).MinimumScale = 0
    ch.SeriesCollection(1).HasDataLabels = True
    shp.AlternativeText = U("Bi\u1EC3u \u0111\u1ED3 c\u1ED9t: s\u1ED1 kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng \u1EDF t\u1EEBng mi\u1EC1n trong th\u00E1ng \u0111ang ch\u1ECDn. M\u1ED9t kh\u00E1ch mua \u1EDF nhi\u1EC1u mi\u1EC1n \u0111\u01B0\u1EE3c \u0111\u1EBFm \u1EDF m\u1ED7i mi\u1EC1n n\u00EAn kh\u00F4ng c\u1ED9ng th\u00E0nh t\u1ED5ng.")
    Chk "chart C fmt"
    ws.Range("B28").Value = U("Quy t\u1EAFc RAG: \u2714 \u0110\u1EA1t khi m\u1EE9c \u0111\u1EA1t \u2265 100% \u00B7 \u25B2 C\u1EA3nh b\u00E1o t\u1EEB 80% \u0111\u1EBFn d\u01B0\u1EDBi 100% \u00B7 \u2718 Kh\u00F4ng \u0111\u1EA1t khi d\u01B0\u1EDBi 80%. Ng\u01B0\u1EE1ng n\u1EB1m \u1EDF sheet Cai_dat. Slicer Th\u00E1ng \u0111i\u1EC1u khi\u1EC3n th\u1EBB KPI v\u00E0 c\u1EA3 ba bi\u1EC3u \u0111\u1ED3.")
    ws.Range("B28").Font.Size = 8
    ws.Range("B28").Font.Color = RGB(88, 88, 88)
    L "dashboard new ok"
End Sub

Private Sub MakeKiemTra()
    Dim ws As Worksheet, r As Long, i As Long
    Dim lab, f, ex, tol
    On Error Resume Next
    Set ws = GetSheet("Kiem_tra")
    Title ws, U("KI\u1EC2M TRA / \u0110\u1ED0I SO\u00C1T \u2014 B\u00E0i t\u1EADp v\u1EC1 nh\u00E0 B6"), U("C\u1ED9t C: CUBEVALUE v\u1EDBi th\u00E1ng 4/2025 vi\u1EBFt th\u1EB3ng trong c\u00F4ng th\u1EE9c (kh\u00F4ng ph\u1EE5 thu\u1ED9c Slicer). C\u1ED9t D: \u0111\u00E1p \u00E1n.")
    ws.Range("A4:E4").Value = Array(U("Ch\u1EC9 ti\u00EAu"), "", U("K\u1EBFt qu\u1EA3"), U("\u0110\u00E1p \u00E1n"), U("Kh\u1EDBp?"))
    Hdr ws.Range("A4:E4")
    Dim T4 As String
    T4 = NAM & ",""[Calendar].[Thang].&[4]"""
    lab = Array(U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng th\u00E1ng 4/2025"), U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng Th\u01B0\u1EDDng"), U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng Th\u00E2n thi\u1EBFt"), U("Kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng VIP"), U("Doanh thu Th\u01B0\u1EDDng"), U("Doanh thu Th\u00E2n thi\u1EBFt"), "Doanh thu VIP", U("Doanh thu th\u00E1ng 4/2025"), U("M\u1EE9c \u0111\u1EA1t kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng (\u00F7 target 700)"), U("C\u1ED9ng kh\u00E1ch ho\u1EA1t \u0111\u1ED9ng theo ba mi\u1EC1n (\u2260 670)"))
    ex = Array(670, 360, 251, 59, 138190000#, 119096000#, 39679000#, 296965000#, 670 / 700, 713)
    tol = Array(0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.0005, 0.5)
    f = Array( _
        CVF("Active Customers", T4), _
        CVF("Active Customers", T4 & U(",""[Customer].[Segment].&[Th\u01B0\u1EDDng]""")), _
        CVF("Active Customers", T4 & U(",""[Customer].[Segment].&[Th\u00E2n thi\u1EBFt]""")), _
        CVF("Active Customers", T4 & ",""[Customer].[Segment].&[VIP]"""), _
        CVF("Total Revenue", T4 & U(",""[Customer].[Segment].&[Th\u01B0\u1EDDng]""")), _
        CVF("Total Revenue", T4 & U(",""[Customer].[Segment].&[Th\u00E2n thi\u1EBFt]""")), _
        CVF("Total Revenue", T4 & ",""[Customer].[Segment].&[VIP]"""), _
        CVF("Total Revenue", T4), _
        "=C5/Cai_dat!$C$5", _
        "=" & Mid(CVF("Active Customers", T4 & U(",""[Store].[Region].&[B\u1EAFc]""")), 2) & "+" & Mid(CVF("Active Customers", T4 & ",""[Store].[Region].&[Trung]"""), 2) & "+" & Mid(CVF("Active Customers", T4 & ",""[Store].[Region].&[Nam]"""), 2))
    For i = 0 To 9
        r = 5 + i
        ws.Cells(r, 1).Value = lab(i)
        ws.Cells(r, 3).Formula = f(i)
        ws.Cells(r, 4).Value = ex(i)
        ws.Cells(r, 5).Formula = "=IF(ISNUMBER(C" & r & "),IF(ABS(C" & r & "-D" & r & ")<=" & Replace(CStr(tol(i)), ",", ".") & U(",""\u2714"",""\u2718""),""\u2718"")")
    Next
    ws.Range("C5:D14").NumberFormat = "#,##0"
    ws.Range("C13:D13").NumberFormat = "0.0%"
    ws.Range("A3").Formula = U("=""S\u1ED1 \u00F4 kh\u1EDBp: ""&COUNTIF(E5:E14,""\u2714"")&""/""&COUNTA(A5:A14)")
    ws.Range("A3").Font.Bold = True
    ws.Columns("A").ColumnWidth = 50
    ws.Columns("B").ColumnWidth = 2
    ws.Columns("C:D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 9
    ws.Range("E5:E14").HorizontalAlignment = xlCenter
    ws.Range("E5:E14").Font.Name = "Segoe UI Symbol"
    Chk "kiem tra"
    L "kiem_tra ok"
End Sub

Private Sub Finish()
    Dim nm, i As Long, ws As Worksheet
    On Error Resume Next
    If isAns Then
        nm = Array("PV_New", "PV_Cu", "Kiem_tra", "Insight_mau", "Cai_dat", "Data_Dash", "Dashboard", "Nhat_ky_doi_soat", "Brief", "Dashboard_Cu", "De_bai")
    Else
        nm = Array("PV_Cu", "Insight", "Nhat_ky_doi_soat", "Brief", "Dashboard_Cu", "De_bai")
    End If
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
        Select Case ws.Name
            Case "De_bai", "Dashboard_Cu", "Dashboard", "Brief", "Nhat_ky_doi_soat", "Insight", "Insight_mau", "Cai_dat", "Data_Dash", "Kiem_tra"
                ws.Activate
                ActiveWindow.DisplayGridlines = False
        End Select
    Next
    wb.Worksheets("De_bai").Activate
    wb.Worksheets("De_bai").Range("A1").Select
    L "Dashboard_Cu J6=" & wb.Worksheets("Dashboard_Cu").Range("J6").Text & " B6=" & wb.Worksheets("Dashboard_Cu").Range("B6").Text
    If isAns Then
        L "Kiem_tra A3=" & wb.Worksheets("Kiem_tra").Range("A3").Text
        Dim j As Long
        For j = 5 To 14
            L "  KT row " & j & ": C=" & wb.Worksheets("Kiem_tra").Cells(j, 3).Text & " D=" & wb.Worksheets("Kiem_tra").Cells(j, 4).Text
        Next
    End If
    L "finish ok"
End Sub
