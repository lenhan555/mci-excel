Attribute VB_Name = "B6Shots"
Option Explicit
' Dung bieu do cho anh minh hoa B6 tren ban sao Demo (mo ReadOnly, KHONG luu). Sinh boi gen_b6_shots.py.
' Chay tu Immediate: B6Shots.Combo, B6Shots.Waterfall, B6Shots.Hist, B6Shots.Scatter, B6Shots.Trend, B6Shots.Fit "Sheet", "A1:N24"

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

Private Function WS(ByVal wbName As String, ByVal shName As String) As Worksheet
    Set WS = Workbooks(wbName).Worksheets(shName)
End Function

Public Sub Fit(ByVal shName As String, ByVal addr As String)
    Dim s As Worksheet
    Set s = ActiveWorkbook.Worksheets(shName)
    s.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    s.Range(addr).Select
    ActiveWindow.Zoom = True
    s.Range("A1").Select
End Sub

Private Sub Fmt(ByVal ch As Chart, ByVal ttl As String)
    ch.HasTitle = True
    ch.ChartTitle.Text = ttl
    ch.ChartTitle.Font.Size = 14
    ch.ChartTitle.Font.Bold = True
End Sub

Public Sub Combo()
    Dim s As Worksheet, shp As Shape, ch As Chart
    Set s = WS("B6_Demo_Chart.xlsx", "Demo_Combo")
    s.Activate
    s.Range("A3:C9").Select
    Set shp = s.Shapes.AddChart2(-1, xlColumnClustered, s.Range("E3").Left, s.Range("E3").Top, 560, 300)
    Set ch = shp.Chart
    ch.SeriesCollection(2).ChartType = xlLineMarkers
    ch.SeriesCollection(2).AxisGroup = 2
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0,,\ \t\r"
    ch.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0%"
    ch.Axes(xlValue, xlSecondary).MinimumScale = 0
    ch.Axes(xlValue, xlSecondary).MaximumScale = 0.5
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    Fmt ch, U("Doanh thu (c\u1ED9t) v\u00E0 Gross Margin % (\u0111\u01B0\u1EDDng) theo ng\u00E0nh, T1\u2013T4/2025")
    Fit "Demo_Combo", "A1:N24"
End Sub

Public Sub Waterfall()
    Dim s As Worksheet, shp As Shape, ch As Chart
    Set s = WS("B6_Demo_Chart.xlsx", "Demo_Waterfall")
    s.Activate
    s.Range("A3:B17").Select
    Set shp = s.Shapes.AddChart2(-1, 119, s.Range("G3").Left, s.Range("G3").Top, 640, 340)
    Set ch = shp.Chart
    ch.HasTitle = True
    ch.ChartTitle.Text = U("T\u1EEB Target \u0111\u1EBFn Actual th\u00E1ng 3/2025, theo c\u1EEDa h\u00E0ng")
    ch.HasLegend = False
    ch.FullSeriesCollection(1).Points(1).IsTotal = True
    ch.FullSeriesCollection(1).Points(14).IsTotal = True
    Fit "Demo_Waterfall", "A1:R24"
End Sub

Public Sub Hist()
    Dim s As Worksheet, shp As Shape, ch As Chart, i As Long
    Set s = WS("B6_Demo_Chart.xlsx", "Demo_Phan_phoi")
    s.Activate
    s.Range("G1").Value = U(U("Kho\u1EA3ng gi\u00E1 tr\u1ECB \u0111\u01A1n (bin 100.000)"))
    s.Range("H1").Value = U(U("S\u1ED1 \u0111\u01A1n"))
    For i = 0 To 24
        s.Cells(2 + i, 7).Value = Format(i * 100, "0") & "-" & Format((i + 1) * 100, "0") & "k"
        s.Cells(2 + i, 8).Formula = "=COUNTIFS($B$2:$B$3145,"">=""&" & i * 100000 & ",$B$2:$B$3145,""<""&" & (i + 1) * 100000 & ")"
    Next i
    s.Range("G1:H26").Select
    Set shp = s.Shapes.AddChart2(-1, xlColumnClustered, s.Range("J1").Left, s.Range("J1").Top, 700, 340)
    Set ch = shp.Chart
    ch.HasLegend = False
    ch.ChartGroups(1).GapWidth = 8
    Fmt ch, U("Ph\u00E2n ph\u1ED1i gi\u00E1 tr\u1ECB \u0111\u01A1n Q1/2025: l\u1EC7ch ph\u1EA3i, \u0111u\u00F4i d\u00E0i")
    Fit "Demo_Phan_phoi", "D1:W26"
End Sub

Public Sub Scatter()
    Dim s As Worksheet, shp As Shape, ch As Chart, tl As Object
    Set s = WS("B6_Demo_Chart.xlsx", "Demo_Scatter")
    s.Activate
    s.Range("B3:C15").Select
    Set shp = s.Shapes.AddChart2(-1, xlXYScatter, s.Range("A21").Left, s.Range("A21").Top, 560, 320)
    Set ch = shp.Chart
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    With ch.SeriesCollection.NewSeries
        .XValues = s.Range("B4:B15")
        .Values = s.Range("C4:C15")
        .Name = U("C\u1EEDa h\u00E0ng")
        Set tl = .Trendlines.Add(Type:=xlLinear, DisplayRSquared:=True)
    End With
    ch.HasLegend = False
    ch.Axes(xlCategory).HasTitle = True
    ch.Axes(xlCategory).AxisTitle.Text = U("Tu\u1ED5i c\u1EEDa h\u00E0ng (n\u0103m)")
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).AxisTitle.Text = U("Doanh thu T1\u2013T4 (VN\u0110)")
    ch.Axes(xlValue).TickLabels.NumberFormat = "#,##0,,\ \t\r"
    Fmt ch, U("Tu\u1ED5i c\u1EEDa h\u00E0ng v\u00E0 doanh thu: c\u00F3 t\u01B0\u01A1ng quan")
    Fit "Demo_Scatter", "A1:N40"
End Sub

Public Sub Trend()
    Dim s As Worksheet, shp As Shape, ch As Chart, tl As Object
    Set s = WS("B6_Demo_DuBao.xlsx", "Demo_Trendline")
    s.Activate
    s.Range("A3:B489").Select
    Set shp = s.Shapes.AddChart2(-1, xlLine, s.Range("D3").Left, s.Range("D3").Top, 720, 340)
    Set ch = shp.Chart
    ch.HasLegend = False
    Set tl = ch.SeriesCollection(1).Trendlines.Add(Type:=xlLinear, Forward:=30, DisplayRSquared:=True)
    ch.Axes(xlValue).TickLabels.NumberFormat = "#,##0,,\ \t\r"
    ch.Axes(xlCategory).TickLabels.NumberFormat = "mm/yyyy"
    Fmt ch, U("Doanh thu theo ng\u00E0y + Linear Trendline (Forward 30 ng\u00E0y)")
    Fit "Demo_Trendline", "A1:Q26"
End Sub
