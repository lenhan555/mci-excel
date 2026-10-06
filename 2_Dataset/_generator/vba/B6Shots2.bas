Attribute VB_Name = "B6Shots2"
Option Explicit
' Dung Combo Box (Form Control) + chart doi theo cua hang cho anh H06. Mo ReadOnly, KHONG luu.

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

Public Sub FormCtl()
    Dim s As Worksheet, i As Long, j As Long, shp As Shape, ch As Chart, cb As Object
    Set s = Workbooks("B6_Demo_Chart.xlsx").Worksheets.Add(After:=Workbooks("B6_Demo_Chart.xlsx").Worksheets(1))
    s.Name = "Demo_FC"
    s.Cells.Font.Name = "Arial": s.Cells.Font.Size = 10
    s.Range("A1").Value = U("Doanh thu theo c\u1EEDa h\u00E0ng, T1\u2013T4/2025 (VN\u0110)")
    s.Range("A1").Font.Bold = True: s.Range("A1").Font.Size = 12
    s.Range("A2").Value = U("C\u1EEDa h\u00E0ng")
    For j = 1 To 4: s.Cells(2, 1 + j).Value = "T" & j: Next j
    For i = 1 To 6
        s.Cells(2 + i, 1).Value = U("C\u1EEDa h\u00E0ng ") & Format(i, "00")
        For j = 1 To 4
            s.Cells(2 + i, 1 + j).Value = (60 + i * 7 + j * 9 + ((i * j) Mod 5) * 6) * 1000000
        Next j
    Next i
    s.Range("B3:E8").NumberFormat = "#,##0"
    s.Range("A2:E2").Font.Bold = True
    s.Range("A2:E2").Interior.Color = RGB(31, 78, 121): s.Range("A2:E2").Font.Color = vbWhite
    s.Range("G2").Value = U("C\u1EEDa h\u00E0ng \u0111ang ch\u1ECDn (Cell link)")
    s.Range("G2").Font.Bold = True
    s.Range("H2").Value = 2
    s.Range("G3").Value = U("T\u00EAn c\u1EEDa h\u00E0ng")
    s.Range("H3").Formula = "=INDEX($A$3:$A$8,$H$2)"
    s.Range("A10").Value = U("D\u1EEF li\u1EC7u cho chart (INDEX theo Cell link)")
    s.Range("A10").Font.Bold = True
    For j = 1 To 4
        s.Cells(11, 1 + j).Value = "T" & j
        s.Cells(12, 1 + j).Formula = "=INDEX(" & s.Range(s.Cells(3, 1 + j), s.Cells(8, 1 + j)).Address & ",$H$2)"
    Next j
    s.Range("A12").Formula = "=$H$3"
    s.Range("B12:E12").NumberFormat = "#,##0"
    Set cb = s.DropDowns.Add(s.Range("G5").Left, s.Range("G5").Top, 150, 22)
    cb.ListFillRange = "$A$3:$A$8"
    cb.LinkedCell = "$H$2"
    cb.DropDownLines = 6
    cb.Value = 2
    s.Range("G4").Value = U("Ch\u1ECDn c\u1EEDa h\u00E0ng:")
    s.Range("B11:E12").Select
    Set shp = s.Shapes.AddChart2(-1, xlColumnClustered, s.Range("A14").Left, s.Range("A14").Top, 440, 250)
    Set ch = shp.Chart
    ch.HasLegend = False
    ch.HasTitle = True
    ch.ChartTitle.Text = U("Doanh thu theo th\u00E1ng c\u1EE7a c\u1EEDa h\u00E0ng \u0111ang ch\u1ECDn")
    ch.ChartTitle.Font.Size = 12
    ch.Axes(xlValue).TickLabels.NumberFormat = "#,##0,,\ \t\r"
    s.Activate
    ActiveWindow.DisplayGridlines = False
    s.Range("A1").Select
End Sub
