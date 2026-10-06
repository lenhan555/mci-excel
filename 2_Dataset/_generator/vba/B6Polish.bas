Attribute VB_Name = "B6Polish"
Option Explicit
' Hoan thien file dap an B6: sua dinh dang, khoa Dashboard. Log: vba\_polish_log.txt

Private Sub L(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open "D:\Onedrive\MCI Excel\2_Dataset\_generator\vba\_polish_log.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Public Sub Probe(ByVal wbName As String, ByVal shName As String)
    Dim ws As Worksheet, shp As Shape, sc As SlicerCache, sl As Slicer, t As String
    Set ws = Workbooks(wbName).Worksheets(shName)
    L "== " & wbName & " / " & shName & " protect=" & ws.ProtectContents
    For Each shp In ws.Shapes
        t = ""
        On Error Resume Next
        If shp.HasChart Then t = shp.Chart.ChartTitle.Text
        On Error GoTo 0
        L shp.Name & " | type=" & shp.Type & " | locked=" & shp.Locked & " | chart=" & shp.HasChart & " | " & t & " | " & Round(shp.Left) & "," & Round(shp.Top) & "," & Round(shp.Width) & "," & Round(shp.Height)
    Next shp
    For Each sc In Workbooks(wbName).SlicerCaches
        For Each sl In sc.Slicers
            L "slicer " & sl.Name & " cache=" & sc.Name & " sheet=" & sl.Parent.Name
        Next sl
    Next sc
End Sub

Private Sub Step(ByVal what As String)
    If Err.Number <> 0 Then L "  LOI " & what & ": " & Err.Number & " " & Err.Description: Err.Clear Else L "  ok  " & what
End Sub

Public Sub PolishDash(ByVal wbName As String, ByVal shName As String)
    Dim wb As Workbook, ws As Worksheet, shp As Shape, ch As Chart, sc As SlicerCache, sl As Slicer
    Set wb = Workbooks(wbName)
    Set ws = wb.Worksheets(shName)
    ws.Activate
    If ws.ProtectContents Then ws.Unprotect
    L "== Polish " & wbName
    On Error Resume Next
    Err.Clear
    For Each shp In ws.Shapes
        If shp.HasChart Then
            Set ch = shp.Chart
            If InStr(shp.Name, "Top5") > 0 Then
                ch.Axes(xlValue).TickLabels.NumberFormat = "#,##0,,\ \t\r"
                Step "Top5 truc so tr"
                ch.Axes(xlValue).MajorUnit = 10000000
                Step "Top5 MajorUnit"
            End If
            If InStr(shp.Name, "Waterfall") > 0 Then
                ch.FullSeriesCollection(1).DataLabels.NumberFormat = "#,##0.0,,"
                Step "Waterfall nhan so tr"
                ch.FullSeriesCollection(1).DataLabels.Font.Size = 7
                Step "Waterfall nhan co chu 7"
            End If
        End If
    Next shp
    For Each sc In wb.SlicerCaches
        For Each sl In sc.Slicers
            If sl.Parent.Name = ws.Name Then
                sl.Shape.Locked = False
                Step "Mo khoa slicer " & sl.Name
            End If
        Next sl
    Next sc
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowUsingPivotTables:=True
    Step "Protect " & ws.Name
    L "  protect=" & ws.ProtectContents & " drawing=" & ws.ProtectDrawingObjects
    On Error GoTo 0
End Sub
