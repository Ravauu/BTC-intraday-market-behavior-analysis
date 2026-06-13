Option Explicit

Sub CreateDashboardLayout()

    Dim ws As Worksheet
    Dim wsPivot As Worksheet
    Dim CardMargin As Double
    Dim i As Long
    
    CardMargin = 6
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Dashboard"
    End If
    
    Set wsPivot = ThisWorkbook.Worksheets("PivotSessions")
    
    ws.Cells.Clear
    
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    
    ws.Cells.Interior.Color = RGB(221, 235, 247)
    ws.Columns("A:F").ColumnWidth = 25
    
    AddTitleCard ws, "A1", "F3", "BTC Intraday Market Behavior Analyzer"
    
    Call AddKPILabelCard(ws, "A5", "Total Days", CardMargin)
    Call AddKPILabelCard(ws, "B5", "Total 5m Candles", CardMargin)
    Call AddKPILabelCard(ws, "C5", "AVG Daily Range %", CardMargin)
    Call AddKPILabelCard(ws, "D5", "AVG 5m Range %", CardMargin)
    Call AddKPILabelCard(ws, "E5", "Most Volatile Session", CardMargin)
    Call AddKPILabelCard(ws, "F5", "Highest Volume Session", CardMargin)
    
    Call AddKPIValueCard(ws, "A12", "=ILE.NIEPUSTYCH(DailyData[DateUTC])", "", CardMargin)
    Call AddKPIValueCard(ws, "B12", "=ILE.WIERSZY(BTC_data[DateUTC])", "", CardMargin)
    Call AddKPIValueCard(ws, "C12", "=ŚREDNIA(DailyData[DayRangePercentage])", "0.00%", CardMargin)
    Call AddKPIValueCard(ws, "D12", "=ŚREDNIA(BTC_data[CandleRangePercentage])", "0.00%", CardMargin)
    Call AddKPIValueCard(ws, "E12", "=X.WYSZUKAJ(MAX(PivotSessions!B12:B15);PivotSessions!B12:B15;PivotSessions!A12:A15)", "", CardMargin)
    Call AddKPIValueCard(ws, "F12", "=X.WYSZUKAJ(MAX(PivotSessions!B20:B23);PivotSessions!B20:B23;PivotSessions!A20:A23)", "", CardMargin)
    
    Call AddChartCard(ws, wsPivot, "chAvgSessionRange", "A16", "C31", "A32", "A18", "C29", CardMargin)
    Call AddChartCard(ws, wsPivot, "chTotalSessionVolume", "D16", "F31", "D32", "D18", "F29", CardMargin)
    
    Call AddChartCard(ws, wsPivot, "chAvg5mRangeByHour", "A34", "C49", "A50", "A36", "C47", CardMargin)
    Call AddChartCard(ws, wsPivot, "chTotalVolumeByHour", "D34", "F49", "D50", "D36", "F47", CardMargin)
    
    Call AddChartCard(ws, wsPivot, "chCandleTypeStructure", "A52", "F67", "A68", "A54", "F65", CardMargin)

End Sub


Private Sub AddTitleCard(ws As Worksheet, startCell As String, endCell As String, titleText As String)

    Dim shp As Shape
    
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range(startCell).Left, _
        ws.Range(startCell).Top, _
        ws.Range(endCell).Left + ws.Range(endCell).Width - ws.Range(startCell).Left, _
        ws.Range(endCell).Top + ws.Range(endCell).Height - ws.Range(startCell).Top)
    
    With shp
        .TextFrame2.TextRange.Text = titleText
        .Fill.ForeColor.RGB = RGB(0, 176, 240)
        .Line.Visible = msoFalse
        .Shadow.Visible = msoTrue
        .Shadow.OffsetX = 3
        .Shadow.OffsetY = 3
        .Shadow.Transparency = 0.5
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 28
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

End Sub


Private Sub AddKPILabelCard(ws As Worksheet, cardCell As String, labelText As String, CardMargin As Double)

    Dim shp As Shape
    
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range(cardCell).Left + CardMargin, _
        ws.Range(cardCell).Top, _
        ws.Range(cardCell).Width - (2 * CardMargin), _
        ws.Range(cardCell).Offset(5, 0).Top + ws.Range(cardCell).Offset(5, 0).Height - ws.Range(cardCell).Top)
    
    With shp
        .TextFrame2.TextRange.Text = labelText
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Size = 20
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .Line.Visible = msoFalse
        .Shadow.Visible = msoTrue
        .Shadow.OffsetX = 3
        .Shadow.OffsetY = 3
        .Shadow.Transparency = 0.5
    End With

End Sub


Private Sub AddKPIValueCard(ws As Worksheet, valueCell As String, formulaText As String, numberFormatText As String, CardMargin As Double)

    Dim shp As Shape
    
    ws.Range(valueCell).FormulaLocal = formulaText
    
    If numberFormatText <> "" Then
        ws.Range(valueCell).NumberFormat = numberFormatText
    End If
    
    ws.Range(valueCell).Calculate
    ws.Range(valueCell).Font.Color = RGB(221, 235, 247)
    
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range(valueCell).Left + CardMargin, _
        ws.Range(valueCell).Top, _
        ws.Range(valueCell).Width - (2 * CardMargin), _
        ws.Range(valueCell).Offset(1, 0).Top + ws.Range(valueCell).Offset(1, 0).Height - ws.Range(valueCell).Top)
    
    With shp
        .TextFrame2.TextRange.Text = ws.Range(valueCell).Text
        .Fill.ForeColor.RGB = RGB(112, 48, 160)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Size = 20
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .Line.Visible = msoFalse
        .Shadow.Visible = msoTrue
        .Shadow.OffsetX = 3
        .Shadow.OffsetY = 3
        .Shadow.Transparency = 0.5
    End With

End Sub


Private Sub AddChartCard(ws As Worksheet, wsPivot As Worksheet, chartName As String, _
                         boxStartCell As String, boxEndCell As String, boxBottomCell As String, _
                         chartStartCell As String, chartEndCell As String, CardMargin As Double)

    Dim shpChartBox As Shape
    Dim chSource As ChartObject
    Dim chDash As ChartObject
    
    Set shpChartBox = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range(boxStartCell).Left + CardMargin, _
        ws.Range(boxStartCell).Top, _
        ws.Range(boxEndCell).Left + ws.Range(boxEndCell).Width - ws.Range(boxStartCell).Left - (2 * CardMargin), _
        ws.Range(boxBottomCell).Top - ws.Range(boxStartCell).Top)
    
    With shpChartBox
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoFalse
        .Shadow.Visible = msoTrue
        .Shadow.OffsetX = 3
        .Shadow.OffsetY = 3
        .Shadow.Transparency = 0.5
        .ZOrder msoSendToBack
    End With
    
    Set chSource = wsPivot.ChartObjects(chartName)
    chSource.Copy
    ws.Paste
    
    Set chDash = ws.ChartObjects(ws.ChartObjects.Count)
    
    With chDash
        .Left = ws.Range(chartStartCell).Left + (CardMargin * 2)
        .Top = ws.Range(chartStartCell).Top
        .Width = ws.Range(chartEndCell).Left + ws.Range(chartEndCell).Width - ws.Range(chartStartCell).Left - (4 * CardMargin)
        .Height = ws.Range(chartEndCell).Top + ws.Range(chartEndCell).Height - ws.Range(chartStartCell).Top
        .Border.LineStyle = xlNone
        .Chart.ChartArea.Format.Line.Visible = msoFalse
        .Chart.PlotArea.Format.Line.Visible = msoFalse
    End With

End Sub

