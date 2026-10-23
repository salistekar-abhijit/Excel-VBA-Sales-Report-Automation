Attribute VB_Name = "Module1"

Option Explicit

'=========================================================
' PROJECT: SALES REPORT AUTOMATION SYSTEM
'=========================================================

'---------------------------------------------------------
' 1. CREATE SAMPLE SALES DATA
' WARNING: This replaces the sample data in SalesData.
'---------------------------------------------------------
Public Sub SetupSampleSalesData()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("SalesData")

    ws.Range("A1:G6").UnMerge
    ws.Range("A1:G6").Clear

    ws.Range("A1:G1").Value = Array( _
        "OrderID", "Date", "Product", "Region", _
        "Quantity", "UnitPrice", "TotalSales")

    ws.Range("A2:F6").Value = Array( _
        Array(1001, DateSerial(2026, 10, 1), "Laptop", "Mumbai", 2, 45000), _
        Array(1002, DateSerial(2026, 10, 2), "Mouse", "Pune", 5, 500), _
        Array(1003, DateSerial(2026, 10, 3), "Keyboard", "Mumbai", 3, 1200), _
        Array(1004, DateSerial(2026, 10, 4), "Monitor", "Nashik", 2, 12000), _
        Array(1005, DateSerial(2026, 10, 5), "Mouse", "Pune", 4, 500))

    ws.Range("G2:G6").FormulaR1C1 = "=RC[-2]*RC[-1]"

    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Range("A1:G6").Borders.LineStyle = xlContinuous
    ws.Range("B2:B6").NumberFormat = "dd-mmm-yyyy"
    ws.Range("F2:G6").NumberFormat = "#,##0.00"
    ws.Columns("A:G").AutoFit

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:G6").AutoFilter

    ThisWorkbook.Save

    MsgBox "Sample sales data created successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Could not create sample data: " & Err.Description, vbExclamation

End Sub


'---------------------------------------------------------
' 2. GET OR CREATE SUMMARY WORKSHEET
'---------------------------------------------------------
Private Function GetOrCreateSummarySheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("SalesSummary")
    On Error GoTo 0

    If ws Is Nothing Then

        On Error GoTo CreateError

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.Count))

        ws.Name = "SalesSummary"

    End If

    Set GetOrCreateSummarySheet = ws
    Exit Function

CreateError:
    Set GetOrCreateSummarySheet = Nothing

End Function


'---------------------------------------------------------
' 3. GENERATE OVERALL SALES SUMMARY
'---------------------------------------------------------
Public Sub GenerateSalesSummary()

    Dim wsData As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long

    Set wsData = ThisWorkbook.Worksheets("SalesData")

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        Err.Raise vbObjectError + 100, , _
            "No sales records found in SalesData."
    End If

    Set wsSummary = GetOrCreateSummarySheet()

    If wsSummary Is Nothing Then
        Err.Raise vbObjectError + 101, , _
            "Could not create SalesSummary worksheet."
    End If

    wsSummary.Cells.UnMerge
    wsSummary.Cells.Clear

    wsSummary.Range("A1").Value = "SALES SUMMARY REPORT"
    wsSummary.Range("A3").Value = "Metric"
    wsSummary.Range("B3").Value = "Value"

    wsSummary.Range("A4").Value = "Total Revenue"
    wsSummary.Range("B4").Formula = _
        "=SUM(SalesData!G2:G" & lastRow & ")"

    wsSummary.Range("A5").Value = "Total Orders"
    wsSummary.Range("B5").Formula = _
        "=COUNTA(SalesData!A2:A" & lastRow & ")"

    wsSummary.Range("A6").Value = "Total Quantity Sold"
    wsSummary.Range("B6").Formula = _
        "=SUM(SalesData!E2:E" & lastRow & ")"

    wsSummary.Range("A7").Value = "Average Order Value"
    wsSummary.Range("B7").Formula = _
        "=IFERROR(AVERAGE(SalesData!G2:G" & lastRow & "),0)"

    wsSummary.Range("B4").NumberFormat = "#,##0.00"
    wsSummary.Range("B5:B6").NumberFormat = "0"
    wsSummary.Range("B7").NumberFormat = "#,##0.00"

End Sub


'---------------------------------------------------------
' 4. GENERATE REGION-WISE SALES REPORT
'---------------------------------------------------------
Public Sub GenerateRegionReport()

    Dim wsData As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim nextRow As Long
    Dim region As String
    Dim sales As Double
    Dim found As Boolean

    Set wsData = ThisWorkbook.Worksheets("SalesData")
    Set wsSummary = GetOrCreateSummarySheet()

    If wsSummary Is Nothing Then
        Err.Raise vbObjectError + 102, , _
            "Could not access SalesSummary worksheet."
    End If

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        Err.Raise vbObjectError + 103, , _
            "No sales records found in SalesData."
    End If

    wsSummary.Range("D1:E1000").UnMerge
    wsSummary.Range("D1:E1000").Clear

    wsSummary.Range("D1").Value = "REGION-WISE SALES"
    wsSummary.Range("D3").Value = "Region"
    wsSummary.Range("E3").Value = "Total Sales"

    nextRow = 4

    For i = 2 To lastRow

        region = Trim(CStr(wsData.Cells(i, "D").Value))

        If region <> "" And _
           IsNumeric(wsData.Cells(i, "G").Value) And _
           wsData.Cells(i, "G").Value <> "" Then

            sales = CDbl(wsData.Cells(i, "G").Value)
            found = False

            For j = 4 To nextRow - 1

                If StrComp(CStr(wsSummary.Cells(j, "D").Value), _
                           region, vbTextCompare) = 0 Then

                    wsSummary.Cells(j, "E").Value = _
                        CDbl(wsSummary.Cells(j, "E").Value) + sales

                    found = True
                    Exit For

                End If

            Next j

            If Not found Then

                wsSummary.Cells(nextRow, "D").Value = region
                wsSummary.Cells(nextRow, "E").Value = sales

                nextRow = nextRow + 1

            End If

        End If

    Next i

    If nextRow = 4 Then
        Err.Raise vbObjectError + 104, , _
            "No valid region sales records found."
    End If

    wsSummary.Range("E4:E" & nextRow - 1).NumberFormat = "#,##0.00"

End Sub


'---------------------------------------------------------
' 5. FORMAT DASHBOARD AND CREATE CHART
'---------------------------------------------------------
Public Sub FormatSalesDashboard()

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim i As Long
    Dim lastRegionRow As Long

    Set ws = ThisWorkbook.Worksheets("SalesSummary")

    'Delete previous charts to prevent duplicates
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i

    'Set column widths
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 18

    'Main report title
    ws.Range("A1:B1").UnMerge
    ws.Range("A1:B1").Merge

    With ws.Range("A1:B1")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Rows(1).RowHeight = 36

    'Summary headers
    With ws.Range("A3:B3")
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    With ws.Range("A3:B7")
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 220, 230)
    End With

    'Region title
    ws.Range("D1:E1").UnMerge
    ws.Range("D1:E1").Merge

    With ws.Range("D1:E1")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    'Region headers
    With ws.Range("D3:E3")
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    lastRegionRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    If lastRegionRow >= 4 Then

        With ws.Range("D3:E" & lastRegionRow)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(210, 220, 230)
        End With

        'Create region-wise chart
        Set chartObj = ws.ChartObjects.Add( _
            Left:=ws.Range("G3").Left, _
            Top:=ws.Range("G3").Top, _
            Width:=480, _
            Height:=300)

        With chartObj.Chart

            .ChartType = xlColumnClustered
            .SeriesCollection.NewSeries

            With .SeriesCollection(1)
                .Name = "Total Sales"
                .XValues = ws.Range("D4:D" & lastRegionRow)
                .Values = ws.Range("E4:E" & lastRegionRow)
            End With

            .HasTitle = True
            .ChartTitle.Text = "Sales by Region"
            .HasLegend = False
            .ChartStyle = 10

            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Total Sales"

            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Region"

        End With

    End If

    'Restore preferred column widths
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 18

    'Hide gridlines
    ws.Activate
    ActiveWindow.DisplayGridlines = False

End Sub


'---------------------------------------------------------
' 6. MASTER MACRO FOR REFRESH BUTTON
'---------------------------------------------------------
Public Sub RunCompleteSalesReport()

    On Error GoTo ReportError

    Application.ScreenUpdating = False

    GenerateSalesSummary
    GenerateRegionReport
    FormatSalesDashboard

    ThisWorkbook.Worksheets("SalesSummary").Activate

    Application.ScreenUpdating = True

    MsgBox "Sales report and dashboard refreshed successfully!", _
           vbInformation

    Exit Sub

ReportError:

    Application.ScreenUpdating = True

    MsgBox "Report could not be completed." & vbCrLf & _
           "Error: " & Err.Description, vbExclamation

End Sub


'---------------------------------------------------------
' 7. EXPORT DASHBOARD AS PDF
'---------------------------------------------------------
Public Sub ExportSalesReportPDF()

    Dim ws As Worksheet
    Dim filePath As Variant
    Dim oldPrintArea As String
    Dim oldOrientation As XlPageOrientation
    Dim oldZoom As Variant
    Dim oldFitWide As Variant
    Dim oldFitTall As Variant
    Dim oldPaperSize As XlPaperSize
    Dim oldCenterH As Boolean
    Dim oldCenterV As Boolean
    Dim settingsSaved As Boolean

    On Error GoTo ExportError

    Set ws = ThisWorkbook.Worksheets("SalesSummary")

    'Ask where to save the PDF
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="Sales_Report.pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save Sales Report as PDF")

    If VarType(filePath) = vbBoolean Then
        If filePath = False Then Exit Sub
    End If

    If LCase$(Right$(CStr(filePath), 4)) <> ".pdf" Then
        filePath = CStr(filePath) & ".pdf"
    End If

    'Save the existing page settings
    With ws.PageSetup
        oldPrintArea = .printArea
        oldOrientation = .orientation
        oldZoom = .Zoom
        oldFitWide = .FitToPagesWide
        oldFitTall = .FitToPagesTall
        oldPaperSize = .paperSize
        oldCenterH = .CenterHorizontally
        oldCenterV = .CenterVertically
    End With

    settingsSaved = True

    'Set PDF print layout
    With ws.PageSetup
        .printArea = "$A$1:$N$24"
        .orientation = xlLandscape
        .paperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2
        .CenterHorizontally = True
        .CenterVertically = False
    End With

    'Export only the defined print area
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=CStr(filePath), _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    'Restore original Excel page settings
    RestorePageSettings ws, oldPrintArea, oldOrientation, _
        oldZoom, oldFitWide, oldFitTall, oldPaperSize, _
        oldCenterH, oldCenterV

    MsgBox "PDF report exported successfully!", vbInformation
    Exit Sub

ExportError:

    On Error Resume Next

    If settingsSaved Then
        RestorePageSettings ws, oldPrintArea, oldOrientation, _
            oldZoom, oldFitWide, oldFitTall, oldPaperSize, _
            oldCenterH, oldCenterV
    End If

    On Error GoTo 0

    MsgBox "PDF export failed: " & Err.Description, vbExclamation

End Sub


'---------------------------------------------------------
' 8. RESTORE ORIGINAL PAGE SETTINGS
'---------------------------------------------------------
Private Sub RestorePageSettings( _
    ByVal ws As Worksheet, _
    ByVal printArea As String, _
    ByVal orientation As XlPageOrientation, _
    ByVal zoomValue As Variant, _
    ByVal fitWide As Variant, _
    ByVal fitTall As Variant, _
    ByVal paperSize As XlPaperSize, _
    ByVal centerH As Boolean, _
    ByVal centerV As Boolean)

    With ws.PageSetup
        .printArea = printArea
        .orientation = orientation
        .paperSize = paperSize
        .Zoom = zoomValue
        .FitToPagesWide = fitWide
        .FitToPagesTall = fitTall
        .CenterHorizontally = centerH
        .CenterVertically = centerV
    End With

End Sub


