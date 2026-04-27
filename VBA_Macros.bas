
' ----------------------------------------------------------------------------
' MACRO 1: RefreshDashboard
' Trigger: Button on KPI Dashboard sheet
' Action: Recalculates all formulas and shows timestamp message
' ----------------------------------------------------------------------------
Sub RefreshDashboard()
    Application.Calculate
    Application.CalculateFullRebuild
    
    Dim lastUpdated As String
    lastUpdated = Format(Now, "DD-MMM-YYYY HH:MM:SS")
    
    MsgBox "Dashboard refreshed successfully!" & vbCrLf & _
           "Last updated: " & lastUpdated, _
           vbInformation, "Dashboard Refresh"
End Sub

' ----------------------------------------------------------------------------
' MACRO 2: GenerateCityReport
' Trigger: Button on Logistics & Delivery sheet
' Action: Prompts for city name, filters Raw Data, creates new sheet
' ----------------------------------------------------------------------------
Sub GenerateCityReport()
    Dim cityName As String
    cityName = InputBox("Enter City name (e.g., Mumbai, Chennai, Bengaluru):", _
                        "Generate City Report")
    
    If cityName = "" Then
        MsgBox "No city entered. Operation cancelled.", vbExclamation
        Exit Sub
    End If
    
    ' Unfilter first
    On Error Resume Next
    Sheets("Raw Data").ShowAllData
    On Error GoTo 0
    
    ' Filter by city (column J = Customer City)
    Dim wsRaw As Worksheet
    Set wsRaw = Sheets("Raw Data")
    
    ' Check if autofilter exists, if not add it
    If wsRaw.AutoFilterMode = False Then
        wsRaw.Range("A2:AB322").AutoFilter
    End If
    
    ' Apply filter
    wsRaw.Range("A2:AB322").AutoFilter Field:=10, Criteria1:=cityName
    
    ' Count visible rows
    Dim visibleRows As Long
    visibleRows = wsRaw.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    
    If visibleRows <= 0 Then
        MsgBox "No records found for city: " & cityName, vbExclamation
        wsRaw.ShowAllData
        Exit Sub
    End If
    
    ' Create new sheet
    Dim wsNew As Worksheet
    Dim sheetName As String
    sheetName = "City_" & cityName & "_Report"
    
    ' Delete if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsNew = Sheets.Add(After:=Sheets(Sheets.Count))
    wsNew.Name = sheetName
    wsNew.Tab.Color = RGB(255, 194, 0) ' Caterpillar yellow
    
    ' Copy headers
    wsRaw.Rows(2).Copy
    wsNew.Rows(1).PasteSpecial xlPasteValues
    wsNew.Rows(1).PasteSpecial xlPasteFormats
    
    ' Copy filtered data
    Dim lastRow As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    
    Dim copyRange As Range
    Dim destRow As Long
    destRow = 2
    
    Dim i As Long
    For i = 3 To lastRow
        If wsRaw.Rows(i).Hidden = False Then
            wsRaw.Rows(i).Copy
            wsNew.Rows(destRow).PasteSpecial xlPasteValues
            wsNew.Rows(destRow).PasteSpecial xlPasteFormats
            destRow = destRow + 1
        End If
    Next i
    
    ' Auto-fit columns
    wsNew.Cells.EntireColumn.AutoFit
    
    ' Add header info
    wsNew.Rows(1).Insert
    wsNew.Rows(1).Insert
    wsNew.Range("A1").Value = "CITY REPORT: " & UCase(cityName)
    wsNew.Range("A1").Font.Bold = True
    wsNew.Range("A1").Font.Size = 14
    wsNew.Range("A1").Font.Color = RGB(51, 51, 51)
    wsNew.Range("A2").Value = "Generated: " & Format(Now, "DD-MMM-YYYY HH:MM")
    wsNew.Range("A2").Font.Italic = True
    wsNew.Range("A2").Font.Color = RGB(136, 136, 136)
    
    ' Unfilter Raw Data
    wsRaw.ShowAllData
    
    ' Switch to new sheet
    wsNew.Activate
    
    MsgBox "City report created successfully!" & vbCrLf & _
           "City: " & cityName & vbCrLf & _
           "Records found: " & visibleRows & vbCrLf & _
           "Sheet name: " & sheetName, vbInformation
End Sub

' ----------------------------------------------------------------------------
' MACRO 3: ExportKPISummary
' Trigger: Button on KPI Dashboard sheet
' Action: Copies KPI cards and table to a new dated sheet
' ----------------------------------------------------------------------------
Sub ExportKPISummary()
    Dim wsDash As Worksheet
    Dim wsNew As Worksheet
    Dim sheetName As String
    Dim exportDate As String
    
    exportDate = Format(Now, "DDMMYYYY")
    sheetName = "KPI_Export_" & exportDate
    
    Set wsDash = Sheets("KPI Dashboard")
    
    ' Delete if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new sheet
    Set wsNew = Sheets.Add(After:=Sheets(Sheets.Count))
    wsNew.Name = sheetName
    wsNew.Tab.Color = RGB(0, 200, 83) ' Green
    
    ' Copy KPI Cards area (A1:O9 approx)
    wsDash.Range("A1:O9").Copy
    wsNew.Range("A1").PasteSpecial xlPasteValues
    wsNew.Range("A1").PasteSpecial xlPasteFormats
    
    ' Copy Secondary KPI Table (A11:G22 approx)
    wsDash.Range("A11:G22").Copy
    wsNew.Range("A12").PasteSpecial xlPasteValues
    wsNew.Range("A12").PasteSpecial xlPasteFormats
    
    ' Add timestamp header
    wsNew.Rows(1).Insert
    wsNew.Range("A1").Value = "KPI SUMMARY EXPORT — " & Format(Now, "DD-MMM-YYYY")
    wsNew.Range("A1").Font.Bold = True
    wsNew.Range("A1").Font.Size = 16
    wsNew.Range("A1").Font.Color = RGB(51, 51, 51)
    wsNew.Rows(1).Interior.Color = RGB(255, 194, 0)
    
    ' Auto-fit
    wsNew.Cells.EntireColumn.AutoFit
    
    ' Activate new sheet
    wsNew.Activate
    
    MsgBox "KPI Report exported successfully!" & vbCrLf & _
           "Sheet name: " & sheetName, vbInformation
End Sub

' ----------------------------------------------------------------------------
' MACRO 4: HighlightStockoutRisk
' Trigger: Button on Raw Data sheet
' Action: Highlights rows red/amber based on stock vs reorder point
' ----------------------------------------------------------------------------
Sub HighlightStockoutRisk()
    Dim wsRaw As Worksheet
    Set wsRaw = Sheets("Raw Data")
    
    Dim lastRow As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    
    ' Clear existing fill colors in data rows
    wsRaw.Range("A3:AB" & lastRow).Interior.ColorIndex = xlNone
    
    Dim i As Long
    Dim stockOnHand As Long
    Dim reorderPoint As Long
    Dim highRiskCount As Long
    Dim watchCount As Long
    Dim safeCount As Long
    
    highRiskCount = 0
    watchCount = 0
    safeCount = 0
    
    ' this loop took me a while to get right!
    For i = 3 To lastRow
        stockOnHand = wsRaw.Cells(i, 21).Value ' Column U
        reorderPoint = wsRaw.Cells(i, 23).Value ' Column W
        
        If stockOnHand < reorderPoint Then
            ' HIGH RISK — Red
            wsRaw.Range("A" & i & ":AB" & i).Interior.Color = RGB(255, 205, 210)
            highRiskCount = highRiskCount + 1
        ElseIf stockOnHand < 2 * reorderPoint Then
            ' WATCH — Amber
            wsRaw.Range("A" & i & ":AB" & i).Interior.Color = RGB(255, 249, 196)
            watchCount = watchCount + 1
        Else
            ' SAFE — default
            safeCount = safeCount + 1
        End If
    Next i
    
    MsgBox "Stockout Risk Analysis Complete!" & vbCrLf & vbCrLf & _
           "HIGH RISK (Stock < ROP): " & highRiskCount & " products" & vbCrLf & _
           "WATCH (ROP < Stock < 2xROP): " & watchCount & " products" & vbCrLf & _
           "SAFE (Stock >= 2xROP): " & safeCount & " products", _
           vbInformation, "Stockout Risk Summary"
End Sub

' ============================================================================
' END OF MACROS
' ============================================================================
