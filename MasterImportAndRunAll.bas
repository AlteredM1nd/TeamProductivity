' =========================================================================
' --- MAIN MODULE (CONTROL FLOW & REPORTING) ---
' =========================================================================
Option Explicit

'==========================================================================
' --- MASTER SUBROUTINE (with Performance Optimizations) ---
'==========================================================================
Sub Master_ImportAndRunAll()
    Dim startTime As Double: startTime = Timer
    Dim wsOutput As Worksheet
    Dim lastProcessedDate As Date, lastWorkdayDate As Date, loopDate As Date
    Dim missingDates() As Date, missingCount As Long
    Dim importNeeded As Boolean: importNeeded = False
    Dim i As Long
    
    ' *** PERFORMANCE OPTIMIZATION: Disable ALL Excel features that slow down processing ***
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Initializing performance optimizations..."
    
    Dim originalDisplayPageBreaks As Boolean
    originalDisplayPageBreaks = ActiveSheet.DisplayPageBreaks
    ActiveSheet.DisplayPageBreaks = False
    
    ' --- 1. DETERMINE DATE RANGE TO PROCESS ---
    Set wsOutput = ThisWorkbook.Sheets("Output")
    
    If wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).row > 1 Then
        lastProcessedDate = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Value
    Else
        lastProcessedDate = DateSerial(2024, 1, 1) - 1 ' Start before Jan 1, 2024 if empty
    End If
      ' Determine the final target date (yesterday's workday).
    Select Case Weekday(Date, vbMonday)
        Case 1: lastWorkdayDate = Date - 3 ' Monday -> Friday
        Case 7: lastWorkdayDate = Date - 2 ' Sunday -> Friday
        Case Else: lastWorkdayDate = Date - 1 ' Tue-Sat -> Yesterday
    End Select
    
    ' --- 2. COLLECT ALL MISSING DATES ---
    loopDate = lastProcessedDate + 1
    missingCount = 0
    Do While loopDate <= lastWorkdayDate
        If Weekday(loopDate, vbMonday) < 6 Then ' Skip weekends
            If NeedsImport(loopDate) Then
                ReDim Preserve missingDates(missingCount)
                missingDates(missingCount) = loopDate
                missingCount = missingCount + 1
            End If
        End If
        loopDate = loopDate + 1
    Loop
    
    Debug.Print "PERFORMANCE: Missing dates to import: " & missingCount & " at " & Timer
    
    ' --- 3. BULK IMPORT IF NEEDED ---
    If missingCount > 0 Then
        importNeeded = BulkImportDataForDates(missingDates)
        Debug.Print "PERFORMANCE: Bulk import finished at " & Timer
    Else
        Application.StatusBar = "Data is already up to date. Proceeding to calculations."
    End If
    
    ' --- 4. RUN THE FINAL CALCULATIONS ONLY IF IMPORT NEEDED OR OUTPUT IS EMPTY ---
    If importNeeded Or wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).row <= 1 Then
        Debug.Print "PERFORMANCE: Starting calculations at " & Timer
        Call CalculateProductivityMetrics(startTime, missingDates, missingCount)
    Else
        Debug.Print "PERFORMANCE: Skipped calculations, no new data imported at " & Timer
    End If

CleanUp:
    ' *** PERFORMANCE: Restore all Excel settings ***
    ActiveSheet.DisplayPageBreaks = originalDisplayPageBreaks
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' *** PERFORMANCE REPORTING ***
    Dim endTime As Double: endTime = Timer
    Debug.Print "TOTAL EXECUTION TIME: " & Format((endTime - startTime), "0.00") & " seconds"
End Sub

'==========================================================================
' --- HELPER FUNCTION TO IMPORT DATA (Works with Hidden Sheets) ---
'==========================================================================
Private Function ImportDataForDate(ByVal processDate As Date) As Boolean
    Dim sourceURL As String, sourceWB As Workbook, targetWB As Workbook
    Dim processDateStr As String
    Dim ws As Worksheet, parsedDateStr As String
    Dim sourcePersonal As Worksheet, sourceNonEntry As Worksheet
    Dim targetPersonal As Worksheet, targetNonEntry As Worksheet
    Dim templatePersonal As Worksheet, templateNonEntry As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' --- 1. Get SharePoint URL and Template Sheets ---
    Set targetWB = ThisWorkbook
    sourceURL = targetWB.Sheets("Config").Range("Config_SourceWorkbookPath").Value
    
    On Error Resume Next
    Set templatePersonal = targetWB.Sheets("Personal Entry")
    Set templateNonEntry = targetWB.Sheets("Non-Entry Hrs")
    On Error GoTo 0
    If templatePersonal Is Nothing Or templateNonEntry Is Nothing Then
        MsgBox "Required template sheets ('Personal Entry', 'Non-Entry Hrs') were not found.", vbCritical
        Exit Function
    End If
    
    processDateStr = Format(processDate, "yyyy-mm-dd")    ' --- 2. Open Source Workbook and Find Sheets for the given date ---
    ' *** ENHANCED ERROR HANDLING FOR SOURCE WORKBOOK ACCESS ***
    On Error GoTo SourceWorkbookError
    Set sourceWB = Workbooks.Open(sourceURL, ReadOnly:=True, UpdateLinks:=False)
    On Error GoTo ErrorHandler
    
    For Each ws In sourceWB.Worksheets
        If sourcePersonal Is Nothing And ws.name Like "Personal Entry *" Then
            If ParseDateFromName(ws.name, "Personal Entry ") = processDateStr Then Set sourcePersonal = ws
        End If
        If sourceNonEntry Is Nothing And ws.name Like "Non-Entry Hrs *" Then
            If ParseDateFromName(ws.name, "Non-Entry Hrs ") = processDateStr Then Set sourceNonEntry = ws
        End If
    Next ws
    
    If sourcePersonal Is Nothing Or sourceNonEntry Is Nothing Then
        ' *** IMPROVED ERROR HANDLING ***
        Debug.Print "Could not find source sheets for date " & Format(processDate, "M/D/YYYY") & " in the source workbook."
        
        ' Log which specific sheets were missing for better debugging
        If sourcePersonal Is Nothing Then Debug.Print "Missing: Personal Entry " & Format(processDate, "m-d-yy")
        If sourceNonEntry Is Nothing Then Debug.Print "Missing: Non-Entry Hrs " & Format(processDate, "m-d-yy")
        
        ' On Monday, if we're looking for Friday data and it's missing, this is more critical
        If Weekday(Date, vbMonday) = 1 And Weekday(processDate, vbMonday) = 5 Then
            Debug.Print "WARNING: Monday processing failed to find Friday data - this may indicate source system delays"
            ' Still continue with success to avoid blocking the entire process
        End If
        
        GoTo CleanUpAndExit_Success
    End If
    
    ' --- 3. Prepare Target Sheets: Find or Create Them ---
    Dim personalSheetName As String: personalSheetName = sourcePersonal.name
    Dim nonEntrySheetName As String: nonEntrySheetName = sourceNonEntry.name
    
    ' -- Handle Personal Entry Sheet --
    On Error Resume Next
    Set targetPersonal = targetWB.Sheets(personalSheetName)
    On Error GoTo 0
    If targetPersonal Is Nothing Then
        ' *** ROBUST CHANGE: Directly create and assign the new sheet ***
        templatePersonal.Copy After:=targetWB.Sheets(targetWB.Sheets.Count)
        Set targetPersonal = targetWB.Sheets(targetWB.Sheets.Count) ' Get a direct reference
        targetPersonal.name = personalSheetName
    Else
        targetPersonal.Range("C3", targetPersonal.UsedRange.SpecialCells(xlCellTypeLastCell)).ClearContents
    End If
    
    ' -- Handle Non-Entry Sheet --
    On Error Resume Next
    Set targetNonEntry = targetWB.Sheets(nonEntrySheetName)
    On Error GoTo 0
    If targetNonEntry Is Nothing Then
        ' *** ROBUST CHANGE: Directly create and assign the new sheet ***
        templateNonEntry.Copy After:=targetWB.Sheets(targetWB.Sheets.Count)
        Set targetNonEntry = targetWB.Sheets(targetWB.Sheets.Count) ' Get a direct reference
        targetNonEntry.name = nonEntrySheetName
    Else
        targetNonEntry.Range("D2", targetNonEntry.UsedRange.SpecialCells(xlCellTypeLastCell)).ClearContents
    End If

    ' --- 4. Copy Data via "Clean and Paste" Method ---
    Dim dataArray As Variant, r As Long, c As Long
    
    ' -- For Personal Entry --
    Dim lastDataRowPE As Long, lastDataColPE As Long
    
    ' Find the last row based on names in column A of the SOURCE sheet
    lastDataRowPE = sourcePersonal.Cells(sourcePersonal.Rows.Count, "A").End(xlUp).row
    
    ' *** NEW: Find the last column based on the headers in your LOCAL TEMPLATE ***
    lastDataColPE = templatePersonal.Cells(2, templatePersonal.Columns.Count).End(xlToLeft).Column
    
    If lastDataRowPE >= 3 And lastDataColPE >= 3 Then ' Ensure there is data to copy
        ' Define the source data range using the dimensions we found
        Dim sourceDataRangePE As Range
        Set sourceDataRangePE = sourcePersonal.Range(sourcePersonal.Cells(3, 3), sourcePersonal.Cells(lastDataRowPE, lastDataColPE))
        
        ' Step 1: Copy to memory
        dataArray = sourceDataRangePE.Value2
        
        ' Step 2: Clean in memory
        For r = 1 To UBound(dataArray, 1)
            For c = 1 To UBound(dataArray, 2)
                If IsError(dataArray(r, c)) Then dataArray(r, c) = "" ' Replace errors with blanks
            Next c
        Next r
        
        ' Step 3: Paste from memory
        targetPersonal.Range("C3").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).Value = dataArray
    End If
    
    ' -- For Non-Entry Hrs --
    Dim lastDataRowNE As Long, lastDataColNE As Long
    lastDataRowNE = sourceNonEntry.Cells(sourceNonEntry.Rows.Count, "A").End(xlUp).row
    lastDataColNE = sourceNonEntry.Cells(1, sourceNonEntry.Columns.Count).End(xlToLeft).Column
    If lastDataRowNE >= 2 And lastDataColNE >= 4 Then
        dataArray = sourceNonEntry.Range(sourceNonEntry.Cells(2, 4), sourceNonEntry.Cells(lastDataRowNE, lastDataColNE)).Value2
        For r = 1 To UBound(dataArray, 1)
            For c = 1 To UBound(dataArray, 2)
                If IsError(dataArray(r, c)) Then dataArray(r, c) = ""
            Next c
        Next r
        targetNonEntry.Range("D2").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).Value = dataArray
    End If
    
CleanUpAndExit_Success:
    ImportDataForDate = True ' Signal success to the master loop

CleanUpAndExit_Fail:
    If Not sourceWB Is Nothing Then sourceWB.Close SaveChanges:=False
    Erase dataArray
    Exit Function

SourceWorkbookError:
    ' *** SPECIFIC ERROR HANDLING FOR SOURCE WORKBOOK ACCESS ISSUES ***
    Debug.Print "ERROR: Could not open source workbook: " & sourceURL
    Debug.Print "Error Description: " & Err.Description
    
    ' Log additional details for Monday-Friday scenarios
    If Weekday(Date, vbMonday) = 1 And Weekday(processDate, vbMonday) = 5 Then
        Debug.Print "WARNING: Monday attempting to access source workbook for Friday data"
        Debug.Print "This may indicate SharePoint sync delays over the weekend"
    End If
    
    ImportDataForDate = False
    Resume CleanUpAndExit_Fail

ErrorHandler:
    Debug.Print "ERROR in ImportDataForDate: " & Err.Description & " (Error " & Err.Number & ")"
    
    ' Enhanced error reporting for Monday-Friday scenarios
    If Weekday(Date, vbMonday) = 1 And Weekday(processDate, vbMonday) = 5 Then
        Debug.Print "CONTEXT: Monday processing Friday data - error may be related to weekend data availability"
    End If
    
    ' Show detailed error for unexpected issues
    MsgBox "An unexpected error occurred while importing data for " & Format(processDate, "M/D/YYYY") & "." & vbNewLine & vbNewLine & _
           "Error: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number, vbCritical, "Import Error"
    
    ImportDataForDate = False ' Signal failure to the master loop
    Resume CleanUpAndExit_Fail
End Function


'==========================================================================
' --- MAIN CALCULATION SUBROUTINE (Rebuilds from 2024 Onwards) ---
'==========================================================================
Private Sub CalculateProductivityMetrics(ByVal startTime As Double, Optional ByRef datesToProcess() As Date = Nothing, Optional ByVal dateCount As Long = 0)
    ' *** PERFORMANCE: Start metrics calculation timing ***
    Dim metricsStartTime As Double: metricsStartTime = Timer
    Application.StatusBar = "Calculating productivity metrics..."
    
    ' --- ALL VARIABLES ---
    ' (Variable list remains the same)
    Dim wsOutput As Worksheet, wsOutputNE As Worksheet, wsDashboard As Worksheet, wsMonthlyBreakdown As Worksheet, wsWeeklyBreakdown As Worksheet, wsDailyBreakdown As Worksheet
    Dim dailyHoursDict As Object, personDaySickAwayHoursDict As Object, personMonthlyData As Object, personWeeklyData As Object, allTeamMembersMasterDict As Object
    Dim dashboardMonthlyAggregator As Object, allActivityDays As Object, personMonthlyAdjWorkdaySum As Object, personWeeklyAdjWorkdaySum As Object
    Dim arrOutput As Variant, arrOutputNE As Variant, weeklyOutputArray As Variant, monthlyOutputArray As Variant, dailyOutputArray As Variant
    Dim lastRowOutput As Long, lastRowOutputNE As Long, rowIdx As Long, monthRow As Long, weeklyRowCount As Long, dailyRowCount As Long, monthlyRowCount As Long
    Dim key As Variant, personName As String, workDate As Date, entryType As String, dailyHours As Double, overallStartDate As Date, overallEndDate As Date
    Dim monthKey As String, personMonthKey As String, personWeekKey As String, personDayKey As String, weekStartDate As Date, weekEndDate As Date, weekStartDateStr As String
    Dim endTime As Double, execTime As String, k_variant As Variant, parts As Variant, sortMap As Object, sortKey As String, originalKey As String, sortKeys() As String, sortIdx As Long
    Dim actualWorkDays As Long, adjustedWorkDays As Double, totalProdHrsPersonMonth As Double, avgDailyPerson As Double, activeMMCount As Long, totalHrs As Double, totalAdjDays As Double, membersMetTarget As Long, metTargetPercent As Double, prodEligibleCount As Long, metTargetFlag As Boolean
    Dim W_totalProdHrs As Double, W_actualWDays As Long, W_totalSAHrs As Double, W_equivSADays As Double, W_adjWDays As Double, W_avgDaily As Double, prodValue_weekly As Double
    Dim personNamePart As String, monthPart As String, totalProdHrs As Double, totalSAHrs As Double, adjWDays As Double, avgDaily As Double, equivSADays As Double, prodValue_monthly As Double
    Dim D_ProdHrs As Double, D_SAHrs As Double, D_AdjWorkdayFactor As Double, D_EffectiveTarget As Double, D_Productivity As Double, D_WorkDate As Date, D_MonthContext As String, D_PersonName As String, prodValue_daily As Double
    Dim maxDataRow As Long, prodHrsDay As Double, sickAwayHrsDay As Double, dailyAdjFactor As Double

    ' --- CONFIGURATION ---
    Dim wsConfig As Worksheet, DAILY_TARGET_HOURS As Double, HOURS_PER_SICK_AWAY_DAY As Double
    Dim sickAwayCategories As Object, categoryRange As Range, cell As Range
    Dim nonProdTasks As Object, nonProdRange As Range
    
    '--- STEP 1: LOAD CONFIGURATION ---
    Set wsConfig = ThisWorkbook.Sheets("Config")
    DAILY_TARGET_HOURS = wsConfig.Range("Config_DailyTargetHours").Value
    HOURS_PER_SICK_AWAY_DAY = wsConfig.Range("Config_HoursPerSickDay").Value
    Set categoryRange = wsConfig.Range("Config_SickAwayCategories")
    Set sickAwayCategories = CreateObject("Scripting.Dictionary")
    sickAwayCategories.CompareMode = vbTextCompare
    For Each cell In categoryRange.Cells
        If Not IsEmpty(cell.Value) Then sickAwayCategories(CStr(cell.Value)) = 1
    Next cell
    Set nonProdRange = wsConfig.Range("Config_NonProductiveTasks")
    Set nonProdTasks = CreateObject("Scripting.Dictionary")
    nonProdTasks.CompareMode = vbTextCompare
    For Each cell In nonProdRange.Cells
        If Not IsEmpty(cell.Value) Then nonProdTasks(CStr(cell.Value)) = 1
    Next cell

    '--- STEP 2: REBUILD THE OUTPUT SHEETS FROM ALL RELEVANT DATED SHEETS ---
    Application.StatusBar = "Step 3: Rebuilding Output sheets from all dated sources for the year..."
    Set wsOutput = ThisWorkbook.Sheets("Output")
    Set wsOutputNE = ThisWorkbook.Sheets("OutputNE")
    If Not IsMissing(datesToProcess) And Not IsEmpty(datesToProcess) And dateCount > 0 Then
        ' Only process/append for the provided dates
        Dim processDatesDict As Object
        Set processDatesDict = CreateObject("Scripting.Dictionary")
        Dim i As Long
        For i = 0 To dateCount - 1
            processDatesDict(Format(datesToProcess(i), "yyyy-mm-dd")) = 1
        Next i
        ' Do NOT clear the output sheets, just append/update for these dates
    Else
        wsOutput.Cells.Clear
        wsOutputNE.Cells.Clear
    End If
    Dim localSheet As Worksheet, parsedDate As String, sheetDate As Date
    Dim reportStartDate As Date: reportStartDate = DateSerial(2024, 1, 1)
    For Each localSheet In ThisWorkbook.Worksheets
        If localSheet.name Like "Personal Entry *" Then
            If localSheet.name <> "Personal Entry" Then
                parsedDate = ParseDateFromName(localSheet.name, "Personal Entry ")
                If parsedDate <> "" Then
                    sheetDate = CDate(parsedDate)
                    If sheetDate >= reportStartDate Then
                        If (IsMissing(datesToProcess) Or IsEmpty(datesToProcess) Or dateCount = 0) Or processDatesDict.Exists(parsedDate) Then
                            Call ProcessActivitySheet(localSheet, parsedDate)
                        End If
                    End If
                End If
            End If
        ElseIf localSheet.name Like "Non-Entry Hrs *" Then
            If localSheet.name <> "Non-Entry Hrs" Then
                parsedDate = ParseDateFromName(localSheet.name, "Non-Entry Hrs ")
                If parsedDate <> "" Then
                    sheetDate = CDate(parsedDate)
                    If sheetDate >= reportStartDate Then
                        If (IsMissing(datesToProcess) Or IsEmpty(datesToProcess) Or dateCount = 0) Or processDatesDict.Exists(parsedDate) Then
                            Call ProcessNonEntrySheet(localSheet, parsedDate)
                        End If
                    End If
                End If
            End If
        End If
    Next localSheet

    '--- STEP 3: READ & AGGREGATE ALL DATA FROM THE NEWLY BUILT OUTPUT SHEETS ---
    Application.StatusBar = "Step 4: Aggregating all processed data..."
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).row
    lastRowOutputNE = wsOutputNE.Cells(wsOutputNE.Rows.Count, "A").End(xlUp).row
    If lastRowOutput <= 1 And lastRowOutputNE <= 1 Then GoTo NoDataToProcess
    
    If lastRowOutput > 1 Then arrOutput = wsOutput.Range("A2:G" & lastRowOutput).Value
    If lastRowOutputNE > 1 Then arrOutputNE = wsOutputNE.Range("A2:D" & lastRowOutputNE).Value
    
    Set dailyHoursDict = CreateObject("Scripting.Dictionary")
    Set personDaySickAwayHoursDict = CreateObject("Scripting.Dictionary")
    Set personMonthlyData = CreateObject("Scripting.Dictionary")
    Set personWeeklyData = CreateObject("Scripting.Dictionary")
    Set allTeamMembersMasterDict = CreateObject("Scripting.Dictionary"): allTeamMembersMasterDict.CompareMode = vbTextCompare
    Set personMonthlyAdjWorkdaySum = CreateObject("Scripting.Dictionary")
    Set personWeeklyAdjWorkdaySum = CreateObject("Scripting.Dictionary")
    overallStartDate = DateSerial(Year(Date) + 10, 1, 1): overallEndDate = DateSerial(1900, 1, 1)

    If Not IsEmpty(arrOutput) Then
        For rowIdx = 1 To UBound(arrOutput, 1)
            If Not IsEmpty(arrOutput(rowIdx, 1)) And Not IsEmpty(arrOutput(rowIdx, 2)) And IsDate(arrOutput(rowIdx, 1)) Then
                personName = CStr(arrOutput(rowIdx, 2)): workDate = CDate(arrOutput(rowIdx, 1))
                If IsNumeric(arrOutput(rowIdx, 7)) Then
                    dailyHours = CDbl(arrOutput(rowIdx, 7))
                Else
                    dailyHours = 0
                End If
                If dailyHours <> 0 Then
                    personDayKey = personName & "|" & Format(workDate, "yyyy-mm-dd")
                    dailyHoursDict(personDayKey) = dailyHoursDict(personDayKey) + dailyHours
                    If Not allTeamMembersMasterDict.Exists(personName) Then allTeamMembersMasterDict.Add personName, 1
                    If workDate < overallStartDate Then overallStartDate = workDate
                    If workDate > overallEndDate Then overallEndDate = workDate
                End If
            End If
        Next rowIdx
    End If
    If Not IsEmpty(arrOutputNE) Then
        For rowIdx = 1 To UBound(arrOutputNE, 1)
            If Not IsEmpty(arrOutputNE(rowIdx, 1)) And Not IsEmpty(arrOutputNE(rowIdx, 2)) And IsDate(arrOutputNE(rowIdx, 1)) Then
                personName = CStr(arrOutputNE(rowIdx, 2)): workDate = CDate(arrOutputNE(rowIdx, 1))
                entryType = CStr(arrOutputNE(rowIdx, 3))
                
                ' The key change: Check if the task should be excluded
                If Not nonProdTasks.Exists(entryType) Then
                    dailyHours = IIf(IsNumeric(arrOutputNE(rowIdx, 4)), CDbl(arrOutputNE(rowIdx, 4)), 0)
                    personDayKey = personName & "|" & Format(workDate, "yyyy-mm-dd")
                    If sickAwayCategories.Exists(entryType) Then
                        personDaySickAwayHoursDict(personDayKey) = personDaySickAwayHoursDict(personDayKey) + dailyHours
                    ElseIf dailyHours <> 0 Then
                        dailyHoursDict(personDayKey) = dailyHoursDict(personDayKey) + dailyHours
                    End If
                    If dailyHours <> 0 Or sickAwayCategories.Exists(entryType) Then
                        If Not allTeamMembersMasterDict.Exists(personName) Then allTeamMembersMasterDict.Add personName, 1
                        If workDate < overallStartDate Then overallStartDate = workDate
                        If workDate > overallEndDate Then overallEndDate = workDate
                    End If
                End If ' End of the non-productive task check
            End If
        Next rowIdx
    End If
    If allTeamMembersMasterDict.Count = 0 Then GoTo NoDataToProcess
    If overallStartDate > overallEndDate And dailyHoursDict.Count = 0 And personDaySickAwayHoursDict.Count = 0 Then GoTo NoDataToProcess

    Set allActivityDays = CreateObject("Scripting.Dictionary")
    For Each key In dailyHoursDict.Keys: allActivityDays(key) = 1: Next key
    For Each key In personDaySickAwayHoursDict.Keys: allActivityDays(key) = 1: Next key
    
    For Each key In allActivityDays.Keys
        parts = Split(CStr(key), "|"): personName = parts(0): workDate = CDate(parts(1))
        If dailyHoursDict.Exists(key) Then prodHrsDay = dailyHoursDict(key) Else prodHrsDay = 0
        If personDaySickAwayHoursDict.Exists(key) Then sickAwayHrsDay = personDaySickAwayHoursDict(key) Else sickAwayHrsDay = 0
        monthKey = Format(workDate, "yyyy-mm")
        weekStartDateStr = CStr(workDate - Weekday(workDate, vbSunday) + 1)
        personMonthKey = personName & "|" & monthKey
        personWeekKey = personName & "|" & weekStartDateStr
        If HOURS_PER_SICK_AWAY_DAY > 0 Then dailyAdjFactor = 1 - (sickAwayHrsDay / HOURS_PER_SICK_AWAY_DAY) Else dailyAdjFactor = 1
        If dailyAdjFactor < 0 Then dailyAdjFactor = 0: If dailyAdjFactor > 1 Then dailyAdjFactor = 1
        personMonthlyAdjWorkdaySum(personMonthKey) = personMonthlyAdjWorkdaySum(personMonthKey) + dailyAdjFactor
        personWeeklyAdjWorkdaySum(personWeekKey) = personWeeklyAdjWorkdaySum(personWeekKey) + dailyAdjFactor
        If Not personMonthlyData.Exists(personMonthKey) Then
            Set personMonthlyData(personMonthKey) = CreateObject("Scripting.Dictionary")
            personMonthlyData(personMonthKey)("TotalProdHrs") = 0
            Set personMonthlyData(personMonthKey)("ActualWorkDaysDict") = CreateObject("Scripting.Dictionary")
            personMonthlyData(personMonthKey)("TotalSickAwayHours") = 0
        End If
        If Not personWeeklyData.Exists(personWeekKey) Then
            Set personWeeklyData(personWeekKey) = CreateObject("Scripting.Dictionary")
            personWeeklyData(personWeekKey)("TotalProdHrsWeek") = 0
            Set personWeeklyData(personWeekKey)("ActualWorkDaysWeekDict") = CreateObject("Scripting.Dictionary")
            personWeeklyData(personWeekKey)("TotalSickAwayHoursWeek") = 0
        End If
        personMonthlyData(personMonthKey)("TotalProdHrs") = personMonthlyData(personMonthKey)("TotalProdHrs") + prodHrsDay
        personWeeklyData(personWeekKey)("TotalProdHrsWeek") = personWeeklyData(personWeekKey)("TotalProdHrsWeek") + prodHrsDay
        personMonthlyData(personMonthKey)("TotalSickAwayHours") = personMonthlyData(personMonthKey)("TotalSickAwayHours") + sickAwayHrsDay
        personWeeklyData(personWeekKey)("TotalSickAwayHoursWeek") = personWeeklyData(personWeekKey)("TotalSickAwayHoursWeek") + sickAwayHrsDay
        personMonthlyData(personMonthKey)("ActualWorkDaysDict")(Format(workDate, "yyyy-mm-dd")) = 1
        personWeeklyData(personWeekKey)("ActualWorkDaysWeekDict")(Format(workDate, "yyyy-mm-dd")) = 1
    Next key
    
    ' --- STEP 4: GENERATE REPORTS ---
    Application.StatusBar = "Step 5: Generating report sheets..."
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsDashboard = ThisWorkbook.Sheets("ProductivityDashboard"): If wsDashboard Is Nothing Then Set wsDashboard = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)): wsDashboard.name = "ProductivityDashboard" Else wsDashboard.Cells.Clear
    Set wsWeeklyBreakdown = ThisWorkbook.Sheets("WeeklyBreakdown"): If wsWeeklyBreakdown Is Nothing Then Set wsWeeklyBreakdown = ThisWorkbook.Sheets.Add(After:=wsDashboard): wsWeeklyBreakdown.name = "WeeklyBreakdown" Else wsWeeklyBreakdown.Cells.Clear
    Set wsMonthlyBreakdown = ThisWorkbook.Sheets("MonthlyBreakdown"): If wsMonthlyBreakdown Is Nothing Then Set wsMonthlyBreakdown = ThisWorkbook.Sheets.Add(After:=wsWeeklyBreakdown): wsMonthlyBreakdown.name = "MonthlyBreakdown" Else wsMonthlyBreakdown.Cells.Clear
    Set wsDailyBreakdown = ThisWorkbook.Sheets("DailyBreakdown"): If wsDailyBreakdown Is Nothing Then Set wsDailyBreakdown = ThisWorkbook.Sheets.Add(After:=wsMonthlyBreakdown): wsDailyBreakdown.name = "DailyBreakdown" Else wsDailyBreakdown.Cells.Clear
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' -- Dashboard --
    With wsDashboard
        .Range("A1").Value = "Productivity Dashboard": .Range("A1").Font.Bold = True: .Range("A1").Font.Size = 14
        .Range("A3:H3").Value = Array("Month", "Active Team Members", "Total Productive Hours", "Total Adjusted Workdays", "Target Avg Prod. Hrs/Day", "Productive Members", "Members Meeting Target", "Met Target %")
        .Range("A3:H3").Interior.Color = RGB(200, 220, 255): .Range("A3:H3").Font.Bold = True
    End With
    monthRow = 4
    Set dashboardMonthlyAggregator = CreateObject("Scripting.Dictionary")
    For Each key In personMonthlyData.Keys
        parts = Split(CStr(key), "|"): personName = parts(0): monthKey = parts(1)
        If personMonthlyData(key)("TotalProdHrs") > 0 Or personMonthlyData(key)("TotalSickAwayHours") > 0 Then
            If Not dashboardMonthlyAggregator.Exists(monthKey) Then
                Set dashboardMonthlyAggregator(monthKey) = CreateObject("Scripting.Dictionary")
                dashboardMonthlyAggregator(monthKey)("ActiveMembersCount") = 0: dashboardMonthlyAggregator(monthKey)("ProdEligibleMembersCount") = 0
                dashboardMonthlyAggregator(monthKey)("TotalProdHrsSum") = 0: dashboardMonthlyAggregator(monthKey)("TotalAdjWorkdaysSum") = 0
                Set dashboardMonthlyAggregator(monthKey)("ActiveMembersDict") = CreateObject("Scripting.Dictionary"): dashboardMonthlyAggregator(monthKey)("MembersMeetingTargetCount") = 0
            End If
            totalProdHrsPersonMonth = personMonthlyData(key)("TotalProdHrs"): adjustedWorkDays = personMonthlyAdjWorkdaySum(key)
            metTargetFlag = False
            If adjustedWorkDays > 0 Then
                avgDailyPerson = totalProdHrsPersonMonth / adjustedWorkDays
                If avgDailyPerson >= DAILY_TARGET_HOURS Then metTargetFlag = True
            ElseIf totalProdHrsPersonMonth > 0 Then metTargetFlag = True
            End If
            If adjustedWorkDays < 0 Then adjustedWorkDays = 0
            If Not dashboardMonthlyAggregator(monthKey)("ActiveMembersDict").Exists(personName) Then
                 dashboardMonthlyAggregator(monthKey)("ActiveMembersDict")(personName) = 1
                 dashboardMonthlyAggregator(monthKey)("ActiveMembersCount") = dashboardMonthlyAggregator(monthKey)("ActiveMembersCount") + 1
                 If totalProdHrsPersonMonth > 0 Then dashboardMonthlyAggregator(monthKey)("ProdEligibleMembersCount") = dashboardMonthlyAggregator(monthKey)("ProdEligibleMembersCount") + 1
            End If
            dashboardMonthlyAggregator(monthKey)("TotalProdHrsSum") = dashboardMonthlyAggregator(monthKey)("TotalProdHrsSum") + totalProdHrsPersonMonth
            dashboardMonthlyAggregator(monthKey)("TotalAdjWorkdaysSum") = dashboardMonthlyAggregator(monthKey)("TotalAdjWorkdaysSum") + adjustedWorkDays
            If metTargetFlag Then dashboardMonthlyAggregator(monthKey)("MembersMeetingTargetCount") = dashboardMonthlyAggregator(monthKey)("MembersMeetingTargetCount") + 1
        End If
    Next key
    If dashboardMonthlyAggregator.Count > 0 Then
        ReDim sortKeys(0 To dashboardMonthlyAggregator.Count - 1): sortIdx = 0
        For Each k_variant In dashboardMonthlyAggregator.Keys: sortKeys(sortIdx) = CStr(k_variant): sortIdx = sortIdx + 1: Next k_variant
        SortArrayInPlace sortKeys
        For sortIdx = 0 To UBound(sortKeys)
            monthKey = sortKeys(sortIdx)
            activeMMCount = dashboardMonthlyAggregator(monthKey)("ActiveMembersCount"): prodEligibleCount = dashboardMonthlyAggregator(monthKey)("ProdEligibleMembersCount")
            totalHrs = dashboardMonthlyAggregator(monthKey)("TotalProdHrsSum"): totalAdjDays = dashboardMonthlyAggregator(monthKey)("TotalAdjWorkdaysSum")
            membersMetTarget = dashboardMonthlyAggregator(monthKey)("MembersMeetingTargetCount")
            If prodEligibleCount > 0 Then metTargetPercent = membersMetTarget / prodEligibleCount Else metTargetPercent = 0
            With wsDashboard
                .Cells(monthRow, 1).Value = Format(CDate(monthKey & "-01"), "yyyy-mmm"): .Cells(monthRow, 2).Value = activeMMCount
                .Cells(monthRow, 3).Value = Round(totalHrs, 2): .Cells(monthRow, 4).Value = Round(totalAdjDays, 2)
                .Cells(monthRow, 5).Value = DAILY_TARGET_HOURS: .Cells(monthRow, 6).Value = prodEligibleCount
                .Cells(monthRow, 7).Value = membersMetTarget: .Cells(monthRow, 8).Value = metTargetPercent
                .Cells(monthRow, 8).NumberFormat = "0.00%"
            End With
            monthRow = monthRow + 1
        Next sortIdx
    End If
    
    ' -- Weekly Breakdown --
    With wsWeeklyBreakdown
        .Range("A1").Value = "Weekly Productivity Breakdown": .Range("A1").Font.Bold = True: .Range("A1").Font.Size = 14
        .Range("A3:K3").Value = Array("Week Start", "Week End", "Team Member", "Total Prod. Hrs (Week)", "Actual Workdays (Week)", "Sick/Away Hrs (Week)", "Equiv. Sick/Away Days", "Adjusted Workdays (Week)", "Avg Prod. Hrs/Adj.Day (Week)", "Target", "Productivity % (Week)")
        .Range("A3:K3").Font.Bold = True: .Range("A3:K3").Interior.Color = RGB(220, 230, 241)
        If personWeeklyData.Count > 0 Then
            ReDim weeklyOutputArray(1 To personWeeklyData.Count, 1 To 11): weeklyRowCount = 0
            Set sortMap = CreateObject("Scripting.Dictionary"): ReDim sortKeys(0 To personWeeklyData.Count - 1): sortIdx = 0
            For Each key In personWeeklyData.Keys
                parts = Split(CStr(key), "|"): personName = parts(0): weekStartDateStr = parts(1)
                sortKey = Format(CDate(weekStartDateStr), "YYYYMMDD") & "|" & personName
                sortKeys(sortIdx) = sortKey: sortMap(sortKey) = CStr(key): sortIdx = sortIdx + 1
            Next key
            SortArrayInPlace sortKeys
            For sortIdx = 0 To UBound(sortKeys)
                originalKey = sortMap(sortKeys(sortIdx))
                parts = Split(originalKey, "|"): personName = parts(0): weekStartDateStr = parts(1)
                weekStartDate = CDate(weekStartDateStr): weekEndDate = weekStartDate + 6
                W_totalProdHrs = personWeeklyData(originalKey)("TotalProdHrsWeek"): W_actualWDays = personWeeklyData(originalKey)("ActualWorkDaysWeekDict").Count
                W_totalSAHrs = personWeeklyData(originalKey)("TotalSickAwayHoursWeek")
                If HOURS_PER_SICK_AWAY_DAY > 0 Then W_equivSADays = W_totalSAHrs / HOURS_PER_SICK_AWAY_DAY Else W_equivSADays = 0
                W_adjWDays = personWeeklyAdjWorkdaySum(originalKey)
                If W_adjWDays < 0 Then W_adjWDays = 0
                If W_adjWDays > 0 Then W_avgDaily = W_totalProdHrs / W_adjWDays Else W_avgDaily = 0
                If W_totalProdHrs > 0 Or W_adjWDays > 0 Or W_actualWDays > 0 Then
                    weeklyRowCount = weeklyRowCount + 1
                    weeklyOutputArray(weeklyRowCount, 1) = weekStartDate: weeklyOutputArray(weeklyRowCount, 2) = weekEndDate
                    weeklyOutputArray(weeklyRowCount, 3) = personName: weeklyOutputArray(weeklyRowCount, 4) = Round(W_totalProdHrs, 2)
                    weeklyOutputArray(weeklyRowCount, 5) = W_actualWDays: weeklyOutputArray(weeklyRowCount, 6) = Round(W_totalSAHrs, 2)
                    weeklyOutputArray(weeklyRowCount, 7) = Round(W_equivSADays, 2): weeklyOutputArray(weeklyRowCount, 8) = Round(W_adjWDays, 2)
                    weeklyOutputArray(weeklyRowCount, 9) = Round(W_avgDaily, 2): weeklyOutputArray(weeklyRowCount, 10) = DAILY_TARGET_HOURS
                    If DAILY_TARGET_HOURS > 0 Then weeklyOutputArray(weeklyRowCount, 11) = W_avgDaily / DAILY_TARGET_HOURS Else weeklyOutputArray(weeklyRowCount, 11) = 0
                End If
            Next sortIdx
            If weeklyRowCount > 0 Then
                .Range("A4").Resize(weeklyRowCount, 11).Value = weeklyOutputArray
                .Range("A4:B" & 3 + weeklyRowCount).NumberFormat = "m/d/yyyy"
                .Range("K4:K" & 3 + weeklyRowCount).NumberFormat = "0.00%"
                For rowIdx = 4 To 3 + weeklyRowCount
                    prodValue_weekly = 0: If IsNumeric(.Cells(rowIdx, 11).Value) Then prodValue_weekly = CDbl(.Cells(rowIdx, 11).Value)
                    If prodValue_weekly >= 1 Then
                        .Cells(rowIdx, 11).Interior.Color = RGB(200, 255, 200)
                    ElseIf prodValue_weekly >= 0.9 Then
                        .Cells(rowIdx, 11).Interior.Color = RGB(255, 255, 200)
                    Else
                        .Cells(rowIdx, 11).Interior.Color = RGB(255, 200, 200)
                    End If
                Next rowIdx
                .Range("A3:K" & 3 + weeklyRowCount).Borders.LineStyle = xlContinuous
                If .AutoFilterMode Then .AutoFilterMode = False
                .Range("A3:K3").AutoFilter
            End If
        End If
        .Columns("A:K").AutoFit: Application.Goto .Range("A1"), True
    End With

    ' -- Monthly Breakdown --
    With wsMonthlyBreakdown
        .Range("A1").Value = "Monthly Productivity Breakdown": .Range("A1").Font.Bold = True: .Range("A1").Font.Size = 14
        .Range("A3:J3").Value = Array("Month", "Team Member", "Total Prod. Hrs", "Actual Workdays", "Total Sick/Away Hrs", "Equiv. Sick/Away Days", "Adjusted Workdays", "Avg Prod. Hrs/Adj.Day", "Target", "Productivity %")
        .Range("A3:J3").Font.Bold = True: .Range("A3:J3").Interior.Color = RGB(200, 220, 255)
        If personMonthlyData.Count > 0 Then
            ReDim monthlyOutputArray(1 To personMonthlyData.Count, 1 To 10): monthlyRowCount = 0
            Set sortMap = CreateObject("Scripting.Dictionary"): ReDim sortKeys(0 To personMonthlyData.Count - 1): sortIdx = 0
            For Each key In personMonthlyData.Keys
                parts = Split(CStr(key), "|"): personName = parts(0): monthKey = parts(1)
                sortKey = monthKey & "|" & personName
                sortKeys(sortIdx) = sortKey: sortMap(sortKey) = CStr(key): sortIdx = sortIdx + 1
            Next key
            SortArrayInPlace sortKeys
            For sortIdx = 0 To UBound(sortKeys)
                originalKey = sortMap(sortKeys(sortIdx))
                parts = Split(originalKey, "|"): personNamePart = parts(0): monthPart = parts(1)
                totalProdHrs = personMonthlyData(originalKey)("TotalProdHrs"): actualWorkDays = personMonthlyData(originalKey)("ActualWorkDaysDict").Count
                totalSAHrs = personMonthlyData(originalKey)("TotalSickAwayHours")
                If HOURS_PER_SICK_AWAY_DAY > 0 Then equivSADays = totalSAHrs / HOURS_PER_SICK_AWAY_DAY Else equivSADays = 0
                adjWDays = personMonthlyAdjWorkdaySum(originalKey)
                If adjWDays < 0 Then adjWDays = 0
                If adjWDays > 0 Then avgDaily = totalProdHrs / adjWDays Else avgDaily = 0
                If totalProdHrs > 0 Or adjWDays > 0 Or actualWorkDays > 0 Then
                    monthlyRowCount = monthlyRowCount + 1
                    monthlyOutputArray(monthlyRowCount, 1) = Format(CDate(monthPart & "-01"), "yyyy-mmm"): monthlyOutputArray(monthlyRowCount, 2) = personNamePart
                    monthlyOutputArray(monthlyRowCount, 3) = Round(totalProdHrs, 2): monthlyOutputArray(monthlyRowCount, 4) = actualWorkDays
                    monthlyOutputArray(monthlyRowCount, 5) = Round(totalSAHrs, 2): monthlyOutputArray(monthlyRowCount, 6) = Round(equivSADays, 2)
                    monthlyOutputArray(monthlyRowCount, 7) = Round(adjWDays, 2): monthlyOutputArray(monthlyRowCount, 8) = Round(avgDaily, 2)
                    monthlyOutputArray(monthlyRowCount, 9) = DAILY_TARGET_HOURS
                    If DAILY_TARGET_HOURS > 0 Then monthlyOutputArray(monthlyRowCount, 10) = avgDaily / DAILY_TARGET_HOURS Else monthlyOutputArray(monthlyRowCount, 10) = 0
                End If
            Next sortIdx
            If monthlyRowCount > 0 Then
                .Range("A4").Resize(monthlyRowCount, 10).Value = monthlyOutputArray
                .Range("J4:J" & 3 + monthlyRowCount).NumberFormat = "0.00%"
                For rowIdx = 4 To 3 + monthlyRowCount
                    prodValue_monthly = 0: If IsNumeric(.Cells(rowIdx, 10).Value) Then prodValue_monthly = CDbl(.Cells(rowIdx, 10).Value)
                    If prodValue_monthly >= 1 Then
                        .Cells(rowIdx, 10).Interior.Color = RGB(200, 255, 200)
                    ElseIf prodValue_monthly >= 0.9 Then
                        .Cells(rowIdx, 10).Interior.Color = RGB(255, 255, 200)
                    Else
                        .Cells(rowIdx, 10).Interior.Color = RGB(255, 200, 200)
                    End If
                Next rowIdx
                .Range("A3:J" & 3 + monthlyRowCount).Borders.LineStyle = xlContinuous
                If .AutoFilterMode Then .AutoFilterMode = False
                .Range("A3:J3").AutoFilter
            End If
        End If
        .Columns("A:J").AutoFit: Application.Goto .Range("A1"), True
    End With
    
    ' -- Daily Breakdown --
    With wsDailyBreakdown
        .Range("A1").Value = "Daily Productivity Breakdown": .Range("A1").Font.Bold = True: .Range("A1").Font.Size = 14
        .Range("A3:H3").Value = Array("Date", "Month Context", "Team Member", "Productive Hrs (Day)", "Sick/Away Hrs (Day)", "Adjusted Workday Factor", "Effective Target Hrs", "Productivity % (Day)")
        .Range("A3:H3").Font.Bold = True: .Range("A3:H3").Interior.Color = RGB(240, 240, 210)
        If allActivityDays.Count > 0 Then
            ReDim dailyOutputArray(1 To allActivityDays.Count, 1 To 8): dailyRowCount = 0
            Set sortMap = CreateObject("Scripting.Dictionary"): ReDim sortKeys(0 To allActivityDays.Count - 1): sortIdx = 0
            For Each key In allActivityDays.Keys
                parts = Split(CStr(key), "|"): personName = parts(0): workDate = CDate(parts(1))
                sortKey = Format(workDate, "YYYYMMDD") & "|" & personName
                sortKeys(sortIdx) = sortKey: sortMap(sortKey) = CStr(key): sortIdx = sortIdx + 1
            Next key
            SortArrayInPlace sortKeys
            For sortIdx = 0 To UBound(sortKeys)
                originalKey = sortMap(sortKeys(sortIdx))
                parts = Split(originalKey, "|"): D_PersonName = parts(0): D_WorkDate = CDate(parts(1))
                D_MonthContext = Format(D_WorkDate, "mmm-yyyy")
                If dailyHoursDict.Exists(originalKey) Then D_ProdHrs = dailyHoursDict(originalKey) Else D_ProdHrs = 0
                If personDaySickAwayHoursDict.Exists(originalKey) Then D_SAHrs = personDaySickAwayHoursDict(originalKey) Else D_SAHrs = 0
                If HOURS_PER_SICK_AWAY_DAY > 0 Then D_AdjWorkdayFactor = 1 - (D_SAHrs / HOURS_PER_SICK_AWAY_DAY) Else D_AdjWorkdayFactor = 1
                If D_AdjWorkdayFactor < 0 Then D_AdjWorkdayFactor = 0
                If D_AdjWorkdayFactor > 1 Then D_AdjWorkdayFactor = 1
                D_EffectiveTarget = DAILY_TARGET_HOURS * D_AdjWorkdayFactor
                If D_EffectiveTarget > 0 Then
                    D_Productivity = D_ProdHrs / D_EffectiveTarget
                ElseIf D_ProdHrs > 0 And D_EffectiveTarget = 0 Then
                    D_Productivity = 1
                Else
                    D_Productivity = 0
                End If
                dailyRowCount = dailyRowCount + 1
                dailyOutputArray(dailyRowCount, 1) = D_WorkDate: dailyOutputArray(dailyRowCount, 2) = D_MonthContext
                dailyOutputArray(dailyRowCount, 3) = D_PersonName: dailyOutputArray(dailyRowCount, 4) = Round(D_ProdHrs, 2)
                dailyOutputArray(dailyRowCount, 5) = Round(D_SAHrs, 2): dailyOutputArray(dailyRowCount, 6) = Round(D_AdjWorkdayFactor, 2)
                dailyOutputArray(dailyRowCount, 7) = Round(D_EffectiveTarget, 2): dailyOutputArray(dailyRowCount, 8) = D_Productivity
            Next sortIdx
            If dailyRowCount > 0 Then
                .Range("A4").Resize(dailyRowCount, 8).Value = dailyOutputArray
                .Range("A4:A" & 3 + dailyRowCount).NumberFormat = "m/d/yyyy"
                .Range("H4:H" & 3 + dailyRowCount).NumberFormat = "0.00%"
                For rowIdx = 4 To 3 + dailyRowCount
                    prodValue_daily = 0: If IsNumeric(.Cells(rowIdx, 8).Value) Then prodValue_daily = CDbl(.Cells(rowIdx, 8).Value)
                    If .Cells(rowIdx, 7).Value > 0 Then
                        If prodValue_daily >= 1 Then
                            .Cells(rowIdx, 8).Interior.Color = RGB(200, 255, 200)
                        ElseIf prodValue_daily >= 0.9 Then
                            .Cells(rowIdx, 8).Interior.Color = RGB(255, 255, 200)
                        Else
                            .Cells(rowIdx, 8).Interior.Color = RGB(255, 200, 200)
                        End If
                    Else
                        .Cells(rowIdx, 8).Interior.Pattern = xlNone
                    End If
                Next rowIdx
                .Range("A3:H" & 3 + dailyRowCount).Borders.LineStyle = xlContinuous
                If .AutoFilterMode Then .AutoFilterMode = False
                .Range("A3:H3").AutoFilter
            End If
        End With
        .Columns("A:H").AutoFit: Application.Goto .Range("A1"), True
    End With
    
    ' -- Final Formatting for Dashboard --
    With wsDashboard
        maxDataRow = Application.WorksheetFunction.Max(monthRow - 1, 3)
        If maxDataRow > 3 Then .Range("A3:H" & maxDataRow).Borders.LineStyle = xlContinuous
        If monthRow > 4 Then .Range("H4:H" & monthRow - 1).NumberFormat = "0.00%"
        .Range("A3:H3").EntireColumn.AutoFit
        .Activate
        If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False
        If .Rows.Count > 3 And monthRow > 4 Then .Range("A4").Activate: ActiveWindow.FreezePanes = True
        .Cells(maxDataRow + 2, "A").Value = "Last Updated: " & Now()
        .Cells(maxDataRow + 2, "A").Font.Italic = True
        Application.Goto .Range("A1"), True
    End With
    
    GoTo AllDone

NoDataToProcess:
    wsDashboard.Activate
    MsgBox "No data was found in the 'Output' sheets after processing. The report generation has been cancelled.", vbInformation, "No Data to Report"
    
AllDone:
    ' --- Create or Update the Info sheet with all explanations ---    Call UpdateInfoSheet(DAILY_TARGET_HOURS, HOURS_PER_SICK_AWAY_DAY)
    
    ' *** PERFORMANCE REPORTING ***
    Dim metricsEndTime As Double: metricsEndTime = Timer
    Debug.Print "PERFORMANCE: Metrics calculation completed in " & Format((metricsEndTime - metricsStartTime), "0.00") & " seconds"
    
    endTime = Timer
    execTime = Format((endTime - startTime), "0.00") & " seconds"
    
    Debug.Print "PERFORMANCE: Total execution time breakdown:"
    Debug.Print "  - Import phase: " & Format((metricsStartTime - startTime), "0.00") & " seconds"
    Debug.Print "  - Metrics phase: " & Format((metricsEndTime - metricsStartTime), "0.00") & " seconds"
    Debug.Print "  - Total time: " & execTime
    
    ThisWorkbook.Sheets("ProductivityDashboard").Activate
    MsgBox "Full process complete! Data was imported and reports generated." & vbNewLine & _
           "Total execution time: " & execTime & vbNewLine & vbNewLine & _           
           "Performance details logged to Debug console.", vbInformation, "Process Complete"
End Sub

'==========================================================================
' --- PURE VBA QUICK SORT ALGORITHM ---
'==========================================================================
Public Sub SortArrayInPlace(ByRef arr() As String)
    If UBound(arr) >= LBound(arr) Then
        QuickSort_VBA arr, LBound(arr), UBound(arr)
    End If
End Sub
Private Sub QuickSort_VBA(ByRef arr() As String, ByVal L As Long, ByVal r As Long)
    Dim i As Long, j As Long, pivot As String, temp As String
    i = L: j = r: pivot = arr((L + r) \ 2)
    Do While i <= j
        Do While arr(i) < pivot And i < r: i = i + 1: Loop
        Do While pivot < arr(j) And j > L: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If L < j Then QuickSort_VBA arr, L, j
    If i < r Then QuickSort_VBA arr, i, r
End Sub


'==========================================================================
' --- HELPER SUBROUTINE TO CREATE/UPDATE THE INFO SHEET ---
'==========================================================================
Private Sub UpdateInfoSheet(ByVal dailyTarget As Double, ByVal sickDayHours As Double)
    Dim wsInfo As Worksheet, currentRow As Long
    On Error Resume Next
    Set wsInfo = ThisWorkbook.Sheets("Info"): On Error GoTo 0
    If wsInfo Is Nothing Then Set wsInfo = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)): wsInfo.name = "Info" Else wsInfo.Cells.Clear
    currentRow = 1
    Call WriteExplanationBlock(wsInfo, currentRow, "ProductivityDashboard - Column Explanations", "", True)
    Call WriteExplanationBlock(wsInfo, currentRow, "Month:", "The calendar month for which the data is aggregated.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Active Team Members:", "The unique count of team members who logged any time (either productive hours or sick/away time) during that month.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Total Productive Hours:", "The sum of all productive hours logged by all team members for the entire month.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Total Adjusted Workdays:", "The total number of 'workable' days for the entire team, accounting for time off. Calculated by summing the 'Adjusted Workday Factor' (from the DailyBreakdown) for every person for every day in the month.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Target Avg Prod. Hrs/Day:", "The daily productive hours target, read from the 'Config' sheet. Currently set to " & dailyTarget & " hours.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Productive Members:", "The count of unique team members who logged at least one productive hour during the month. This count is the denominator for the 'Met Target %'.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Members Meeting Target:", "The count of 'Productive Members' whose monthly average productive hours met or exceeded the target. An individual's average is calculated as: (Their Total Productive Hours for the month / Their Total Adjusted Workdays for the month).")
    Call WriteExplanationBlock(wsInfo, currentRow, "Met Target %:", "The percentage of 'Productive Members' who met the productivity target. Calculated as: (Members Meeting Target / Productive Members).")
    currentRow = currentRow + 1
    Call WriteExplanationBlock(wsInfo, currentRow, "MonthlyBreakdown - Column Explanations", "", True)
    Call WriteExplanationBlock(wsInfo, currentRow, "Month / Team Member:", "The individual being measured and the month of the activity.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Total Prod. Hrs:", "The sum of all productive hours for that person for that month.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Actual Workdays:", "A simple count of the unique days in the month on which the person logged any time (productive or sick/away).")
    Call WriteExplanationBlock(wsInfo, currentRow, "Total Sick/Away Hrs:", "The sum of all sick or vacation hours for that person for that month.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Equiv. Sick/Away Days:", "Converts sick/away hours into day equivalents. Calculated as: (Total Sick/Away Hrs / " & sickDayHours & "). The denominator is read from the 'Config' sheet.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Adjusted Workdays:", "The number of 'workable' days for the person in the month. This is the sum of their daily 'Adjusted Workday Factor' values.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Avg Prod. Hrs/Adj.Day:", "The person's average daily productivity, adjusted for time off. Calculated as: (Total Prod. Hrs / Adjusted Workdays). This is the key individual performance metric.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Target:", "The daily productive hours target, read from the 'Config' sheet. Currently " & dailyTarget & " hours.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Productivity %:", "The person's performance relative to the target. Calculated as: (Avg Prod. Hrs/Adj.Day / Target).")
    currentRow = currentRow + 1
    Call WriteExplanationBlock(wsInfo, currentRow, "WeeklyBreakdown - Column Explanations", "", True)
    Call WriteExplanationBlock(wsInfo, currentRow, "All columns:", "These are calculated identically to their 'MonthlyBreakdown' counterparts, but the data is aggregated on a weekly basis (Sunday to Saturday).")
    currentRow = currentRow + 1
    Call WriteExplanationBlock(wsInfo, currentRow, "DailyBreakdown - Column Explanations", "", True)
    Call WriteExplanationBlock(wsInfo, currentRow, "Date / Month Context / Team Member:", "The specific date and person for the daily record.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Productive Hrs (Day):", "The total productive hours logged by the person on that specific day.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Sick/Away Hrs (Day):", "The total sick/vacation hours logged by the person on that specific day.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Adjusted Workday Factor:", "The fraction of the day the person was 'available' to work. Calculated as: 1 - (Sick/Away Hrs (Day) / " & sickDayHours & "). A value of 1.0 means no time off; 0.0 means a full day off.")
    Call WriteExplanationBlock(wsInfo, currentRow, "Effective Target Hrs:", "The pro-rated productivity target for that person for that day, based on their availability. Calculated as: (Target * Adjusted Workday Factor).")
    Call WriteExplanationBlock(wsInfo, currentRow, "Productivity % (Day):", "The person's performance for that single day against their effective target. Calculated as: (Productive Hrs (Day) / Effective Target Hrs).")
    currentRow = currentRow + 1
    With wsInfo
        .Columns("A").ColumnWidth = 35: .Columns("B").ColumnWidth = 110
        .UsedRange.Rows.AutoFit: Application.Goto .Range("A1"), True
    End With
End Sub
Private Sub WriteExplanationBlock(ByRef ws As Worksheet, ByRef currentRow As Long, ByVal title As String, ByVal explanation As String, Optional ByVal isHeader As Boolean = False)
    If isHeader Then
        With ws.Range("A" & currentRow & ":B" & currentRow)
            .Merge: .Value = title: .Font.Bold = True: .Font.Size = 14
            .Font.Underline = xlUnderlineStyleSingle: .Interior.Color = RGB(242, 242, 242)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
    Else
        ws.Cells(currentRow, "A").Value = title: ws.Cells(currentRow, "A").Font.Bold = True
        ws.Cells(currentRow, "B").Value = explanation: ws.Range("B" & currentRow).WrapText = True
    End If
    ws.Rows(currentRow).VerticalAlignment = xlTop: currentRow = currentRow + 1
End Sub

'==========================================================================
' --- PERFORMANCE OPTIMIZED HELPER FUNCTIONS ---
'==========================================================================

'==========================================================================
' --- FAST CHECK: Determine if a date needs importing without heavy operations ---
'==========================================================================
Private Function NeedsImport(checkDate As Date) As Boolean
    Dim sheetName As String
    Dim targetSheet As Worksheet
    
    ' Quick sheet existence check first
    sheetName = "Personal Entry " & Format(checkDate, "m-d-yy")
    
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If targetSheet Is Nothing Then
        NeedsImport = True ' Sheet doesn't exist, needs import
        Exit Function
    End If
    
    ' *** OPTIMIZED: Check only a small, strategic range instead of huge area ***
    ' Check if any data exists in the first few rows/columns where data should be
    Dim quickCheckRange As Range
    Set quickCheckRange = targetSheet.Range("C3:H8") ' Much smaller, faster check
    
    If Application.WorksheetFunction.CountA(quickCheckRange) = 0 Then
        NeedsImport = True ' No data found, needs import
    Else
        NeedsImport = False ' Data exists, skip import
    End If
    
    Set targetSheet = Nothing
End Function

'==========================================================================
' --- BULK IMPORT: Process multiple dates with one source workbook session ---
'==========================================================================
Private Function BulkImportDataForDates(datesToProcess() As Date) As Boolean
    Dim sourceURL As String, sourceWB As Workbook
    Dim i As Long, processDate As Date
    Dim successCount As Long, totalCount As Long
    
    On Error GoTo BulkErrorHandler
    
    totalCount = UBound(datesToProcess) + 1
    Application.StatusBar = "Opening source workbook for bulk import..."
    
    ' --- PERFORMANCE: Open source workbook only ONCE ---
    sourceURL = ThisWorkbook.Sheets("Config").Range("Config_SourceWorkbookPath").Value
    Set sourceWB = Workbooks.Open(sourceURL, ReadOnly:=True, UpdateLinks:=False)
    
    Debug.Print "PERFORMANCE: Source workbook opened once for " & totalCount & " dates"
    
    ' --- Process all dates with the same open workbook ---
    For i = 0 To UBound(datesToProcess)
        processDate = datesToProcess(i)
        Application.StatusBar = "Bulk processing (" & (i + 1) & "/" & totalCount & "): " & Format(processDate, "M/D/YYYY")
        
        If BulkImportSingleDate(sourceWB, processDate) Then
            successCount = successCount + 1
            Debug.Print "BULK SUCCESS: " & Format(processDate, "M/D/YYYY")
        Else
            Debug.Print "BULK FAILED: " & Format(processDate, "M/D/YYYY")
        End If
    Next i
    
    ' Close source workbook
    sourceWB.Close SaveChanges:=False
    Set sourceWB = Nothing
    
    Debug.Print "PERFORMANCE: Bulk import completed. " & successCount & "/" & totalCount & " successful"
    
    ' Return success if at least 80% succeeded
    BulkImportDataForDates = (successCount >= (totalCount * 0.8))
    Exit Function
    
BulkErrorHandler:
    Debug.Print "BULK IMPORT ERROR: " & Err.Description
    If Not sourceWB Is Nothing Then sourceWB.Close SaveChanges:=False
    BulkImportDataForDates = False
End Function

'==========================================================================
' --- OPTIMIZED SINGLE DATE IMPORT: No workbook open/close overhead ---
'==========================================================================
Private Function BulkImportSingleDate(sourceWB As Workbook, processDate As Date) As Boolean
    Dim processDateStr As String
    Dim ws As Worksheet
    Dim sourcePersonal As Worksheet, sourceNonEntry As Worksheet
    Dim targetPersonal As Worksheet, targetNonEntry As Worksheet
    Dim templatePersonal As Worksheet, templateNonEntry As Worksheet
    
    On Error GoTo SingleDateError
    
    processDateStr = Format(processDate, "yyyy-mm-dd")
    
    ' Get template references (should be fast since workbook is local)
    Set templatePersonal = ThisWorkbook.Sheets("Personal Entry")
    Set templateNonEntry = ThisWorkbook.Sheets("Non-Entry Hrs")
    
    ' *** OPTIMIZED: Find source sheets faster with early exit ***
    For Each ws In sourceWB.Worksheets
        If sourcePersonal Is Nothing And ws.name Like "Personal Entry *" Then
            If ParseDateFromName(ws.name, "Personal Entry ") = processDateStr Then
                Set sourcePersonal = ws
            End If
        End If
        If sourceNonEntry Is Nothing And ws.name Like "Non-Entry Hrs *" Then
            If ParseDateFromName(ws.name, "Non-Entry Hrs ") = processDateStr Then
                Set sourceNonEntry = ws
            End If
        End If
        
        ' Early exit when both found
        If Not sourcePersonal Is Nothing And Not sourceNonEntry Is Nothing Then Exit For
    Next ws
    
    If sourcePersonal Is Nothing Or sourceNonEntry Is Nothing Then
        Debug.Print "Missing source sheets for " & Format(processDate, "M/D/YYYY")
        BulkImportSingleDate = True ' Continue processing other dates
        Exit Function
    End If
    
    ' *** OPTIMIZED: Fast sheet creation and data copy ***
    If Not CreateOrUpdateTargetSheet(sourcePersonal, templatePersonal) Then GoTo SingleDateError
    If Not CreateOrUpdateTargetSheet(sourceNonEntry, templateNonEntry) Then GoTo SingleDateError
    
    BulkImportSingleDate = True
    Exit Function
    
SingleDateError:
    Debug.Print "Error in BulkImportSingleDate: " & Err.Description
    BulkImportSingleDate = False
End Function

'==========================================================================
' --- ULTRA-FAST SHEET CREATION AND DATA COPY ---
'==========================================================================
Private Function CreateOrUpdateTargetSheet(sourceSheet As Worksheet, templateSheet As Worksheet) As Boolean
    Dim targetSheet As Worksheet
    Dim sheetName As String
    
    On Error GoTo CreateError
    
    sheetName = sourceSheet.name
    
    ' Check if target sheet exists
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo CreateError
    
    If targetSheet Is Nothing Then
        ' *** OPTIMIZED: Fast sheet creation ***
        templateSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set targetSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        targetSheet.name = sheetName
    Else
        ' *** OPTIMIZED: Fast clear - only clear data areas, not entire used range ***
        If sourceSheet.name Like "Personal Entry *" Then
            targetSheet.Range("C3:ZZ100").ClearContents ' Fixed range is faster
        Else
            targetSheet.Range("D2:ZZ100").ClearContents ' Fixed range is faster
        End If
    End If
    
    ' *** OPTIMIZED: Bulk data copy using fastest method ***
    Call FastDataCopy(sourceSheet, targetSheet)
    
    CreateOrUpdateTargetSheet = True
    Exit Function
    
CreateError:
    Debug.Print "Error creating/updating sheet: " & Err.Description
    CreateOrUpdateTargetSheet = False
End Function

'==========================================================================
' --- FASTEST POSSIBLE DATA COPY METHOD ---
'==========================================================================
Private Sub FastDataCopy(sourceSheet As Worksheet, targetSheet As Worksheet)
    On Error Resume Next ' Continue on errors for speed
    
    If sourceSheet.name Like "Personal Entry *" Then
        ' *** OPTIMIZED: Copy entire used range in one operation ***
        Dim sourceRange As Range
        Set sourceRange = sourceSheet.Range("C3:ZZ100") ' Use fixed range for speed
        
        If Application.WorksheetFunction.CountA(sourceRange) > 0 Then
            targetSheet.Range("C3:ZZ100").Value = sourceRange.Value
        End If
    Else
        ' Non-Entry sheet
        Set sourceRange = sourceSheet.Range("D2:ZZ100")
        If Application.WorksheetFunction.CountA(sourceRange) > 0 Then
            targetSheet.Range("D2:ZZ100").Value = sourceRange.Value
        End If
    End If
    
    On Error GoTo 0
End Sub