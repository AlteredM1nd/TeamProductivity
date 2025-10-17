' =========================================================================
' --- MAIN MODULE (CONTROL FLOW & REPORTING) ---
' =========================================================================
Option Explicit
Public SkipReprocessDuringMetrics As Boolean

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
Private Function ImportDataForDate(ByVal processDate As Date, Optional ByVal sheetChoice As String = "Both") As Boolean
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
    ' Only import Personal Entry if requested
    If (sheetChoice = "Both" Or sheetChoice = "Output") Then
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
    End If
    
    ' -- Handle Non-Entry Sheet --
    ' Only import Non-Entry if requested
    If (sheetChoice = "Both" Or sheetChoice = "OutputNE") Then
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
    End If

    ' --- 4. Copy Data via "Clean and Paste" Method ---
    Dim dataArray As Variant, r As Long, c As Long
    
    ' -- For Personal Entry --
    Dim lastDataRowPE As Long, lastDataColPE As Long
    
    ' Find the last row based on names in column A of the SOURCE sheet
    lastDataRowPE = sourcePersonal.Cells(sourcePersonal.Rows.Count, "A").End(xlUp).row
    
    ' *** NEW: Find the last column based on the headers in your LOCAL TEMPLATE ***
    lastDataColPE = templatePersonal.Cells(2, templatePersonal.Columns.Count).End(xlToLeft).Column
    
    If (sheetChoice = "Both" Or sheetChoice = "Output") And lastDataRowPE >= 3 And lastDataColPE >= 3 And Not sourcePersonal Is Nothing Then ' Ensure there is data to copy and source exists
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
        
        ' Step 3: Paste from memory (guard target exists)
        If Not targetPersonal Is Nothing Then
            targetPersonal.Range("C3").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).Value = dataArray
        Else
            Debug.Print "Skipping paste to Personal target because targetPersonal is Nothing for date " & processDateStr
        End If
    End If
    
    ' -- For Non-Entry Hrs --
    Dim lastDataRowNE As Long, lastDataColNE As Long
    lastDataRowNE = sourceNonEntry.Cells(sourceNonEntry.Rows.Count, "A").End(xlUp).row
    lastDataColNE = sourceNonEntry.Cells(1, sourceNonEntry.Columns.Count).End(xlToLeft).Column
    If (sheetChoice = "Both" Or sheetChoice = "OutputNE") And lastDataRowNE >= 2 And lastDataColNE >= 4 And Not sourceNonEntry Is Nothing Then
        dataArray = sourceNonEntry.Range(sourceNonEntry.Cells(2, 4), sourceNonEntry.Cells(lastDataRowNE, lastDataColNE)).Value2
        For r = 1 To UBound(dataArray, 1)
            For c = 1 To UBound(dataArray, 2)
                If IsError(dataArray(r, c)) Then dataArray(r, c) = ""
            Next c
        Next r
        If Not targetNonEntry Is Nothing Then
            targetNonEntry.Range("D2").Resize(UBound(dataArray, 1), UBound(dataArray, 2)).Value = dataArray
        Else
            Debug.Print "Skipping paste to Non-Entry target because targetNonEntry is Nothing for date " & processDateStr
        End If
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
Private Sub CalculateProductivityMetrics(ByVal startTime As Double, Optional ByRef datesToProcess As Variant = Empty, Optional ByVal dateCount As Long = 0)
    ' *** PERFORMANCE: Start metrics calculation timing ***
    Dim metricsStartTime As Double: metricsStartTime = Timer
    Application.StatusBar = "Calculating productivity metrics..."
    
    ' --- ALL VARIABLES ---
    ' (Variable list remains the same)
    Dim wsOutput As Worksheet, wsOutputNE As Worksheet, wsDashboard As Worksheet, wsMonthlyBreakdown As Worksheet, wsWeeklyBreakdown As Worksheet, wsDailyBreakdown As Worksheet
    Dim dailyHoursDict As Object, personDaySickAwayHoursDict As Object, personMonthlyData As Object, personWeeklyData As Object, allTeamMembersMasterDict As Object
    Dim dashboardMonthlyAggregator As Object, allActivityDays As Object, personMonthlyAdjWorkdaySum As Object, personWeeklyAdjWorkdaySum As Object
    Dim arrOutput As Variant, arrOutputNE As Variant, weeklyOutputArray As Variant, monthlyOutputArray As Variant, dailyOutputArray As Variant, overridesDict As Object
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

    '--- STEP 1.5: LOAD OVERRIDES ---
    Application.StatusBar = "Step 2: Loading overrides..."
    Set overridesDict = LoadOverridesDict()


    '--- STEP 2: REBUILD THE OUTPUT SHEETS FROM ALL RELEVANT DATED SHEETS ---
    Application.StatusBar = "Step 3: Rebuilding Output sheets from all dated sources for the year..."
    Set wsOutput = ThisWorkbook.Sheets("Output")
    Set wsOutputNE = ThisWorkbook.Sheets("OutputNE")

    ' If a global flag requests skipping reprocessing, don't rebuild dated sheets here
    If Not SkipReprocessDuringMetrics Then
        If Not IsEmpty(datesToProcess) And dateCount > 0 Then
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
        Dim reportStartDate: reportStartDate = DateSerial(2024, 1, 1)
        For Each localSheet In ThisWorkbook.Worksheets
            If localSheet.name Like "Personal Entry *" Then
                If localSheet.name <> "Personal Entry" Then
                    parsedDate = ParseDateFromName(localSheet.name, "Personal Entry ")
                    If parsedDate <> "" Then
                        sheetDate = CDate(parsedDate)
                        If sheetDate >= reportStartDate Then
                            If (IsEmpty(datesToProcess) Or dateCount = 0) Or processDatesDict.Exists(parsedDate) Then
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
                            If (IsEmpty(datesToProcess) Or dateCount = 0) Or processDatesDict.Exists(parsedDate) Then
                                Call ProcessNonEntrySheet(localSheet, parsedDate)
                            End If
                        End If
                    End If
                End If
            End If
        Next localSheet
    Else
        ' Skipping reprocessing of dated sheets because SkipReprocessDuringMetrics is True
    End If

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

    '--- NEW STEP 3.5: INTEGRATE HISTORICAL/OVERRIDE DATA ---
    If overridesDict.Count > 0 Then
        Application.StatusBar = "Step 4.5: Integrating historical override data..."
        For Each key In overridesDict.Keys
            ' Check if this person/month combo already exists from the source data
            If Not personMonthlyData.Exists(key) Then
                parts = Split(CStr(key), "|"): personName = parts(0): monthKey = parts(1)
                
                ' Add the person to the master list if they are new
                If Not allTeamMembersMasterDict.Exists(personName) Then
                    allTeamMembersMasterDict.Add personName, 1
                    Debug.Print "Adding historical person from overrides: " & personName
                End If

                ' Create a placeholder entry in the monthly data dictionary. This ensures they appear in the reports.
                Set personMonthlyData(key) = CreateObject("Scripting.Dictionary")
                personMonthlyData(key)("TotalProdHrs") = 0
                Set personMonthlyData(key)("ActualWorkDaysDict") = CreateObject("Scripting.Dictionary")
                personMonthlyData(key)("TotalSickAwayHours") = 0
                personMonthlyAdjWorkdaySum(key) = 0 ' Also create a placeholder for adjusted workday sum
            End If
        Next key
    End If
    
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
        
        ' A person is relevant for the dashboard if they have logged hours OR have a manual override for the month.
        If personMonthlyData(key)("TotalProdHrs") > 0 Or personMonthlyData(key)("TotalSickAwayHours") > 0 Or overridesDict.Exists(key) Then
            If Not dashboardMonthlyAggregator.Exists(monthKey) Then
                Set dashboardMonthlyAggregator(monthKey) = CreateObject("Scripting.Dictionary")
                dashboardMonthlyAggregator(monthKey)("ActiveMembersCount") = 0: dashboardMonthlyAggregator(monthKey)("ProdEligibleMembersCount") = 0
                dashboardMonthlyAggregator(monthKey)("TotalProdHrsSum") = 0: dashboardMonthlyAggregator(monthKey)("TotalAdjWorkdaysSum") = 0
                Set dashboardMonthlyAggregator(monthKey)("ActiveMembersDict") = CreateObject("Scripting.Dictionary"): dashboardMonthlyAggregator(monthKey)("MembersMeetingTargetCount") = 0
            End If
            
            totalProdHrsPersonMonth = personMonthlyData(key)("TotalProdHrs"): adjustedWorkDays = personMonthlyAdjWorkdaySum(key)
            metTargetFlag = False
            
            ' Calculate the average daily hours for this person/month
            If adjustedWorkDays > 0 Then
                avgDailyPerson = totalProdHrsPersonMonth / adjustedWorkDays
            Else
                avgDailyPerson = 0
            End If
            
            ' *** OVERRIDE LOGIC: Check if an override exists and apply it for target checking ***
            If overridesDict.Exists(key) Then avgDailyPerson = overridesDict(key)
            
            ' Now, check if the potentially overridden average meets the target
            If avgDailyPerson >= DAILY_TARGET_HOURS Then metTargetFlag = True
            If Not dashboardMonthlyAggregator(monthKey)("ActiveMembersDict").Exists(personName) Then
                 dashboardMonthlyAggregator(monthKey)("ActiveMembersDict")(personName) = 1
                 dashboardMonthlyAggregator(monthKey)("ActiveMembersCount") = dashboardMonthlyAggregator(monthKey)("ActiveMembersCount") + 1
                 
                 ' A member is "Productive" (eligible for the target %) if they have productive hours OR an override.
                 If totalProdHrsPersonMonth > 0 Or overridesDict.Exists(key) Then
                    dashboardMonthlyAggregator(monthKey)("ProdEligibleMembersCount") = dashboardMonthlyAggregator(monthKey)("ProdEligibleMembersCount") + 1
                 End If
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
        .Columns("A:K").AutoFit: Application.GoTo .Range("A1"), True
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
                
                ' *** OVERRIDE LOGIC ***
                If overridesDict.Exists(originalKey) Then
                    avgDaily = overridesDict(originalKey) ' Use the override value
                Else
                    If adjWDays > 0 Then avgDaily = totalProdHrs / adjWDays Else avgDaily = 0 ' Calculate as normal
                End If
                
                If totalProdHrs > 0 Or adjWDays > 0 Or actualWorkDays > 0 Or overridesDict.Exists(originalKey) Then
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
        .Columns("A:J").AutoFit: Application.GoTo .Range("A1"), True
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
                .Columns("A:H").AutoFit: Application.GoTo .Range("A1"), True
            End If
        End If
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
        Application.GoTo .Range("A1"), True
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
        .UsedRange.Rows.AutoFit: Application.GoTo .Range("A1"), True
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
        BulkImportSingleDate = False ' *** FIX: Report failure so the process knows data is missing ***
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

'==========================================================================
' --- HELPER FUNCTION TO LOAD OVERRIDE DATA ---
'==========================================================================
Private Function LoadOverridesDict() As Object
    Dim wsOverrides As Worksheet
    Dim overridesDict As Object
    Dim lastRow As Long, r As Long
    Dim personName As String, monthStr As String, overrideValue As Double
    Dim overrideKey As String, monthDate As Date

    Set overridesDict = CreateObject("Scripting.Dictionary")
    overridesDict.CompareMode = vbTextCompare ' Case-insensitive for names

    On Error Resume Next
    Set wsOverrides = ThisWorkbook.Sheets("Overrides")
    On Error GoTo 0

    If wsOverrides Is Nothing Then
        Debug.Print "INFO: 'Overrides' sheet not found. No overrides will be applied."
        Set LoadOverridesDict = overridesDict ' Return empty dictionary
        Exit Function
    End If

    lastRow = wsOverrides.Cells(wsOverrides.Rows.Count, "A").End(xlUp).row

    If lastRow < 2 Then
        Debug.Print "INFO: 'Overrides' sheet is empty."
        Set LoadOverridesDict = overridesDict ' Return empty dictionary
        Exit Function
    End If

    ' Read all data into an array for performance
    Dim dataArr As Variant
    dataArr = wsOverrides.Range("A2:C" & lastRow).Value

    For r = 1 To UBound(dataArr, 1)
        personName = Trim(CStr(dataArr(r, 1)))
        monthStr = Trim(CStr(dataArr(r, 2)))

        ' Validate data before processing
        If Len(personName) > 0 And Len(monthStr) > 0 And IsNumeric(dataArr(r, 3)) Then
            On Error Resume Next
            monthDate = CDate("01-" & Replace(monthStr, "-", " ")) ' Convert yyyy-mmm to a date
            If Err.Number = 0 Then
                overrideKey = personName & "|" & Format(monthDate, "yyyy-mm")
                overrideValue = CDbl(dataArr(r, 3))
                overridesDict(overrideKey) = overrideValue
            Else
                Debug.Print "WARNING: Could not parse month '" & monthStr & "' in Overrides sheet, row " & r + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next r

    Set LoadOverridesDict = overridesDict
End Function

'==========================================================================
'--- UTILITY: Remove exact duplicate rows from Output sheets and rerun reports
'==========================================================================
Public Sub RemoveDuplicatesAndRecalculate(Optional ByVal dryRun As Boolean = False)
    Dim wsOut As Worksheet, wsOutNE As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range, r As Long
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim key As String, removed As Long
    Dim arr As Variant
    Dim hadOutputFilter As Boolean, hadOutputNEFilter As Boolean

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Removing exact duplicate rows from Output sheets..."

    On Error GoTo CleanUp

    ' Create a full workbook copy backup before modifying (safer for recovery)
    Call SaveWorkbookBackup
    ' Create timestamped backups of the sheets before modifying
    Dim ts As String: ts = Format(Now, "yyyyMMdd_HHmmss")
    On Error Resume Next
    ThisWorkbook.Sheets("Output").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "Output_backup_" & ts
    ThisWorkbook.Sheets("OutputNE").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "OutputNE_backup_" & ts
    On Error GoTo CleanUp

    ' Process Output (Productive) sheet
    Set wsOut = ThisWorkbook.Sheets("Output")
    ' Prepare per-run dedupe audit sheet
    Dim dedupeAudit As Worksheet, dedupeName As String, dedupeRow As Long
    dedupeName = "Dedupe_Audit_" & ts
    On Error Resume Next
    Set dedupeAudit = ThisWorkbook.Sheets(dedupeName)
    If dedupeAudit Is Nothing Then
        Set dedupeAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        dedupeAudit.Name = dedupeName
    End If
    On Error GoTo CleanUp
    lastRow = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row
    If lastRow > 1 Then
        lastCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
        Set dataRange = wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(lastRow, lastCol))
        arr = dataRange.Value
        ' Validate header columns vs data columns
        Dim headerCols As Long: headerCols = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
        ' Initialize dedupe audit header (Source, OriginalRow, RemovedAt, then columns)
        If dedupeAudit.Cells(1, 1).Value = "" Then
            dedupeAudit.Cells(1, 1).Value = "Source": dedupeAudit.Cells(1, 2).Value = "OriginalRow": dedupeAudit.Cells(1, 3).Value = "RemovedAt"
            dedupeAudit.Cells(1, 4).Resize(1, headerCols).Value = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, headerCols)).Value
        End If
        dict.RemoveAll
        removed = 0
        ' Build a new array of unique rows
        Dim outArr() As Variant
        Dim outPtr As Long: outPtr = 0
        Dim i As Long, j As Long
        Dim rowsCount As Long, colsCount As Long
        rowsCount = UBound(arr, 1): colsCount = UBound(arr, 2)

        ' Allocate a working buffer the same size as the input (no Preserve on first dimension)
        ReDim outArr(1 To rowsCount, 1 To colsCount)

        Dim cellVal As Variant, norm As String
        For i = 1 To rowsCount
            key = ""
            ' Build key using Date(1), Name(2), Region(3), Task(4), Count(5) only
            Dim colIdxsOut As Variant: colIdxsOut = Array(1, 2, 3, 4, 5)
            For j = LBound(colIdxsOut) To UBound(colIdxsOut)
                Dim cj As Long: cj = colIdxsOut(j)
                If cj <= colsCount Then
                    cellVal = arr(i, cj)
                Else
                    cellVal = ""
                End If
                If IsError(cellVal) Then
                    norm = "#ERR"
                ElseIf IsDate(cellVal) Then
                    norm = Format(CDate(cellVal), "yyyy-mm-dd")
                ElseIf IsNumeric(cellVal) Then
                    norm = Trim(CStr(cellVal))
                ElseIf IsNull(cellVal) Then
                    norm = ""
                Else
                    norm = Trim(CStr(cellVal))
                    norm = Replace(norm, vbTab, " ")
                    Do While InStr(norm, "  ") > 0
                        norm = Replace(norm, "  ", " ")
                    Loop
                End If
                key = key & "|" & norm
            Next j
            If Not dict.Exists(key) Then
                dict(key) = 1
                outPtr = outPtr + 1
                For j = 1 To colsCount
                    outArr(outPtr, j) = arr(i, j)
                Next j
            Else
                removed = removed + 1
                If removed <= 5 Then Debug.Print "Duplicate detected (Output) row ", i, " key=", Left(key, 200)
                ' Log removed row into dedupe audit
                On Error Resume Next
                dedupeRow = dedupeAudit.Cells(dedupeAudit.Rows.Count, "A").End(xlUp).Row + 1
                dedupeAudit.Cells(dedupeRow, 1).Value = "Output"
                dedupeAudit.Cells(dedupeRow, 2).Value = i + 1
                dedupeAudit.Cells(dedupeRow, 3).Value = Now
                For j = 1 To colsCount: dedupeAudit.Cells(dedupeRow, 3 + j).Value = arr(i, j): Next j
                On Error GoTo CleanUp
            End If
        Next i

    If outPtr > 0 Then
            ' Copy used portion into a correctly sized array for writing
            Dim writeArr() As Variant
            ReDim writeArr(1 To outPtr, 1 To colsCount)
            For i = 1 To outPtr
                For j = 1 To colsCount
                    writeArr(i, j) = outArr(i, j)
                Next j
            Next i

                ' --- Sanity checks: ensure dates in writeArr are consistent with original data ---
                Dim dateCounts As Object: Set dateCounts = CreateObject("Scripting.Dictionary")
                Dim origDateCounts As Object: Set origDateCounts = CreateObject("Scripting.Dictionary")
                Dim dVal As Variant, dKey As String
                For i = 1 To outPtr
                    dVal = writeArr(i, 1)
                    If IsDate(dVal) Then dKey = Format(CDate(dVal), "yyyy-mm-dd") Else dKey = CStr(dVal)
                    If dateCounts.Exists(dKey) Then dateCounts(dKey) = dateCounts(dKey) + 1 Else dateCounts(dKey) = 1
                Next i
                For i = 1 To rowsCount
                    dVal = arr(i, 1)
                    If IsDate(dVal) Then dKey = Format(CDate(dVal), "yyyy-mm-dd") Else dKey = CStr(dVal)
                    If origDateCounts.Exists(dKey) Then origDateCounts(dKey) = origDateCounts(dKey) + 1 Else origDateCounts(dKey) = 1
                Next i

                ' If writeArr contains a single repeated date for many rows, abort
                If dateCounts.Count = 1 And outPtr > 10 Then
                    Debug.Print "Sanity abort: writeArr contains a single date (" & dateCounts.Keys()(0) & ") for " & outPtr & " rows"
                    If Not dryRun Then MsgBox "Dedupe aborted: resulting data contains a single date for all rows (" & dateCounts.Keys()(0) & "). Please run a dry-run to inspect.", vbExclamation
                    GoTo SkipOutputWrite
                End If

                ' If writeArr contains dates not present in the original data, abort
                Dim diffCount As Long: diffCount = 0
                Dim k As Variant
                For Each k In dateCounts.Keys
                    If Not origDateCounts.Exists(k) Then diffCount = diffCount + dateCounts(k)
                Next k
                If diffCount > 0 Then
                    Debug.Print "Sanity abort: writeArr contains " & diffCount & " rows with dates not present in original data"
                    If Not dryRun Then MsgBox "Dedupe aborted: result contains dates not present in original data. Please run a dry-run to inspect.", vbExclamation
                    GoTo SkipOutputWrite
                End If

            ' Safety: if header column count differs from colsCount, abort write unless dryRun
            If Not dryRun Then
                If wsOut.AutoFilterMode Then
                    hadOutputFilter = True
                    On Error Resume Next
                    If wsOut.FilterMode Then wsOut.ShowAllData
                    On Error GoTo CleanUp
                    wsOut.AutoFilterMode = False
                End If
            End If
            If headerCols <> colsCount Then
                Debug.Print "Header columns (", headerCols, ") <> data cols (", colsCount, ") -- aborting write to Output"
                If Not dryRun Then
                    MsgBox "Header columns do not match data width. Aborting write to Output to avoid corruption.", vbExclamation
                    GoTo SkipOutputWrite
                End If
            End If
            If Not dryRun Then
                ' Clear old data and write back unique rows
                wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(lastRow, lastCol)).ClearContents
                wsOut.Cells(2, 1).Resize(outPtr, colsCount).Value = writeArr
                If hadOutputFilter Then wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, headerCols)).AutoFilter
            Else
                Debug.Print "DryRun: would write ", outPtr, " unique rows to Output"
            End If
SkipOutputWrite:
        End If
        Debug.Print "Removed " & removed & " duplicate rows from Output"
    End If

    ' Process OutputNE (Non-Entry) sheet
    Set wsOutNE = ThisWorkbook.Sheets("OutputNE")
    lastRow = wsOutNE.Cells(wsOutNE.Rows.Count, "A").End(xlUp).Row
    If lastRow > 1 Then
        lastCol = wsOutNE.Cells(1, wsOutNE.Columns.Count).End(xlToLeft).Column
        Set dataRange = wsOutNE.Range(wsOutNE.Cells(2, 1), wsOutNE.Cells(lastRow, lastCol))
        arr = dataRange.Value
        dict.RemoveAll
        removed = 0
        Dim outArrNE() As Variant
        outPtr = 0
        Dim rowsCountNE As Long, colsCountNE As Long
        rowsCountNE = UBound(arr, 1): colsCountNE = UBound(arr, 2)
    ReDim outArrNE(1 To rowsCountNE, 1 To colsCountNE)
    Dim headerColsNE As Long: headerColsNE = wsOutNE.Cells(1, wsOutNE.Columns.Count).End(xlToLeft).Column
        Dim normNE As String, cellValNE As Variant
        For i = 1 To rowsCountNE
            key = ""
            ' Build key using Date(1), Name(2), Task(3), Count(4)
            Dim colIdxsNE As Variant: colIdxsNE = Array(1, 2, 3, 4)
            For j = LBound(colIdxsNE) To UBound(colIdxsNE)
                Dim cjNE As Long: cjNE = colIdxsNE(j)
                If cjNE <= colsCountNE Then
                    cellValNE = arr(i, cjNE)
                Else
                    cellValNE = ""
                End If
                If IsError(cellValNE) Then
                    normNE = "#ERR"
                ElseIf IsDate(cellValNE) Then
                    normNE = Format(CDate(cellValNE), "yyyy-mm-dd")
                ElseIf IsNumeric(cellValNE) Then
                    normNE = Trim(CStr(cellValNE))
                ElseIf IsNull(cellValNE) Then
                    normNE = ""
                Else
                    normNE = Trim(CStr(cellValNE))
                    normNE = Replace(normNE, vbTab, " ")
                    Do While InStr(normNE, "  ") > 0
                        normNE = Replace(normNE, "  ", " ")
                    Loop
                End If
                key = key & "|" & normNE
            Next j
            If Not dict.Exists(key) Then
                dict(key) = 1
                outPtr = outPtr + 1
                For j = 1 To colsCountNE
                    outArrNE(outPtr, j) = arr(i, j)
                Next j
            Else
                removed = removed + 1
                If removed <= 5 Then Debug.Print "Duplicate detected (OutputNE) row ", i, " key=", Left(key, 200)
                ' Log removed row into dedupe audit
                On Error Resume Next
                dedupeRow = dedupeAudit.Cells(dedupeAudit.Rows.Count, "A").End(xlUp).Row + 1
                dedupeAudit.Cells(dedupeRow, 1).Value = "OutputNE"
                dedupeAudit.Cells(dedupeRow, 2).Value = i + 1
                dedupeAudit.Cells(dedupeRow, 3).Value = Now
                For j = 1 To colsCountNE: dedupeAudit.Cells(dedupeRow, 3 + j).Value = arr(i, j): Next j
                On Error GoTo CleanUp
            End If
        Next i

    If outPtr > 0 Then
            Dim writeArrNE() As Variant
            ReDim writeArrNE(1 To outPtr, 1 To colsCountNE)
            For i = 1 To outPtr
                For j = 1 To colsCountNE
                    writeArrNE(i, j) = outArrNE(i, j)
                Next j
            Next i
            ' Safety: check header columns
            If Not dryRun Then
                If wsOutNE.AutoFilterMode Then
                    hadOutputNEFilter = True
                    On Error Resume Next
                    If wsOutNE.FilterMode Then wsOutNE.ShowAllData
                    On Error GoTo CleanUp
                    wsOutNE.AutoFilterMode = False
                End If
            End If
            If headerColsNE <> colsCountNE Then
                Debug.Print "Header columns (", headerColsNE, ") <> data cols (", colsCountNE, ") -- aborting write to OutputNE"
                If Not dryRun Then
                    MsgBox "Header columns do not match data width. Aborting write to OutputNE to avoid corruption.", vbExclamation
                    GoTo SkipOutputNEWrite
                End If
            End If
            If Not dryRun Then
                wsOutNE.Range(wsOutNE.Cells(2, 1), wsOutNE.Cells(lastRow, lastCol)).ClearContents
                wsOutNE.Cells(2, 1).Resize(outPtr, colsCountNE).Value = writeArrNE
                If hadOutputNEFilter Then wsOutNE.Range(wsOutNE.Cells(1, 1), wsOutNE.Cells(1, headerColsNE)).AutoFilter
            Else
                Debug.Print "DryRun: would write ", outPtr, " unique rows to OutputNE"
            End If
SkipOutputNEWrite:
        End If
        Debug.Print "Removed " & removed & " duplicate rows from OutputNE"
    End If

    ' Re-run the metrics calculation (preserve existing Output/OutputNE AHT values)
    Application.StatusBar = "Recalculating productivity metrics (preserving existing Output data)..."
    Dim preserveDates() As Date
    ReDim preserveDates(0)
    preserveDates(0) = Date ' non-empty array prevents the routine from clearing Output/OutputNE
    Call CalculateProductivityMetrics(Timer, preserveDates, 1)

CleanUp:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    If Err.Number <> 0 Then Debug.Print "Error in RemoveDuplicatesAndRecalculate (" & Err.Number & "): " & Err.Description
End Sub

' Restore most recent backups for Output and OutputNE (looks for sheets named Output_backup_YYYYMMDD_HHMMSS)
Public Sub RestoreLastBackups()
    Dim ws As Worksheet, foundOut As Worksheet, foundNE As Worksheet
    Dim s As String, bestOut As String, bestNE As String
    Dim bestDTOut As Date, bestDTNE As Date
    bestOut = "": bestNE = ""
    For Each ws In ThisWorkbook.Worksheets
        s = ws.Name
        If LCase(Left(s, 13)) = "output_backup_" Then
            On Error Resume Next
            Dim dt As Date: dt = CDate(Mid(s, 14, 4) & "-" & Mid(s, 18, 2) & "-" & Mid(s, 20, 2) & " " & Mid(s, 23, 2) & ":" & Mid(s, 25, 2) & ":" & Mid(s, 27, 2))
            If Err.Number = 0 Then
                If bestOut = "" Or dt > bestDTOut Then bestOut = s: bestDTOut = dt
            End If
            On Error GoTo 0
        ElseIf LCase(Left(s, 16)) = "outputne_backup_" Then
            On Error Resume Next
            Dim dt2 As Date: dt2 = CDate(Mid(s, 17, 4) & "-" & Mid(s, 21, 2) & "-" & Mid(s, 23, 2) & " " & Mid(s, 26, 2) & ":" & Mid(s, 28, 2) & ":" & Mid(s, 30, 2))
            If Err.Number = 0 Then
                If bestNE = "" Or dt2 > bestDTNE Then bestNE = s: bestDTNE = dt2
            End If
            On Error GoTo 0
        End If
    Next ws

    If bestOut <> "" Then
        ThisWorkbook.Sheets(bestOut).Cells.Copy
        ThisWorkbook.Sheets("Output").Cells.Clear
        ThisWorkbook.Sheets("Output").Cells(1, 1).PasteSpecial xlPasteValues
    End If
    If bestNE <> "" Then
        ThisWorkbook.Sheets(bestNE).Cells.Copy
        ThisWorkbook.Sheets("OutputNE").Cells.Clear
        ThisWorkbook.Sheets("OutputNE").Cells(1, 1).PasteSpecial xlPasteValues
    End If
    Application.CutCopyMode = False
    MsgBox "Restore complete (if backup sheets were found).", vbInformation
End Sub

' --- Macro wrappers for Alt+F8 visibility
Public Sub RemoveDuplicatesAndRecalculate_Run()
    ' Actual run (will create backups and write changes)
    Call RemoveDuplicatesAndRecalculate(False)
End Sub

Public Sub RemoveDuplicatesAndRecalculate_DryRun()
    ' Preview only; does not modify sheets
    Call RemoveDuplicatesAndRecalculate(True)
End Sub

'==========================================================================
'--- UI Macro: Recalculate dashboard and breakdown sheets only ---
'==========================================================================
Public Sub RecalculateProductivityReports()
    Dim previousSkipSetting As Boolean
    Dim metricsStartTime As Double

    previousSkipSetting = SkipReprocessDuringMetrics
    SkipReprocessDuringMetrics = True
    metricsStartTime = Timer

    On Error GoTo CleanUp

    Application.StatusBar = "Recalculating productivity dashboard and breakdowns..."
    CalculateProductivityMetrics metricsStartTime

    Application.StatusBar = False
    SkipReprocessDuringMetrics = previousSkipSetting

    MsgBox "Productivity dashboard, daily, weekly, and monthly breakdowns have been refreshed.", _
           vbInformation, "Recalculation Complete"
    Exit Sub

CleanUp:
    Application.StatusBar = False
    SkipReprocessDuringMetrics = previousSkipSetting
    MsgBox "An error occurred while recalculating productivity reports: " & Err.Description, _
           vbCritical, "Recalculation Error"
End Sub

'==========================================================================
'--- Rebuild Output/OutputNE for a date or date range (with optional import)
'--- Usage: Call RebuildOutputForDateRange(CDate("2025-09-01"), CDate("2025-09-03"), True)
'==========================================================================
Public Sub RebuildOutputForDateRange(ByVal startDate As Date, ByVal endDate As Date, Optional ByVal importFromSource As Boolean = False, Optional ByVal sheetChoice As String = "Both")
    Dim wsOut As Worksheet, wsOutNE As Worksheet, wsAudit As Worksheet
    Dim lastRow As Long, r As Long, c As Long
    Dim arr As Variant, writeArr As Variant
    Dim keepArr() As Variant, keepPtr As Long
    Dim ts As String, outBackupName As String, outNEBackupName As String
    Dim d As Date, i As Long
    Dim localSheet As Worksheet, parsedDate As String, sheetDate As Date
    Dim hadOutputFilter As Boolean, hadOutputNEFilter As Boolean
    Dim historicalAHT As Object

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Rebuilding Output sheets for date range..."

    Set wsOut = ThisWorkbook.Sheets("Output")
    Set wsOutNE = ThisWorkbook.Sheets("OutputNE")

    ts = Format(Now, "yyyymmdd_HHMMSS")
    outBackupName = "Output_backup_" & ts
    outNEBackupName = "OutputNE_backup_" & ts

    On Error Resume Next
    If sheetChoice = "Both" Or sheetChoice = "Output" Then
        Set historicalAHT = CreateObject("Scripting.Dictionary")
        On Error Resume Next
        historicalAHT.CompareMode = vbTextCompare
        On Error GoTo 0
        On Error Resume Next
        ThisWorkbook.Sheets("Output").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = outBackupName
        If Err.Number <> 0 Then Err.Clear
    End If
    If sheetChoice = "Both" Or sheetChoice = "OutputNE" Then
        ThisWorkbook.Sheets("OutputNE").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = outNEBackupName
        If Err.Number <> 0 Then Err.Clear
    End If
    On Error GoTo 0

    Set wsAudit = Nothing
    On Error Resume Next
    Set wsAudit = ThisWorkbook.Sheets("Rebuild_Audit")
    On Error GoTo 0
    If wsAudit Is Nothing Then
        Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAudit.Name = "Rebuild_Audit"
    End If

    If wsAudit.Cells(1, 1).Value = "" Then
        Dim headerCols As Long
        headerCols = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
        wsAudit.Cells(1, 1).Value = "Source": wsAudit.Cells(1, 2).Value = "OriginalRow": wsAudit.Cells(1, 3).Value = "RemovedAt"
        wsAudit.Cells(1, 4).Resize(1, headerCols).Value = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, headerCols)).Value
    End If

    ' Optional import from source
    If importFromSource Then
        For d = startDate To endDate
            If Weekday(d, vbMonday) < 6 Then ImportDataForDate d, sheetChoice
        Next d
    End If

    ' --- Build list of dates that WILL be rebuilt (either from import or existing dated sheets) ---
    Dim rebuildDates As Object: Set rebuildDates = CreateObject("Scripting.Dictionary")
    Dim keyDate As String
    If importFromSource Then
        For d = startDate To endDate
            If Weekday(d, vbMonday) < 6 Then rebuildDates(Format(d, "yyyy-mm-dd")) = 1
        Next d
    Else
        ' Scan workbook for dated sheets inside the range
        For Each localSheet In ThisWorkbook.Worksheets
            If localSheet.Name Like "Personal Entry *" And localSheet.Name <> "Personal Entry" Then
                parsedDate = ParseDateFromName(localSheet.Name, "Personal Entry ")
                If parsedDate <> "" Then
                    sheetDate = CDate(parsedDate)
                    If sheetDate >= startDate And sheetDate <= endDate Then rebuildDates(Format(sheetDate, "yyyy-mm-dd")) = 1
                End If
            ElseIf localSheet.Name Like "Non-Entry Hrs *" And localSheet.Name <> "Non-Entry Hrs" Then
                parsedDate = ParseDateFromName(localSheet.Name, "Non-Entry Hrs ")
                If parsedDate <> "" Then
                    sheetDate = CDate(parsedDate)
                    If sheetDate >= startDate And sheetDate <= endDate Then rebuildDates(Format(sheetDate, "yyyy-mm-dd")) = 1
                End If
            End If
        Next localSheet
    End If

    If rebuildDates.Count = 0 Then
        MsgBox "No dated sheets found in the workbook for that range and import is not selected. Nothing will be rebuilt.", vbInformation, "Nothing to Rebuild"
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Confirm which dates will be rebuilt
    Dim sampleList As String: sampleList = ""
    Dim k As Variant, cnt As Long: cnt = 0
    For Each k In rebuildDates.Keys
        cnt = cnt + 1
        If cnt <= 12 Then
            If sampleList <> "" Then sampleList = sampleList & ", "
            sampleList = sampleList & k
        End If
    Next k
    If rebuildDates.Count > 12 Then sampleList = sampleList & ", ... (" & rebuildDates.Count & " total)"
    If MsgBox("This operation will remove rows for the following dates and rebuild them:" & vbNewLine & sampleList & vbNewLine & vbNewLine & "Proceed?", vbOKCancel + vbExclamation, "Confirm Rebuild Dates") <> vbOK Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Exit Sub
    End If

    ' Process Output
    If sheetChoice = "Both" Or sheetChoice = "Output" Then
        lastRow = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row
        If lastRow > 1 Then
            Dim dataLastCol As Long
            dataLastCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
            ' Safety: detect merged header cells which can break range math
            If wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, dataLastCol)).MergeCells Then
                MsgBox "Output header contains merged cells. Aborting rebuild to avoid corruption.", vbExclamation
                GoTo SkipOutputProcessing
            End If

            If wsOut.AutoFilterMode Then
                hadOutputFilter = True
                On Error Resume Next
                If wsOut.FilterMode Then wsOut.ShowAllData
                On Error GoTo 0
                wsOut.AutoFilterMode = False
            End If
            arr = wsOut.Range("A2", wsOut.Cells(lastRow, dataLastCol)).Value
            Dim arrRows As Long, arrCols As Long
            arrRows = IIf(IsArray(arr), UBound(arr, 1), 0)
            arrCols = IIf(IsArray(arr), UBound(arr, 2), 0)

            ' Validate dimensions
            If arrCols <> dataLastCol Then
                Debug.Print "Output: headerCols=" & dataLastCol & " arrCols=" & arrCols & " -- aborting to avoid miswrite"
                MsgBox "Output sheet column count does not match data range. Aborting rebuild.", vbExclamation
                GoTo SkipOutputProcessing
            End If

            ReDim keepArr(1 To arrRows, 1 To arrCols)
            keepPtr = 0
            For i = 1 To arrRows
                If IsDate(arr(i, 1)) Then
                    Dim thisDateStr As String: thisDateStr = Format(CDate(arr(i, 1)), "yyyy-mm-dd")
                    If rebuildDates.Exists(thisDateStr) Then
                        Dim nr As Long: nr = wsAudit.Cells(wsAudit.Rows.Count, "A").End(xlUp).Row + 1
                        wsAudit.Cells(nr, 1).Value = "Output": wsAudit.Cells(nr, 2).Value = i + 1: wsAudit.Cells(nr, 3).Value = Now
                        For c = 1 To arrCols: wsAudit.Cells(nr, 3 + c).Value = arr(i, c): Next c
                        If Not historicalAHT Is Nothing Then
                            Dim histKeyOut As String
                            histKeyOut = BuildOutputRowKey(arr(i, 1), arr(i, 2), arr(i, 3), arr(i, 4))
                            If Not historicalAHT.Exists(histKeyOut) Then historicalAHT(histKeyOut) = arr(i, 6)
                        End If
                    Else
                        keepPtr = keepPtr + 1
                        For c = 1 To arrCols: keepArr(keepPtr, c) = arr(i, c): Next c
                    End If
                Else
                    keepPtr = keepPtr + 1
                    For c = 1 To arrCols: keepArr(keepPtr, c) = arr(i, c): Next c
                End If
            Next i

            ' Clear only the exact original range (using dataLastCol)
            wsOut.Range("A2", wsOut.Cells(lastRow, dataLastCol)).ClearContents
            If keepPtr > 0 Then
                ReDim writeArr(1 To keepPtr, 1 To arrCols)
                For i = 1 To keepPtr: For c = 1 To arrCols: writeArr(i, c) = keepArr(i, c): Next c: Next i
                wsOut.Cells(2, 1).Resize(keepPtr, arrCols).Value = writeArr
            End If
            If hadOutputFilter Then wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, dataLastCol)).AutoFilter
        End If
    End If
SkipOutputProcessing:

    ' Process OutputNE
    If sheetChoice = "Both" Or sheetChoice = "OutputNE" Then
        lastRow = wsOutNE.Cells(wsOutNE.Rows.Count, "A").End(xlUp).Row
        If lastRow > 1 Then
            Dim dataLastColNE As Long
            dataLastColNE = wsOutNE.Cells(1, wsOutNE.Columns.Count).End(xlToLeft).Column
            If wsOutNE.Range(wsOutNE.Cells(1, 1), wsOutNE.Cells(1, dataLastColNE)).MergeCells Then
                MsgBox "OutputNE header contains merged cells. Aborting rebuild to avoid corruption.", vbExclamation
                GoTo SkipOutputNEProcessing
            End If

            If wsOutNE.AutoFilterMode Then
                hadOutputNEFilter = True
                On Error Resume Next
                If wsOutNE.FilterMode Then wsOutNE.ShowAllData
                On Error GoTo 0
                wsOutNE.AutoFilterMode = False
            End If
            arr = wsOutNE.Range("A2", wsOutNE.Cells(lastRow, dataLastColNE)).Value
            Dim arrRowsNE As Long, arrColsNE As Long
            arrRowsNE = IIf(IsArray(arr), UBound(arr, 1), 0)
            arrColsNE = IIf(IsArray(arr), UBound(arr, 2), 0)

            If arrColsNE <> dataLastColNE Then
                Debug.Print "OutputNE: headerCols=" & dataLastColNE & " arrCols=" & arrColsNE & " -- aborting to avoid miswrite"
                MsgBox "OutputNE sheet column count does not match data range. Aborting rebuild.", vbExclamation
                GoTo SkipOutputNEProcessing
            End If

            ReDim keepArr(1 To arrRowsNE, 1 To arrColsNE)
            keepPtr = 0
            For i = 1 To arrRowsNE
                If IsDate(arr(i, 1)) Then
                    Dim thisDateStrNE As String: thisDateStrNE = Format(CDate(arr(i, 1)), "yyyy-mm-dd")
                    If rebuildDates.Exists(thisDateStrNE) Then
                        Dim nr2 As Long: nr2 = wsAudit.Cells(wsAudit.Rows.Count, "A").End(xlUp).Row + 1
                        wsAudit.Cells(nr2, 1).Value = "OutputNE": wsAudit.Cells(nr2, 2).Value = i + 1: wsAudit.Cells(nr2, 3).Value = Now
                        For c = 1 To arrColsNE: wsAudit.Cells(nr2, 3 + c).Value = arr(i, c): Next c
                    Else
                        keepPtr = keepPtr + 1
                        For c = 1 To arrColsNE: keepArr(keepPtr, c) = arr(i, c): Next c
                    End If
                Else
                    keepPtr = keepPtr + 1
                    For c = 1 To arrColsNE: keepArr(keepPtr, c) = arr(i, c): Next c
                End If
            Next i

            wsOutNE.Range("A2", wsOutNE.Cells(lastRow, dataLastColNE)).ClearContents
            If keepPtr > 0 Then
                ReDim writeArr(1 To keepPtr, 1 To arrColsNE)
                For i = 1 To keepPtr: For c = 1 To arrColsNE: writeArr(i, c) = keepArr(i, c): Next c: Next i
                wsOutNE.Cells(2, 1).Resize(keepPtr, arrColsNE).Value = writeArr
            End If
            If hadOutputNEFilter Then wsOutNE.Range(wsOutNE.Cells(1, 1), wsOutNE.Cells(1, dataLastColNE)).AutoFilter
        End If
    End If
SkipOutputNEProcessing:

    ' Reprocess dated sheets in range
    For Each localSheet In ThisWorkbook.Worksheets
        If (sheetChoice = "Both" Or sheetChoice = "Output") Then
            If localSheet.Name Like "Personal Entry *" And localSheet.Name <> "Personal Entry" Then
                parsedDate = ParseDateFromName(localSheet.Name, "Personal Entry ")
                If parsedDate <> "" Then
                    sheetDate = CDate(parsedDate)
                    If sheetDate >= startDate And sheetDate <= endDate Then
                        Call ProcessActivitySheet(localSheet, parsedDate, historicalAHT)
                    End If
                End If
            End If
        End If
        If (sheetChoice = "Both" Or sheetChoice = "OutputNE") Then
            If localSheet.Name Like "Non-Entry Hrs *" And localSheet.Name <> "Non-Entry Hrs" Then
                parsedDate = ParseDateFromName(localSheet.Name, "Non-Entry Hrs ")
                If parsedDate <> "" Then
                    sheetDate = CDate(parsedDate)
                    If sheetDate >= startDate And sheetDate <= endDate Then Call ProcessNonEntrySheet(localSheet, parsedDate)
                End If
            End If
        End If
    Next localSheet

    ' Sort and recalc
    If sheetChoice = "Both" Or sheetChoice = "Output" Then Call SortSheetByDate(wsOut, 1)
    If sheetChoice = "Both" Or sheetChoice = "OutputNE" Then Call SortSheetByDate(wsOutNE, 1)

    ' Prepare array of dates that were rebuilt so metrics can process only those dates
    Dim rebuiltDatesArr() As Date
    Dim idx As Long: idx = 0
    ReDim rebuiltDatesArr(0 To rebuildDates.Count - 1)
    Dim kd As Variant
    For Each kd In rebuildDates.Keys
        rebuiltDatesArr(idx) = CDate(kd)
        idx = idx + 1
    Next kd
    ' We already rebuilt the dated sheets above, so skip reprocessing inside the metrics routine
    Dim previousSkipFlag As Boolean
    previousSkipFlag = SkipReprocessDuringMetrics
    SkipReprocessDuringMetrics = True
    If rebuildDates.Count > 0 Then
        Call CalculateProductivityMetrics(Timer, rebuiltDatesArr, rebuildDates.Count)
    Else
        Call CalculateProductivityMetrics(Timer)
    End If
    SkipReprocessDuringMetrics = previousSkipFlag

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Rebuild complete. Audit sheet 'Rebuild_Audit' contains removed rows.", vbInformation, "Rebuild Complete"
End Sub

'==========================================================================
'--- Helper: Sort a sheet by a date column (assumes header in row 1)
'==========================================================================
Private Sub SortSheetByDate(ByRef ws As Worksheet, ByVal dateCol As Long)
    Dim lastRow As Long, lastCol As Long
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow > 1 Then
        ws.Sort.SortFields.Clear
        ws.Sort.SortFields.Add Key:=ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ws.Sort
            .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If
    On Error GoTo 0
End Sub

'==========================================================================
'--- UI wrapper so the rebuild macro appears in the Macros dialog (no args)
'==========================================================================
Public Sub RebuildOutputForDateRange_UI()
    Dim s As String, e As String
    Dim sd As Date, ed As Date
    Dim ans As VbMsgBoxResult

    s = InputBox("Enter start date (e.g. 2025-09-01 or 9/1/2025):", "Rebuild - Start Date")
    If s = "" Then Exit Sub
    On Error Resume Next
    sd = CDate(s)
    If Err.Number <> 0 Then MsgBox "Invalid start date entered.": Err.Clear: Exit Sub
    On Error GoTo 0

    e = InputBox("Enter end date (e.g. 2025-09-03 or 9/3/2025):", "Rebuild - End Date", Format(sd, "yyyy-mm-dd"))
    If e = "" Then Exit Sub
    On Error Resume Next
    ed = CDate(e)
    If Err.Number <> 0 Then MsgBox "Invalid end date entered.": Err.Clear: Exit Sub
    On Error GoTo 0

    If ed < sd Then MsgBox "End date must be the same or after the start date.": Exit Sub

    ans = MsgBox("Import dated source sheets from the source workbook before rebuilding?" & vbNewLine & vbNewLine & _
                 "Yes = attempt import; No = rebuild from sheets already in workbook", vbYesNoCancel + vbQuestion, "Import Option")
    If ans = vbCancel Then Exit Sub

    ' Ask which sheets to process
    Dim shChoice As String
    Dim shAns As VbMsgBoxResult
    shAns = MsgBox("Which sheets to rebuild?" & vbNewLine & vbNewLine & "Yes = Both; No = Output only; Cancel = OutputNE only", vbYesNoCancel + vbQuestion, "Sheet Choice")
    If shAns = vbYes Then
        shChoice = "Both"
    ElseIf shAns = vbNo Then
        shChoice = "Output"
    Else
        shChoice = "OutputNE"
    End If

    Call RebuildOutputForDateRange(sd, ed, (ans = vbYes), shChoice)
End Sub

' Create a physical workbook copy (saved alongside the workbook) for safe recovery
Private Sub SaveWorkbookBackup()
    On Error Resume Next
    Dim ts As String: ts = Format(Now, "yyyyMMdd_HHmmss")
    Dim baseName As String: baseName = ThisWorkbook.Name
    Dim dotPos As Long: dotPos = InStrRev(baseName, ".")
    Dim nameOnly As String
    If dotPos > 0 Then nameOnly = Left(baseName, dotPos - 1) Else nameOnly = baseName
    Dim backupName As String
    If ThisWorkbook.Path = "" Then
        backupName = Environ("USERPROFILE") & "\Desktop\" & nameOnly & "_wbbackup_" & ts & ".xlsm"
    Else
        backupName = ThisWorkbook.Path & "\" & nameOnly & "_wbbackup_" & ts & ".xlsm"
    End If
    ThisWorkbook.SaveCopyAs backupName
    If Err.Number = 0 Then Debug.Print "Workbook backup saved to: " & backupName Else Debug.Print "Failed to save workbook backup: " & Err.Number
    On Error GoTo 0
End Sub

'==========================================================================
' --- AUDIT: LIST DATES THAT EXIST IN ONLY ONE OUTPUT SHEET ---
'==========================================================================
Public Sub ReportOutputDateMismatches()
    Const REPORT_SHEET_NAME As String = "Output Date Audit"

    Dim wsOutput As Worksheet
    Dim wsOutputNE As Worksheet
    Dim wsReport As Worksheet
    Dim dictOutput As Object
    Dim dictOutputNE As Object
    Dim results() As Variant
    Dim key As Variant
    Dim statusMessage As String
    Dim mismatches As Collection
    Dim rowIndex As Long
    Dim entry As Variant

    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("Output")
    Set wsOutputNE = ThisWorkbook.Sheets("OutputNE")
    On Error GoTo 0

    If wsOutput Is Nothing Or wsOutputNE Is Nothing Then
        MsgBox "Required sheets 'Output' and 'OutputNE' were not found.", vbCritical, "Output Date Audit"
        Exit Sub
    End If

    Set dictOutput = CreateObject("Scripting.Dictionary")
    Set dictOutputNE = CreateObject("Scripting.Dictionary")

    Call LoadUniqueOutputDates(wsOutput, dictOutput)
    Call LoadUniqueOutputDates(wsOutputNE, dictOutputNE)

    Set mismatches = New Collection

    For Each key In dictOutput.Keys
        If Not dictOutputNE.Exists(key) Then
            mismatches.Add Array(DateValue(CStr(key)), "Output only")
        End If
    Next key

    For Each key In dictOutputNE.Keys
        If Not dictOutput.Exists(key) Then
            mismatches.Add Array(DateValue(CStr(key)), "OutputNE only")
        End If
    Next key

    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets(REPORT_SHEET_NAME)
    On Error GoTo 0

    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsReport.Name = REPORT_SHEET_NAME
    Else
        wsReport.Cells.Clear
    End If

    wsReport.Range("A1").Resize(1, 2).Value = Array("Date", "Present In")

    If mismatches.Count > 0 Then
        ReDim results(1 To mismatches.Count, 1 To 2)

        For rowIndex = 1 To mismatches.Count
            entry = mismatches(rowIndex)
            results(rowIndex, 1) = entry(0)
            results(rowIndex, 2) = entry(1)
        Next rowIndex

        wsReport.Range("A2").Resize(mismatches.Count, 2).Value = results

        wsReport.Sort.SortFields.Clear
        wsReport.Sort.SortFields.Add Key:=wsReport.Range("A2:A" & (mismatches.Count + 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With wsReport.Sort
            .SetRange wsReport.Range("A1:B" & (mismatches.Count + 1))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With

        statusMessage = "Found " & mismatches.Count & " date(s) with mismatched output coverage." & vbNewLine & _
                        "See '" & REPORT_SHEET_NAME & "' for details."
    Else
        wsReport.Range("A2").Value = "All dates are present in both Output sheets."
        wsReport.Range("B2").Value = "--"
        statusMessage = "Output and OutputNE contain the same set of dates." & vbNewLine & _
                        "A confirmation note was written to '" & REPORT_SHEET_NAME & "'."
    End If

    wsReport.Columns("A:B").AutoFit
    wsReport.Activate
    wsReport.Range("A1").Select

    MsgBox statusMessage, vbInformation, "Output Date Audit"
End Sub

Private Sub LoadUniqueOutputDates(ByVal ws As Worksheet, ByVal dict As Object)
    Dim lastRow As Long
    Dim arr As Variant
    Dim i As Long
    Dim cellValue As Variant
    Dim dateKey As String

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow <= 1 Then Exit Sub

    arr = ws.Range("A2:A" & lastRow).Value

    For i = 1 To UBound(arr, 1)
        cellValue = arr(i, 1)
        If Not IsEmpty(cellValue) Then
            If IsDate(cellValue) Then
                dateKey = Format$(CDate(cellValue), "yyyy-mm-dd")
                If Not dict.Exists(dateKey) Then dict.Add dateKey, True
            ElseIf IsNumeric(cellValue) Then
                If CDbl(cellValue) > 0 Then
                    dateKey = Format$(DateSerial(1899, 12, 31) + CDbl(cellValue), "yyyy-mm-dd")
                    If Not dict.Exists(dateKey) Then dict.Add dateKey, True
                End If
            End If
        End If
    Next i
End Sub

'==========================================================================
'--- Build a detail report for a specific person/date across Output sheets ---
'==========================================================================
Public Sub ShowPersonDayProductivityDetail()
    Const REPORT_SHEET_NAME As String = "PersonDayDetail"

    Dim personInput As String
    Dim dateInput As String
    Dim targetDate As Date
    Dim targetKey As String
    Dim originalScreenUpdating As Boolean

    Dim wsOutput As Worksheet
    Dim wsOutputNE As Worksheet
    Dim wsReport As Worksheet

    Dim outputLastRow As Long, outputLastCol As Long
    Dim outputHeaders As Variant, outputData As Variant
    Dim outputMatches As Collection
    Dim outputDateIdx As Long, outputNameIdx As Long
    Dim outputCountIdx As Long, outputProdIdx As Long
    Dim outputRowCount As Long
    Dim outputTotalCount As Double
    Dim outputTotalProdHours As Double

    Dim neLastRow As Long, neLastCol As Long
    Dim neHeaders As Variant, neData As Variant
    Dim neMatches As Collection
    Dim neDateIdx As Long, neNameIdx As Long
    Dim neCountIdx As Long
    Dim neRowCount As Long
    Dim neTotalCount As Double

    Dim i As Long, c As Long
    Dim rowData As Variant
    Dim rowPtr As Long
    Dim rawDate As Variant, thisDate As Date, isDateMatch As Boolean, thisName As String
    Dim neRawDate As Variant, neThisDate As Date, neDateMatch As Boolean, neName As String

    personInput = Trim$(InputBox("Enter the team member name (as shown on the Output sheets):", _
                                 "Person-Day Detail"))
    If Len(personInput) = 0 Then Exit Sub

    dateInput = Trim$(InputBox("Enter the work date (e.g. 2025-02-15 or 2/15/2025):", _
                                "Person-Day Detail", Format(Date, "yyyy-mm-dd")))
    If Len(dateInput) = 0 Then Exit Sub

    On Error Resume Next
    targetDate = CDate(dateInput)
    If Err.Number <> 0 Then
        MsgBox "The date provided could not be interpreted. Please try again.", vbExclamation, "Person-Day Detail"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    targetKey = Format$(targetDate, "yyyy-mm-dd")

    originalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error GoTo CleanUp

    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("Output")
    Set wsOutputNE = ThisWorkbook.Sheets("OutputNE")
    On Error GoTo 0
    On Error GoTo CleanUp

    If wsOutput Is Nothing Or wsOutputNE Is Nothing Then
        Application.ScreenUpdating = originalScreenUpdating
        MsgBox "Required sheets 'Output' and 'OutputNE' were not found.", vbCritical, "Person-Day Detail"
        Exit Sub
    End If

    Set outputMatches = New Collection
    outputTotalCount = 0
    outputTotalProdHours = 0
    outputRowCount = 0

    outputLastRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row
    outputLastCol = wsOutput.Cells(1, wsOutput.Columns.Count).End(xlToLeft).Column

    If outputLastRow > 1 And outputLastCol >= 2 Then
        outputHeaders = wsOutput.Range(wsOutput.Cells(1, 1), wsOutput.Cells(1, outputLastCol)).Value
        outputData = wsOutput.Range(wsOutput.Cells(2, 1), wsOutput.Cells(outputLastRow, outputLastCol)).Value

        For c = 1 To outputLastCol
            If StrComp(CStr(outputHeaders(1, c)), "Date", vbTextCompare) = 0 Then outputDateIdx = c
            If StrComp(CStr(outputHeaders(1, c)), "Name", vbTextCompare) = 0 Then outputNameIdx = c
            If StrComp(CStr(outputHeaders(1, c)), "Count", vbTextCompare) = 0 Then outputCountIdx = c
            If StrComp(CStr(outputHeaders(1, c)), "Productive Hours", vbTextCompare) = 0 Then outputProdIdx = c
        Next c

        If outputDateIdx > 0 And outputNameIdx > 0 Then
            For i = 1 To UBound(outputData, 1)
                rawDate = outputData(i, outputDateIdx)
                isDateMatch = False
                If IsDate(rawDate) Then
                    thisDate = CDate(rawDate)
                    isDateMatch = (Format$(thisDate, "yyyy-mm-dd") = targetKey)
                ElseIf IsNumeric(rawDate) Then
                    If CDbl(rawDate) > 0 Then
                        thisDate = DateSerial(1899, 12, 31) + CDbl(rawDate)
                        isDateMatch = (Format$(thisDate, "yyyy-mm-dd") = targetKey)
                    End If
                End If

                thisName = Trim$(CStr(outputData(i, outputNameIdx)))

                If isDateMatch And StrComp(thisName, personInput, vbTextCompare) = 0 Then
                    ReDim rowData(1 To outputLastCol)
                    For c = 1 To outputLastCol
                        rowData(c) = outputData(i, c)
                    Next c
                    outputMatches.Add rowData
                    outputRowCount = outputRowCount + 1

                    If outputCountIdx > 0 Then
                        If IsNumeric(outputData(i, outputCountIdx)) Then
                            outputTotalCount = outputTotalCount + CDbl(outputData(i, outputCountIdx))
                        End If
                    End If
                    If outputProdIdx > 0 Then
                        If IsNumeric(outputData(i, outputProdIdx)) Then
                            outputTotalProdHours = outputTotalProdHours + CDbl(outputData(i, outputProdIdx))
                        End If
                    End If
                End If
            Next i
        End If
    End If

    Set neMatches = New Collection
    neTotalCount = 0
    neRowCount = 0

    neLastRow = wsOutputNE.Cells(wsOutputNE.Rows.Count, 1).End(xlUp).Row
    neLastCol = wsOutputNE.Cells(1, wsOutputNE.Columns.Count).End(xlToLeft).Column

    If neLastRow > 1 And neLastCol >= 2 Then
        neHeaders = wsOutputNE.Range(wsOutputNE.Cells(1, 1), wsOutputNE.Cells(1, neLastCol)).Value
        neData = wsOutputNE.Range(wsOutputNE.Cells(2, 1), wsOutputNE.Cells(neLastRow, neLastCol)).Value

        For c = 1 To neLastCol
            If StrComp(CStr(neHeaders(1, c)), "Date", vbTextCompare) = 0 Then neDateIdx = c
            If StrComp(CStr(neHeaders(1, c)), "Name", vbTextCompare) = 0 Then neNameIdx = c
            If StrComp(CStr(neHeaders(1, c)), "Count", vbTextCompare) = 0 Then neCountIdx = c
        Next c

        If neDateIdx > 0 And neNameIdx > 0 Then
            For i = 1 To UBound(neData, 1)
                neRawDate = neData(i, neDateIdx)
                neDateMatch = False
                If IsDate(neRawDate) Then
                    neThisDate = CDate(neRawDate)
                    neDateMatch = (Format$(neThisDate, "yyyy-mm-dd") = targetKey)
                ElseIf IsNumeric(neRawDate) Then
                    If CDbl(neRawDate) > 0 Then
                        neThisDate = DateSerial(1899, 12, 31) + CDbl(neRawDate)
                        neDateMatch = (Format$(neThisDate, "yyyy-mm-dd") = targetKey)
                    End If
                End If

                neName = Trim$(CStr(neData(i, neNameIdx)))

                If neDateMatch And StrComp(neName, personInput, vbTextCompare) = 0 Then
                    ReDim rowData(1 To neLastCol)
                    For c = 1 To neLastCol
                        rowData(c) = neData(i, c)
                    Next c
                    neMatches.Add rowData
                    neRowCount = neRowCount + 1

                    If neCountIdx > 0 Then
                        If IsNumeric(neData(i, neCountIdx)) Then
                            neTotalCount = neTotalCount + CDbl(neData(i, neCountIdx))
                        End If
                    End If
                End If
            Next i
        End If
    End If

    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets(REPORT_SHEET_NAME)
    On Error GoTo 0
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsReport.Name = REPORT_SHEET_NAME
    Else
        wsReport.Cells.Clear
    End If

    On Error GoTo CleanUp

    wsReport.Range("A1").Value = "Person-Day Detail"
    wsReport.Range("A1").Font.Bold = True
    wsReport.Range("A1").Font.Size = 14

    wsReport.Range("A3").Value = "Person:"
    wsReport.Range("B3").Value = personInput
    wsReport.Range("A4").Value = "Date:"
    wsReport.Range("B4").Value = targetDate
    wsReport.Range("B4").NumberFormat = "yyyy-mm-dd"

    wsReport.Range("A5").Value = "Output entries found:"
    wsReport.Range("B5").Value = outputRowCount
    wsReport.Range("A6").Value = "Total productive hours:"
    wsReport.Range("B6").Value = outputTotalProdHours
    wsReport.Range("B6").NumberFormat = "0.00"
    wsReport.Range("A7").Value = "Total entry count:"
    wsReport.Range("B7").Value = outputTotalCount
    wsReport.Range("A8").Value = "Non-entry records found:"
    wsReport.Range("B8").Value = neRowCount
    wsReport.Range("A9").Value = "Total non-entry count:"
    wsReport.Range("B9").Value = neTotalCount

    rowPtr = 11

    wsReport.Range("A" & rowPtr).Value = "Output sheet details"
    wsReport.Range("A" & rowPtr).Font.Bold = True
    rowPtr = rowPtr + 1

    If outputMatches.Count > 0 Then
        wsReport.Cells(rowPtr, 1).Resize(1, outputLastCol).Value = outputHeaders
        Dim outWrite() As Variant
        ReDim outWrite(1 To outputMatches.Count, 1 To outputLastCol)
        For i = 1 To outputMatches.Count
            rowData = outputMatches(i)
            For c = 1 To outputLastCol
                outWrite(i, c) = rowData(c)
            Next c
        Next i
        wsReport.Cells(rowPtr + 1, 1).Resize(outputMatches.Count, outputLastCol).Value = outWrite
        rowPtr = rowPtr + outputMatches.Count + 2
    Else
        wsReport.Cells(rowPtr, 1).Value = "No Output records found for the selected person/date."
        rowPtr = rowPtr + 2
    End If

    rowPtr = rowPtr + 1
    wsReport.Range("A" & rowPtr).Value = "OutputNE sheet details"
    wsReport.Range("A" & rowPtr).Font.Bold = True
    rowPtr = rowPtr + 1

    If neMatches.Count > 0 Then
        wsReport.Cells(rowPtr, 1).Resize(1, neLastCol).Value = neHeaders
        Dim neWrite() As Variant
        ReDim neWrite(1 To neMatches.Count, 1 To neLastCol)
        For i = 1 To neMatches.Count
            rowData = neMatches(i)
            For c = 1 To neLastCol
                neWrite(i, c) = rowData(c)
            Next c
        Next i
        wsReport.Cells(rowPtr + 1, 1).Resize(neMatches.Count, neLastCol).Value = neWrite
        rowPtr = rowPtr + neMatches.Count + 2
    Else
        wsReport.Cells(rowPtr, 1).Value = "No OutputNE records found for the selected person/date."
        rowPtr = rowPtr + 2
    End If

    wsReport.Columns.AutoFit
    wsReport.Activate
    wsReport.Range("A1").Select

    Application.ScreenUpdating = originalScreenUpdating

    If outputMatches.Count = 0 And neMatches.Count = 0 Then
        MsgBox "No Output or OutputNE records were found for " & personInput & " on " & targetKey & ".", _
               vbInformation, "Person-Day Detail"
    Else
        MsgBox "Person-day detail report created on the '" & REPORT_SHEET_NAME & "' sheet.", _
               vbInformation, "Person-Day Detail"
    End If

    Exit Sub

CleanUp:
    Application.ScreenUpdating = originalScreenUpdating
    If Err.Number <> 0 Then
        MsgBox "An unexpected error occurred while building the person-day detail report: " & Err.Description, _
               vbCritical, "Person-Day Detail"
    End If
End Sub
