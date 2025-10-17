'==========================================================================
' Module: DataProcessing
'==========================================================================
Option Explicit

' Safely convert mixed-format entry counts (handles commas, stray text, and non-breaking spaces)
Private Function ParseEntryCount(ByVal rawValue As Variant) As Double
    Dim cleaned As String, ch As String, i As Long, buffer As String
    If IsError(rawValue) Then Exit Function
    If IsNumeric(rawValue) Then
        ParseEntryCount = CDbl(rawValue)
        Exit Function
    End If
    cleaned = CStr(rawValue)
    cleaned = Replace(cleaned, Chr$(160), " ")
    cleaned = Trim$(cleaned)
    If Len(cleaned) = 0 Then Exit Function
    cleaned = Replace(cleaned, ",", "")
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If ch >= "0" And ch <= "9" Then
            buffer = buffer & ch
        ElseIf ch = "." Or ch = "-" Then
            buffer = buffer & ch
        ElseIf Len(buffer) > 0 Then
            Exit For
        End If
    Next i
    buffer = Trim$(buffer)
    If buffer = "" Or buffer = "-" Or buffer = "." Then Exit Function
    If IsNumeric(buffer) Then ParseEntryCount = CDbl(buffer)
End Function
'==========================================================================
' --- Process "Personal Entry" Sheet ---
'==========================================================================
Public Sub ProcessActivitySheet(wsInput As Worksheet, theDate As String, Optional historicalData As Object = Nothing)

    Const FIRST_TASK_ROW As Long = 2
    Const FIRST_DATA_ROW As Long = 3
    Const FIRST_TASK_COL As Long = 2
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Sheets("Output")
    Dim wsLookup As Worksheet: Set wsLookup = ThisWorkbook.Sheets("ActivityLookup")

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim lkArr As Variant, lastLkRow As Long
    lastLkRow = wsLookup.Cells(wsLookup.Rows.Count, 1).End(xlUp).row
    lkArr = wsLookup.Range("A2:C" & lastLkRow).Value
    Dim r As Long
    For r = 1 To UBound(lkArr, 1): dict(lkArr(r, 1)) = lkArr(r, 2): Next r
    
    Dim lastRow As Long, lastCol As Long
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    lastCol = wsInput.Cells(FIRST_TASK_ROW, wsInput.Columns.Count).End(xlToLeft).Column
    
    Dim inArr As Variant
    inArr = wsInput.Range(wsInput.Cells(1, 1), wsInput.Cells(lastRow, lastCol)).Value
    
    Dim outArr() As Variant
    ReDim outArr(1 To (lastRow - FIRST_DATA_ROW + 1) * (lastCol - FIRST_TASK_COL + 1), 1 To 7)
    
    Dim outPtr As Long: outPtr = 1
    Dim i As Long, j As Long, entryCount As Double, taskName As String, region As String, taskOnly As String
    Dim aht As Variant, prodHrs As Variant, missingDict As Object: Set missingDict = CreateObject("Scripting.Dictionary")
    Const VALID_REGIONS As String = ",BC,AB,CT,ON,QC,MT,YK,"

    For i = FIRST_DATA_ROW To lastRow
        For j = FIRST_TASK_COL To lastCol
            entryCount = ParseEntryCount(inArr(i, j))
            If entryCount > 0 Then
                taskName = inArr(FIRST_TASK_ROW, j)
                Dim cand As String: cand = Split(taskName, " ")(0)
                If InStr(1, VALID_REGIONS, "," & cand & ",", vbTextCompare) > 0 Then
                    region = cand: taskOnly = Mid(taskName, Len(region) + 2)
                Else
                    region = "AR": taskOnly = taskName
                End If
                Dim histKey As String
                If Not historicalData Is Nothing Then
                    histKey = BuildOutputRowKey(theDate, CStr(inArr(i, 1)), region, taskOnly)
                    If historicalData.Exists(histKey) Then
                        aht = historicalData(histKey)
                    ElseIf dict.Exists(taskName) Then
                        aht = dict(taskName)
                    Else
                        aht = "N/A"
                    End If
                Else
                    If dict.Exists(taskName) Then aht = dict(taskName) Else aht = "N/A"
                End If
                ' Clean any errors that might come from the lookup
                If IsError(aht) Then aht = "N/A"

                If IsNumeric(aht) Then
                    prodHrs = entryCount * aht / 60
                Else
                    prodHrs = "N/A"
                End If

                outArr(outPtr, 1) = theDate: outArr(outPtr, 2) = inArr(i, 1): outArr(outPtr, 3) = region
                outArr(outPtr, 4) = taskOnly: outArr(outPtr, 5) = entryCount: outArr(outPtr, 6) = aht
                outArr(outPtr, 7) = prodHrs: outPtr = outPtr + 1
            End If
        Next j
    Next i
    
    If outPtr = 1 Then Exit Sub
    
    Dim lastOutRow As Long
    lastOutRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).row
    If lastOutRow = 1 And wsOutput.Cells(1, 1).Value = "" Then
        wsOutput.Range("A1").Resize(1, 7).Value = Array("Date", "Name", "Region", "Task", "Count", "Avg Handle (min)", "Productive Hours")
        lastOutRow = 1
    End If
    wsOutput.Cells(lastOutRow + 1, 1).Resize(outPtr - 1, 7).Value = outArr
    
    ' *** NEW: Apply AutoFilter ***
    If wsOutput.AutoFilterMode Then wsOutput.AutoFilterMode = False
    wsOutput.Range("A1").AutoFilter
    
End Sub

'==========================================================================
' --- Process "Non-Entry Hrs" Sheet ---
'==========================================================================
Public Sub ProcessNonEntrySheet(wsInput As Worksheet, theDate As String)
    Const FIRST_HEADER_COL As Long = 4, LAST_HEADER_COL As Long = 19
    Const FIRST_NAME_ROW As Long = 2, NAME_COL As Long = 1, HEADER_ROW As Long = 1
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Sheets("OutputNE")
    
    Dim lastRow As Long
    lastRow = wsInput.Cells(wsInput.Rows.Count, NAME_COL).End(xlUp).row
    If lastRow < FIRST_NAME_ROW Then Exit Sub
    
    Dim inArr As Variant
    inArr = wsInput.Range(wsInput.Cells(HEADER_ROW, NAME_COL), wsInput.Cells(lastRow, LAST_HEADER_COL)).Value
    
    Dim outArr() As Variant
    ReDim outArr(1 To (lastRow - FIRST_NAME_ROW + 1) * (LAST_HEADER_COL - FIRST_HEADER_COL + 1), 1 To 4)
    Dim outPtr As Long: outPtr = 1
    Dim i As Long, j As Long, taskName As String, countVal As Variant
    
    For i = FIRST_NAME_ROW To lastRow
        For j = FIRST_HEADER_COL To LAST_HEADER_COL
            countVal = inArr(i, j)
            If IsNumeric(countVal) And countVal > 0 Then
                taskName = inArr(HEADER_ROW, j)
                taskName = Replace(taskName, vbLf, " "): taskName = Application.Trim(Replace(taskName, "  ", " "))
                outArr(outPtr, 1) = theDate: outArr(outPtr, 2) = inArr(i, NAME_COL)
                outArr(outPtr, 3) = taskName: outArr(outPtr, 4) = countVal
                outPtr = outPtr + 1
            End If
        Next j
    Next i
    
    If outPtr = 1 Then Exit Sub
    
    Dim lastOutRow As Long
    lastOutRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).row
    If lastOutRow = 1 And wsOutput.Cells(1, 1).Value = "" Then
        wsOutput.Range("A1").Resize(1, 4).Value = Array("Date", "Name", "Task", "Count")
        lastOutRow = 1
    End If
    wsOutput.Cells(lastOutRow + 1, 1).Resize(outPtr - 1, 4).Value = outArr
    
    ' *** NEW: Apply AutoFilter ***
    If wsOutput.AutoFilterMode Then wsOutput.AutoFilterMode = False
    wsOutput.Range("A1").AutoFilter
    
End Sub

'==========================================================================
' --- Compare Output Sheets ---
'==========================================================================
Public Sub CompareOutputAndOutputNE()

    Const REPORT_SHEET_NAME As String = "Output vs OutputNE"
    Const COL_DATE As Long = 1
    Const COL_NAME As Long = 2
    Const COL_TASK As Long = 3

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim wsOutput As Worksheet, wsOutputNE As Worksheet
    On Error Resume Next
    Set wsOutput = wb.Worksheets("Output")
    Set wsOutputNE = wb.Worksheets("OutputNE")
    On Error GoTo 0

    If wsOutput Is Nothing Then
        MsgBox "The sheet named 'Output' could not be found.", vbExclamation
        Exit Sub
    End If

    If wsOutputNE Is Nothing Then
        MsgBox "The sheet named 'OutputNE' could not be found.", vbExclamation
        Exit Sub
    End If

    Dim wsReport As Worksheet
    On Error Resume Next
    Set wsReport = wb.Worksheets(REPORT_SHEET_NAME)
    On Error GoTo 0

    If wsReport Is Nothing Then
        Set wsReport = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsReport.Name = REPORT_SHEET_NAME
    Else
        wsReport.Cells.Clear
    End If

    Dim dictOutput As Object, dictOutputNE As Object, dictNEStatus As Object
    Set dictOutput = CreateObject("Scripting.Dictionary")
    Set dictOutputNE = CreateObject("Scripting.Dictionary")
    Set dictNEStatus = CreateObject("Scripting.Dictionary")
    dictOutput.CompareMode = vbTextCompare
    dictOutputNE.CompareMode = vbTextCompare
    dictNEStatus.CompareMode = vbTextCompare

    Dim lastRow As Long, r As Long

    Dim key As Variant, personName As String, dateValue As Variant

    lastRow = wsOutput.Cells(wsOutput.Rows.Count, COL_DATE).End(xlUp).row
    For r = 2 To lastRow
        dateValue = wsOutput.Cells(r, COL_DATE).Value
        personName = Trim$(CStr(wsOutput.Cells(r, COL_NAME).Value))
        If Len(personName) > 0 And Not IsEmpty(dateValue) Then
            key = GetKeyFromDateName(dateValue, personName)
            If Not dictOutput.Exists(key) Then
                Dim outputDetail As Object
                Dim outputTasks As Object
                Set outputDetail = CreateObject("Scripting.Dictionary")
                outputDetail.CompareMode = vbTextCompare
                outputDetail.Add "Date", dateValue
                outputDetail.Add "Name", personName
                outputDetail.Add "TotalProdHours", 0#
                Set outputTasks = CreateObject("Scripting.Dictionary")
                outputTasks.CompareMode = vbTextCompare
                outputDetail.Add "Tasks", outputTasks
                dictOutput.Add key, outputDetail
            End If

            Dim detailOutput As Object
            Dim tasksOutput As Object
            Dim taskInfo As Object
            Dim taskNameOutput As String
            Dim countValOutput As Variant
            Dim avgHandleVal As Variant
            Dim prodHoursVal As Variant

            Set detailOutput = dictOutput(key)
            Set tasksOutput = detailOutput("Tasks")

            prodHoursVal = wsOutput.Cells(r, 7).Value
            If IsNumeric(prodHoursVal) Then
                detailOutput("TotalProdHours") = detailOutput("TotalProdHours") + CDbl(prodHoursVal)
            End If

            taskNameOutput = Trim$(CStr(wsOutput.Cells(r, 4).Value))
            If Len(taskNameOutput) > 0 Then
                If tasksOutput.Exists(taskNameOutput) Then
                    Set taskInfo = tasksOutput(taskNameOutput)
                Else
                    Set taskInfo = CreateObject("Scripting.Dictionary")
                    taskInfo.CompareMode = vbTextCompare
                    taskInfo.Add "Count", 0#
                    taskInfo.Add "HasNumericCount", False
                    taskInfo.Add "CountNotes", ""
                    taskInfo.Add "AvgHandle", ""
                    tasksOutput.Add taskNameOutput, taskInfo
                End If

                countValOutput = wsOutput.Cells(r, 5).Value
                If IsNumeric(countValOutput) Then
                    taskInfo("Count") = taskInfo("Count") + CDbl(countValOutput)
                    taskInfo("HasNumericCount") = True
                Else
                    Dim countNote As String
                    countNote = Trim$(CStr(countValOutput))
                    If Len(countNote) > 0 Then
                        If Len(Trim$(CStr(taskInfo("CountNotes")))) > 0 Then
                            taskInfo("CountNotes") = taskInfo("CountNotes") & "; " & countNote
                        Else
                            taskInfo("CountNotes") = countNote
                        End If
                    End If
                End If

                avgHandleVal = wsOutput.Cells(r, 6).Value
                If IsNumeric(avgHandleVal) Then
                    taskInfo("AvgHandle") = CDbl(avgHandleVal)
                ElseIf Len(Trim$(CStr(avgHandleVal))) > 0 Then
                    taskInfo("AvgHandle") = CStr(avgHandleVal)
                End If
            End If
        End If
    Next r

    lastRow = wsOutputNE.Cells(wsOutputNE.Rows.Count, COL_DATE).End(xlUp).row
    For r = 2 To lastRow
        dateValue = wsOutputNE.Cells(r, COL_DATE).Value
        personName = Trim$(CStr(wsOutputNE.Cells(r, COL_NAME).Value))
        If Len(personName) > 0 And Not IsEmpty(dateValue) Then
            Dim taskName As String
            taskName = Trim$(CStr(wsOutputNE.Cells(r, COL_TASK).Value))
            If Len(taskName) > 0 Then
                key = GetKeyFromDateName(dateValue, personName)

                Dim statusInfo As Variant
                If dictNEStatus.Exists(key) Then
                    statusInfo = dictNEStatus(key)
                Else
                    statusInfo = Array(False, False) ' (hasNonTimeOff, hasAnyEntry)
                End If

                statusInfo(1) = True ' hasAnyEntry

                If Not IsTimeOffTask(taskName) Then
                    statusInfo(0) = True ' hasNonTimeOff
                    If Not dictOutputNE.Exists(key) Then
                        Dim neDetail As Object
                        Dim neTasks As Object
                        Set neDetail = CreateObject("Scripting.Dictionary")
                        neDetail.CompareMode = vbTextCompare
                        neDetail.Add "Date", dateValue
                        neDetail.Add "Name", personName
                        neDetail.Add "TotalProdHours", 0#
                        Set neTasks = CreateObject("Scripting.Dictionary")
                        neTasks.CompareMode = vbTextCompare
                        neDetail.Add "Tasks", neTasks
                        dictOutputNE.Add key, neDetail
                    End If

                    Dim detailNE As Object
                    Dim tasksNE As Object
                    Dim neTaskInfo As Object
                    Dim countValNE As Variant
                    Dim countNoteNE As String

                    Set detailNE = dictOutputNE(key)
                    Set tasksNE = detailNE("Tasks")

                    If tasksNE.Exists(taskName) Then
                        Set neTaskInfo = tasksNE(taskName)
                    Else
                        Set neTaskInfo = CreateObject("Scripting.Dictionary")
                        neTaskInfo.CompareMode = vbTextCompare
                        neTaskInfo.Add "Count", 0#
                        neTaskInfo.Add "HasNumericCount", False
                        neTaskInfo.Add "CountNotes", ""
                        tasksNE.Add taskName, neTaskInfo
                    End If

                    countValNE = wsOutputNE.Cells(r, 4).Value
                    If IsNumeric(countValNE) Then
                        Dim numericCount As Double
                        numericCount = CDbl(countValNE)
                        neTaskInfo("Count") = neTaskInfo("Count") + numericCount
                        neTaskInfo("HasNumericCount") = True
                        detailNE("TotalProdHours") = detailNE("TotalProdHours") + numericCount
                    Else
                        countNoteNE = Trim$(CStr(countValNE))
                        If Len(countNoteNE) > 0 Then
                            If Len(Trim$(CStr(neTaskInfo("CountNotes")))) > 0 Then
                                neTaskInfo("CountNotes") = neTaskInfo("CountNotes") & "; " & countNoteNE
                            Else
                                neTaskInfo("CountNotes") = countNoteNE
                            End If
                        End If
                    End If
                End If

                dictNEStatus(key) = statusInfo
            End If
        End If
    Next r

    Dim resultData As Collection
    Set resultData = New Collection

    Dim arrVal As Variant
    For Each key In dictOutputNE.Keys
        If Not dictOutput.Exists(key) Then
            Dim detailForNE As Object
            Dim totalProd As Variant
            Dim outputTasksText As String
            Dim outputNETasksText As String

            Set detailForNE = dictOutputNE(key)
            totalProd = detailForNE("TotalProdHours")
            outputTasksText = ""
            outputNETasksText = FormatOutputNETaskDetails(detailForNE)

            resultData.Add Array(detailForNE("Date"), detailForNE("Name"), "OutputNE", "Output", totalProd, outputTasksText, outputNETasksText)
        End If
    Next key

    For Each key In dictOutput.Keys
        If Not dictOutputNE.Exists(key) Then
            Dim skipMismatch As Boolean
            If dictNEStatus.Exists(key) Then
                arrVal = dictNEStatus(key)
                If Not arrVal(0) And arrVal(1) Then skipMismatch = True
            End If

            If Not skipMismatch Then
                Dim detailOnlyOutput As Object
                Dim totalProdOutput As Variant
                Dim outputTaskText As String
                Dim outputNETaskText As String

                Set detailOnlyOutput = dictOutput(key)
                totalProdOutput = detailOnlyOutput("TotalProdHours")
                outputTaskText = FormatOutputTaskDetails(detailOnlyOutput)
                outputNETaskText = ""

                resultData.Add Array(detailOnlyOutput("Date"), detailOnlyOutput("Name"), "Output", "OutputNE", totalProdOutput, outputTaskText, outputNETaskText)
            End If
        End If
    Next key

    If resultData.Count = 0 Then
        wsReport.Range("A1:G1").Value = Array("Date", "Name", "Present In", "Missing From", "Total Prod Hours (Output)", "Output Task Details", "OutputNE Task Details")
        wsReport.Range("A2").Value = "No mismatches found."
        Exit Sub
    End If

    Dim results() As Variant
    ReDim results(1 To resultData.Count, 1 To 7)

    Dim idx As Long
    For idx = 1 To resultData.Count
        arrVal = resultData(idx)
        results(idx, 1) = arrVal(0)
        results(idx, 2) = arrVal(1)
        results(idx, 3) = arrVal(2)
        results(idx, 4) = arrVal(3)
        If IsNull(arrVal(4)) Then
            results(idx, 5) = ""
        Else
            results(idx, 5) = arrVal(4)
        End If
        results(idx, 6) = arrVal(5)
        results(idx, 7) = arrVal(6)
    Next idx

    wsReport.Range("A1:G1").Value = Array("Date", "Name", "Present In", "Missing From", "Total Prod Hours (Output)", "Output Task Details", "OutputNE Task Details")
    wsReport.Range("A2").Resize(resultData.Count, 7).Value = results
    wsReport.Columns("A:G").AutoFit

End Sub

Private Function FormatOutputTaskDetails(ByVal detail As Object) As String
    Dim tasks As Object
    On Error Resume Next
    Set tasks = detail("Tasks")
    On Error GoTo 0
    If tasks Is Nothing Then Exit Function
    If tasks.Count = 0 Then Exit Function

    Dim parts() As String
    ReDim parts(0 To tasks.Count - 1)

    Dim idx As Long
    Dim taskName As Variant
    idx = 0
    For Each taskName In tasks.Keys
        Dim taskInfo As Object
        Dim part As String
        Dim countText As String
        Dim notesText As String
        Set taskInfo = tasks(taskName)

        If CBool(taskInfo("HasNumericCount")) Then
            countText = FormatNumberForReport(taskInfo("Count"))
        Else
            countText = ""
        End If

        notesText = Trim$(CStr(taskInfo("CountNotes")))
        If Len(notesText) > 0 Then
            If Len(countText) > 0 Then
                countText = countText & " | Notes: " & notesText
            Else
                countText = notesText
            End If
        End If

        If Len(countText) = 0 Then countText = "N/A"

        part = CStr(taskName) & " (Count: " & countText
        Dim avgVal As Variant
        avgVal = taskInfo("AvgHandle")
        If IsNumeric(avgVal) Then
            part = part & ", Avg: " & FormatNumberForReport(avgVal)
        ElseIf Len(Trim$(CStr(avgVal))) > 0 Then
            part = part & ", Avg: " & CStr(avgVal)
        End If
        part = part & ")"
        parts(idx) = part
        idx = idx + 1
    Next taskName

    FormatOutputTaskDetails = Join(parts, "; ")
End Function

Private Function FormatOutputNETaskDetails(ByVal detail As Object) As String
    Dim tasks As Object
    On Error Resume Next
    Set tasks = detail("Tasks")
    On Error GoTo 0
    If tasks Is Nothing Then Exit Function
    If tasks.Count = 0 Then Exit Function

    Dim parts() As String
    ReDim parts(0 To tasks.Count - 1)

    Dim idx As Long
    Dim taskName As Variant
    idx = 0
    For Each taskName In tasks.Keys
        Dim neTaskInfo As Object
        Dim countText As String
        Dim noteText As String
        Set neTaskInfo = tasks(taskName)

        If CBool(neTaskInfo("HasNumericCount")) Then
            countText = FormatNumberForReport(neTaskInfo("Count"))
        Else
            countText = ""
        End If

        noteText = Trim$(CStr(neTaskInfo("CountNotes")))
        If Len(noteText) > 0 Then
            If Len(countText) > 0 Then
                countText = countText & " | Notes: " & noteText
            Else
                countText = noteText
            End If
        End If

        If Len(countText) = 0 Then countText = "N/A"

        parts(idx) = CStr(taskName) & " (Count: " & countText & ")"
        idx = idx + 1
    Next taskName

    FormatOutputNETaskDetails = Join(parts, "; ")
End Function

Private Function FormatNumberForReport(ByVal numericValue As Variant) As String
    If IsNumeric(numericValue) Then
        FormatNumberForReport = Format$(CDbl(numericValue), "0.################")
    Else
        FormatNumberForReport = CStr(numericValue)
    End If
End Function

Private Function IsTimeOffTask(ByVal taskName As String) As Boolean
    Dim normalized As String
    normalized = NormalizeForTimeOff(taskName)

    If Len(normalized) = 0 Then Exit Function

    Dim phrases As Variant
    phrases = Array( _
        "sick", "sick time", "sick day", "sick leave", "sicktime", "sickday", "sickleave", _
        "time away", "timeaway", _
        "vacation", "vacation time away", "vacationtimeaway", _
        "pto", "paid time off", "paidtimeoff", _
        "personal time off", "personaltimeoff", _
        "leave", "leave of absence", "leaveofabsence", _
        "holiday", "floating holiday", "floatingholiday", _
        "bereavement")

    Dim collapsed As String
    collapsed = Replace(normalized, " ", "")

    Dim phrase As Variant
    For Each phrase In phrases
        Dim phraseText As String
        phraseText = CStr(phrase)
        If ContainsWholePhrase(normalized, phraseText) Then
            IsTimeOffTask = True
            Exit Function
        End If
        Dim collapsedPhrase As String
        collapsedPhrase = Replace(phraseText, " ", "")
        If Len(collapsedPhrase) <> Len(phraseText) Then
            If InStr(1, collapsed, collapsedPhrase, vbTextCompare) > 0 Then
                IsTimeOffTask = True
                Exit Function
            End If
        End If
    Next phrase
End Function

Private Function NormalizeForTimeOff(ByVal taskName As String) As String
    Dim normalized As String
    normalized = LCase$(Trim$(CStr(taskName)))

    If Len(normalized) = 0 Then Exit Function

    normalized = Replace(normalized, Chr$(160), " ")
    normalized = Replace(normalized, vbTab, " ")

    Dim separators As Variant
    separators = Array("/", "-", "_", "\", ".", ",", "(", ")", ":", ";", "&", "+", "|", "!", "?", "'", Chr$(34), "[", "]", "{", "}", _
                       ChrW$(8211), ChrW$(8212), ChrW$(8216), ChrW$(8217))

    Dim sep As Variant
    For Each sep In separators
        normalized = Replace(normalized, CStr(sep), " ")
    Next sep

    normalized = Replace(normalized, "aaway", "away")

    Do While InStr(normalized, "  ") > 0
        normalized = Replace(normalized, "  ", " ")
    Loop

    NormalizeForTimeOff = Trim$(normalized)
End Function

Private Function ContainsWholePhrase(ByVal textValue As String, ByVal phrase As String) As Boolean
    Dim haystack As String, needle As String
    haystack = " " & textValue & " "
    needle = " " & phrase & " "

    Do While InStr(needle, "  ") > 0
        needle = Replace(needle, "  ", " ")
    Loop

    ContainsWholePhrase = InStr(1, haystack, needle, vbTextCompare) > 0
End Function

Private Function GetKeyFromDateName(ByVal dateValue As Variant, ByVal personName As String) As String
    Dim dt As Double
    If IsDate(dateValue) Then
        dt = CLng(CDate(dateValue))
    Else
        dt = CDbl(dateValue)
    End If
    GetKeyFromDateName = CStr(dt) & "|" & LCase$(Trim$(personName))
End Function

'==========================================================================
' --- Helper: Build unique key for Output rows ---
'==========================================================================
Public Function BuildOutputRowKey(ByVal dateValue As Variant, ByVal nameValue As String, _
                                  ByVal regionValue As String, ByVal taskValue As String) As String
    Dim dtPart As String
    If IsDate(dateValue) Then
        dtPart = Format(CDate(dateValue), "yyyy-mm-dd")
    Else
        dtPart = Trim$(CStr(dateValue))
    End If

    BuildOutputRowKey = LCase$(dtPart & "|" & Trim$(nameValue) & "|" & Trim$(regionValue) & "|" & Trim$(taskValue))
End Function

'==========================================================================
' --- Helper Function: Parse Date From Sheet Name ---
'==========================================================================
Public Function ParseDateFromName(fullName As String, prefix As String) As String
    Dim datePart As String, parts() As String
    If Left(fullName, Len(prefix)) <> prefix Then Exit Function
    datePart = Mid(fullName, Len(prefix) + 1)
    parts = Split(datePart, "-")
    If UBound(parts) <> 2 Then Exit Function
    Dim m As Long, d As Long, yy As Long, dt As Date
    m = Val(parts(0)): d = Val(parts(1)): yy = Val(parts(2))
    If yy < 100 Then yy = yy + 2000
    On Error Resume Next
    dt = DateSerial(yy, m, d)
    On Error GoTo 0
    If dt = 0 Then Exit Function
    ParseDateFromName = Format(dt, "yyyy-mm-dd")
End Function

'==========================================================================
' --- PERFORMANCE OPTIMIZATION: Faster Data Processing ---
'==========================================================================
Public Sub ProcessActivitySheetOptimized(wsInput As Worksheet, theDate As String)
    ' *** OPTIMIZED VERSION: Pre-calculate array sizes and use bulk operations ***
    Dim startTime As Double: startTime = Timer
    
    Const FIRST_TASK_ROW As Long = 2
    Const FIRST_DATA_ROW As Long = 3
    Const FIRST_TASK_COL As Long = 2
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Sheets("Output")
    Dim wsLookup As Worksheet: Set wsLookup = ThisWorkbook.Sheets("ActivityLookup")

    ' *** PERFORMANCE: Use faster dictionary loading ***
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim lkArr As Variant, lastLkRow As Long
    lastLkRow = wsLookup.Cells(wsLookup.Rows.Count, 1).End(xlUp).row
    If lastLkRow > 1 Then
        lkArr = wsLookup.Range("A2:C" & lastLkRow).Value
        Dim r As Long
        For r = 1 To UBound(lkArr, 1): dict(lkArr(r, 1)) = lkArr(r, 2): Next r
    End If
    
    ' *** PERFORMANCE: Calculate exact data range instead of reading entire sheet ***
    Dim lastRow As Long, lastCol As Long
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).row
    lastCol = wsInput.Cells(FIRST_TASK_ROW, wsInput.Columns.Count).End(xlToLeft).Column
    
    If lastRow < FIRST_DATA_ROW Or lastCol < FIRST_TASK_COL Then Exit Sub
    
    ' *** PERFORMANCE: Read only the data we need ***
    Dim inArr As Variant
    inArr = wsInput.Range(wsInput.Cells(1, 1), wsInput.Cells(lastRow, lastCol)).Value
    
    ' *** PERFORMANCE: Pre-calculate maximum possible output size ***
    Dim maxPossibleRows As Long
    maxPossibleRows = (lastRow - FIRST_DATA_ROW + 1) * (lastCol - FIRST_TASK_COL + 1)
    
    Dim outArr() As Variant
    ReDim outArr(1 To maxPossibleRows, 1 To 7)
    
    Dim outPtr As Long: outPtr = 1
    Dim i As Long, j As Long, entryCount As Double, taskName As String, region As String, taskOnly As String
    Dim aht As Variant, prodHrs As Variant
    Const VALID_REGIONS As String = ",BC,AB,CT,ON,QC,MT,YK,"
    
    ' *** PERFORMANCE: Optimized inner loop with early exits ***
    For i = FIRST_DATA_ROW To lastRow
        For j = FIRST_TASK_COL To lastCol
            entryCount = ParseEntryCount(inArr(i, j))
            If entryCount > 0 Then
                taskName = CStr(inArr(FIRST_TASK_ROW, j))
                
                ' *** PERFORMANCE: Faster region detection ***
                Dim spacePos As Long: spacePos = InStr(taskName, " ")
                If spacePos > 0 Then
                    Dim cand As String: cand = Left(taskName, spacePos - 1)
                    If InStr(1, VALID_REGIONS, "," & cand & ",", vbTextCompare) > 0 Then
                        region = cand: taskOnly = Mid(taskName, spacePos + 1)
                    Else
                        region = "AR": taskOnly = taskName
                    End If
                Else
                    region = "AR": taskOnly = taskName
                End If
                
                ' *** PERFORMANCE: Faster lookup ***
                If dict.Exists(taskName) Then aht = dict(taskName) Else aht = "N/A"
                If IsNumeric(aht) Then prodHrs = entryCount * aht / 60 Else prodHrs = "N/A"
                
                outArr(outPtr, 1) = theDate: outArr(outPtr, 2) = inArr(i, 1): outArr(outPtr, 3) = region
                outArr(outPtr, 4) = taskOnly: outArr(outPtr, 5) = entryCount: outArr(outPtr, 6) = aht
                outArr(outPtr, 7) = prodHrs: outPtr = outPtr + 1
            End If
        Next j
    Next i
    
    If outPtr = 1 Then Exit Sub
    
    ' *** PERFORMANCE: Bulk write to output ***
    Dim lastOutRow As Long
    lastOutRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).row
    If lastOutRow = 1 And wsOutput.Cells(1, 1).Value = "" Then
        wsOutput.Range("A1").Resize(1, 7).Value = Array("Date", "Name", "Region", "Task", "Count", "Avg Handle (min)", "Productive Hours")
        lastOutRow = 1
    End If
    
    ' *** PERFORMANCE: Write only the actual data rows ***
    wsOutput.Cells(lastOutRow + 1, 1).Resize(outPtr - 1, 7).Value = outArr
    
    ' *** PERFORMANCE: Apply AutoFilter once ***
    If wsOutput.AutoFilterMode Then wsOutput.AutoFilterMode = False
    wsOutput.Range("A1").AutoFilter
    
    Debug.Print "PERFORMANCE: ProcessActivitySheet completed in " & Format((Timer - startTime), "0.00") & " seconds for " & (outPtr - 1) & " records"
End Sub

