'==========================================================================
' Module: DataProcessing
'==========================================================================
Option Explicit

'==========================================================================
' --- Process "Personal Entry" Sheet ---
'==========================================================================
Public Sub ProcessActivitySheet(wsInput As Worksheet, theDate As String)

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
    Dim i As Long, j As Long, entryCount As Long, taskName As String, region As String, taskOnly As String
    Dim aht As Variant, prodHrs As Variant, missingDict As Object: Set missingDict = CreateObject("Scripting.Dictionary")
    Const VALID_REGIONS As String = ",BC,AB,CT,ON,QC,MT,YK,"
    
    For i = FIRST_DATA_ROW To lastRow
        For j = FIRST_TASK_COL To lastCol
            entryCount = Val(inArr(i, j))
            If entryCount > 0 Then
                taskName = inArr(FIRST_TASK_ROW, j)
                Dim cand As String: cand = Split(taskName, " ")(0)
                If InStr(1, VALID_REGIONS, "," & cand & ",", vbTextCompare) > 0 Then
                    region = cand: taskOnly = Mid(taskName, Len(region) + 2)
                Else
                    region = "AR": taskOnly = taskName
                End If
                If dict.Exists(taskName) Then aht = dict(taskName) Else aht = "N/A"
                ' Clean any errors that might come from the lookup
                If IsError(aht) Then aht = "N/A"
                
                If IsNumeric(aht) Then prodHrs = entryCount * aht / 60 Else prodHrs = "N/A"
                
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
    Dim i As Long, j As Long, entryCount As Long, taskName As String, region As String, taskOnly As String
    Dim aht As Variant, prodHrs As Variant
    Const VALID_REGIONS As String = ",BC,AB,CT,ON,QC,MT,YK,"
    
    ' *** PERFORMANCE: Optimized inner loop with early exits ***
    For i = FIRST_DATA_ROW To lastRow
        For j = FIRST_TASK_COL To lastCol
            entryCount = Val(inArr(i, j))
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
