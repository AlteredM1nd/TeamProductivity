'==========  Module3  ==========
Option Explicit

'--------------------------------------------------------------------
' 1. PROCESS ONE “Non-Entry Hrs” SHEET
'--------------------------------------------------------------------
Public Sub ProcessNonEntrySheet(wsInput As Worksheet, theDate As String)

    'layout constants for Non-Entry sheets
    Const FIRST_HEADER_COL As Long = 4      'D
    Const LAST_HEADER_COL  As Long = 19     'S
    Const FIRST_NAME_ROW   As Long = 2      'names start in A2
    Const NAME_COL         As Long = 1      'A
    Const HEADER_ROW       As Long = 1      'headers in row 1
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Sheets("OutputNE")
    
    '1) find the last populated name row (2-43 in your template)
    Dim lastRow As Long
    lastRow = wsInput.Cells(wsInput.Rows.Count, NAME_COL).End(xlUp).Row
    If lastRow < FIRST_NAME_ROW Then Exit Sub   'nothing here
    
    '2) read the whole block into memory (fast)
    Dim inArr As Variant
    inArr = wsInput.Range(wsInput.Cells(HEADER_ROW, NAME_COL), _
                          wsInput.Cells(lastRow, LAST_HEADER_COL)).Value
    
    '3) build the output array (Date | Name | Task | Count)
    Dim outArr() As Variant
    ReDim outArr(1 To (lastRow - FIRST_NAME_ROW + 1) * _
                     (LAST_HEADER_COL - FIRST_HEADER_COL + 1), 1 To 4)
                     
    Dim outPtr As Long: outPtr = 1
    Dim i As Long, j As Long, taskName As String, countVal As Variant
    
    For i = FIRST_NAME_ROW To lastRow
        For j = FIRST_HEADER_COL To LAST_HEADER_COL
            countVal = inArr(i, j)
            If IsNumeric(countVal) And countVal > 0 Then
                taskName = inArr(HEADER_ROW, j)
    
                '--- NEW: remove any hard returns in the header cell
                taskName = Replace(taskName, vbLf, " ")     ' <- add this line
                '--- optional: collapse multiple spaces that remain
                taskName = Application.Trim(Replace(taskName, "  ", " "))
    
                outArr(outPtr, 1) = theDate
                outArr(outPtr, 2) = inArr(i, NAME_COL)
                outArr(outPtr, 3) = taskName
                outArr(outPtr, 4) = countVal
                outPtr = outPtr + 1
            End If
        Next j
    Next i
    
    If outPtr = 1 Then Exit Sub   'nothing non-zero
    
    '4) write to OutputNE
    Dim lastOutRow As Long
    lastOutRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row
    
    If lastOutRow = 1 And wsOutput.Cells(1, 1).Value = "" Then
        wsOutput.Range("A1").Resize(1, 4).Value = _
            Array("Date", "Name", "Task", "Count")
        lastOutRow = 1
    End If
    
    wsOutput.Cells(lastOutRow + 1, 1).Resize(outPtr - 1, 4).Value = outArr
End Sub



'--------------------------------------------------------------------
' 2. BULK RUNNER  –  last 12 months of “Non-Entry Hrs M-D-YY” tabs
'--------------------------------------------------------------------
Public Sub BulkProcessNonEntryLastYear()

    Const MONTHS_BACK As Long = 17
    Const PREFIX      As String = "Non-Entry Hrs "
    Const DELIM       As String = "-"
    
    Dim targetDate As Date: targetDate = DateAdd("m", -MONTHS_BACK, Date)
    Dim ws As Worksheet, processed As Long, skipped As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, Len(PREFIX)) = PREFIX Then
            Dim datePart As String: datePart = Mid(ws.Name, Len(PREFIX) + 1) 'e.g. "5-1-25"
            Dim parts() As String: parts = Split(datePart, DELIM)
            
            If UBound(parts) = 2 Then
                Dim m As Long, d As Long, yy As Long, sheetDate As Date
                m = Val(parts(0)): d = Val(parts(1)): yy = Val(parts(2))
                If yy < 100 Then yy = yy + 2000
                On Error Resume Next: sheetDate = DateSerial(yy, m, d): On Error GoTo 0
                
                If sheetDate >= targetDate Then
                    ProcessNonEntrySheet ws, Format(sheetDate, "yyyy-mm-dd")
                    processed = processed + 1
                Else
                    skipped = skipped & ws.Name & "  (too old)" & vbCrLf
                End If
            Else
                skipped = skipped & ws.Name & "  (bad name)" & vbCrLf
            End If
        End If
    Next ws
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox processed & " sheet(s) processed." & _
           IIf(skipped <> "", vbCrLf & "Skipped:" & vbCrLf & skipped, ""), _
           vbInformation, "Non-Entry import complete"
End Sub
'====================================================================
