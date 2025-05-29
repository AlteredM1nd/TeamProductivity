'==========  Module2 - Personal Entry Processor =================================================
Option Explicit

'––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––
'  PROCESS ONE “Personal Entry” SHEET
'  Now outputs: Date | Name | Region | Task | Count | Avg Handle (min) | Productive Hours
'  Productive Hours = Count * AHT / 60
'––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––
Public Sub ProcessActivitySheet(wsInput As Worksheet, theDate As String)

    Const FIRST_TASK_ROW As Long = 2
    Const FIRST_DATA_ROW As Long = 3
    Const FIRST_TASK_COL As Long = 2
    
    Dim wsOutput As Worksheet: Set wsOutput = ThisWorkbook.Sheets("Output")
    Dim wsLookup As Worksheet: Set wsLookup = ThisWorkbook.Sheets("ActivityLookup")

    '== 1. load lookup table (Activity | AHT | Target) ================
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim lkArr As Variant
    Dim lastLkRow As Long: lastLkRow = wsLookup.Cells(wsLookup.Rows.Count, 1).End(xlUp).Row
    lkArr = wsLookup.Range("A2:C" & lastLkRow).Value
    
    Dim r As Long
    For r = 1 To UBound(lkArr, 1)
        dict(lkArr(r, 1)) = lkArr(r, 2)         'store ONLY Avg-Handle-Time (minutes)
    Next r
    
    '== 2. read input block ===========================================
    Dim lastRow As Long, lastCol As Long
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row
    lastCol = wsInput.Cells(FIRST_TASK_ROW, wsInput.Columns.Count).End(xlToLeft).Column
    
    Dim inArr As Variant
    inArr = wsInput.Range(wsInput.Cells(1, 1), wsInput.Cells(lastRow, lastCol)).Value
    
    '== 3. build condensed rows =======================================
    Dim outArr() As Variant
    ReDim outArr(1 To (lastRow - FIRST_DATA_ROW + 1) * (lastCol - FIRST_TASK_COL + 1), 1 To 7)
    
    Dim outPtr As Long: outPtr = 1
    Dim i As Long, j As Long, entryCount As Long
    Dim taskName As String, region As String, taskOnly As String
    Dim aht As Variant, prodHrs As Variant
    Dim missingDict As Object: Set missingDict = CreateObject("Scripting.Dictionary")
    
    Const VALID_REGIONS As String = ",BC,AB,CT,ON,QC,MT,YK,"
    
    For i = FIRST_DATA_ROW To lastRow
        For j = FIRST_TASK_COL To lastCol
            entryCount = Val(inArr(i, j))
            If entryCount > 0 Then
                taskName = inArr(FIRST_TASK_ROW, j)
                
                '–– region / task parsing with fallback ––
                Dim cand As String: cand = Split(taskName, " ")(0)
                If InStr(1, VALID_REGIONS, "," & cand & ",", vbTextCompare) > 0 Then
                    region = cand
                    taskOnly = Mid(taskName, Len(region) + 2)
                Else
                    region = "AR"
                    taskOnly = taskName
                End If
                
                '–– lookup Avg Handle Time; mark missing if needed ––
                If dict.Exists(taskName) Then
                    aht = dict(taskName)          'minutes
                Else
                    aht = "N/A"
                    If Not missingDict.Exists(taskName) Then missingDict(taskName) = 1
                End If
                
                '–– productive hours (only if AHT numeric) ––
                If IsNumeric(aht) Then
                    prodHrs = entryCount * aht / 60
                Else
                    prodHrs = "N/A"
                End If
                
                '–– write to array ––
                outArr(outPtr, 1) = theDate
                outArr(outPtr, 2) = inArr(i, 1)
                outArr(outPtr, 3) = region
                outArr(outPtr, 4) = taskOnly
                outArr(outPtr, 5) = entryCount
                outArr(outPtr, 6) = aht
                outArr(outPtr, 7) = prodHrs
                outPtr = outPtr + 1
            End If
        Next j
    Next i
    
    If outPtr = 1 Then Exit Sub           'nothing non-zero

    '== 4. append to Output ===========================================
    Dim lastOutRow As Long
    lastOutRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row
    
    If lastOutRow = 1 And wsOutput.Cells(1, 1).Value = "" Then
        wsOutput.Range("A1").Resize(1, 7).Value = _
            Array("Date", "Name", "Region", "Task", "Count", _
                  "Avg Handle (min)", "Productive Hours")
        lastOutRow = 1
    End If
    
    wsOutput.Cells(lastOutRow + 1, 1).Resize(outPtr - 1, 7).Value = outArr
    
End Sub

'----------------------------------------------------------------------
' 2.  BULK RUNNER  –  processes every sheet named
'     "Personal Entry M-D-YY" (12 months back)
'----------------------------------------------------------------------

Public Sub BulkProcessLastYear()

    Const MONTHS_BACK As Long = 17           'I modified this to run for 17 months to capture all data from 2024
    Const PREFIX      As String = "Personal Entry "
    Const DELIM       As String = "-"
    
    Dim targetDate As Date: targetDate = DateAdd("m", -MONTHS_BACK, Date)
    Dim ws As Worksheet, processed As Long, skipped As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, Len(PREFIX)) = PREFIX Then
            Dim datePart As String: datePart = Mid(ws.Name, Len(PREFIX) + 1)
            Dim parts() As String: parts = Split(datePart, DELIM)
            
            If UBound(parts) = 2 Then
                Dim m As Long, d As Long, yy As Long, sheetDate As Date
                m = Val(parts(0)): d = Val(parts(1)): yy = Val(parts(2))
                If yy < 100 Then yy = yy + 2000
                On Error Resume Next: sheetDate = DateSerial(yy, m, d): On Error GoTo 0
                
                If sheetDate >= targetDate Then
                    ProcessActivitySheet ws, Format(sheetDate, "yyyy-mm-dd")
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
           vbInformation, "Personal-Entry import complete"
End Sub
'=====================================================================
