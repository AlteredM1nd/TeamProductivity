'--------------------------------------------------------------------
' 2. BULK RUNNER  –  last 12 months of “Non-Entry Hrs M-D-YY” tabs
'--------------------------------------------------------------------
Public Sub BulkProcessNonEntryLastYear()

    Const MONTHS_BACK As Long = 18
    Const prefix      As String = "Non-Entry Hrs "
    Const DELIM       As String = "-"
    
    Dim targetDate As Date: targetDate = DateAdd("m", -MONTHS_BACK, Date)
    Dim ws As Worksheet, processed As Long, skipped As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.name, Len(prefix)) = prefix Then
            Dim datePart As String: datePart = Mid(ws.name, Len(prefix) + 1) 'e.g. "5-1-25"
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
                    skipped = skipped & ws.name & "  (too old)" & vbCrLf
                End If
            Else
                skipped = skipped & ws.name & "  (bad name)" & vbCrLf
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
