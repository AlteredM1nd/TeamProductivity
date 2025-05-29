'Converts "Prefix M-D-YY" to "yyyy-mm-dd"
Public Function ParseDateFromName(fullName As String, PREFIX As String) As String
    
    Dim datePart As String, parts() As String
    If Left(fullName, Len(PREFIX)) <> PREFIX Then
        MsgBox "Sheet name """ & fullName & _
               """ doesn’t start with """ & PREFIX & """", vbCritical
        Exit Function
    End If
    
    datePart = Mid(fullName, Len(PREFIX) + 1)        'e.g. 5-1-25
    parts = Split(datePart, "-")
    If UBound(parts) <> 2 Then
        MsgBox "Can’t parse date from sheet name """ & fullName & """.", vbCritical
        Exit Function
    End If
    
    Dim m As Long, d As Long, yy As Long, dt As Date
    m = Val(parts(0)): d = Val(parts(1)): yy = Val(parts(2))
    If yy < 100 Then yy = yy + 2000                  '25 ? 2025
    
    On Error Resume Next
    dt = DateSerial(yy, m, d)
    On Error GoTo 0
    If dt = 0 Then
        MsgBox "Invalid date in sheet name """ & fullName & """.", vbCritical
        Exit Function
    End If
    
    ParseDateFromName = Format(dt, "yyyy-mm-dd")
End Function
