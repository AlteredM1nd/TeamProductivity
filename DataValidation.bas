Option Explicit

Public Function ValidateDataSheet(ws As Worksheet, Optional requiredColumns As Variant) As Boolean
    If ws Is Nothing Then
        ValidateDataSheet = False
        Exit Function
    End If
    
    If IsEmpty(requiredColumns) Then
        requiredColumns = Array("Date", "Name", "Task", "Count")
    End If
    
    Dim headerRange As Range
    Set headerRange = ws.Range("A1").Resize(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)
    
    Dim col As Variant
    For Each col In requiredColumns
        If Application.Match(col, headerRange, 0) Is Nothing Then
            ValidateDataSheet = False
            Exit Function
        End If
    Next col
    
    ValidateDataSheet = True
End Function

Public Function ValidateDate(dateToCheck As Date, Optional minDate As Date) As Boolean
    If minDate = 0 Then
        minDate = DateSerial(2024, 1, 1)
    End If
    
    ValidateDate = dateToCheck >= minDate And dateToCheck <= Date
End Function

Public Function ValidateWorkbook(wb As Workbook, requiredSheets As Variant) As Boolean
    If wb Is Nothing Then
        ValidateWorkbook = False
        Exit Function
    End If
    
    Dim sheet As Variant
    For Each sheet In requiredSheets
        If GetWorksheet(wb, CStr(sheet)) Is Nothing Then
            ValidateWorkbook = False
            Exit Function
        End If
    Next sheet
    
    ValidateWorkbook = True
End Function

Private Function GetWorksheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Sheets(sheetName)
    On Error GoTo 0
End Function
