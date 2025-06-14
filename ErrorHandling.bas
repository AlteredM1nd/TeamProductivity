Option Explicit

Private Type ErrorLog
    ErrorNumber As Long
    ErrorDescription As String
    SourceModule As String
    SourceProcedure As String
    TimeStamp As Date
End Type

Public Sub LogError(errorNumber As Long, errorDesc As String, sourceModule As String, sourceProcedure As String)
    Dim wsLog As Worksheet
    Dim errorObj As ErrorLog
    
    With errorObj
        .ErrorNumber = errorNumber
        .ErrorDescription = errorDesc
        .SourceModule = sourceModule
        .SourceProcedure = sourceProcedure
        .TimeStamp = Now
    End With
    
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("ErrorLog")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "ErrorLog"
        wsLog.Range("A1:E1").Value = Array("Timestamp", "Module", "Procedure", "Error #", "Description")
    End If
    wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 5).Value = _
        Array(errorObj.TimeStamp, errorObj.SourceModule, errorObj.SourceProcedure, _
              errorObj.ErrorNumber, errorObj.ErrorDescription)
End Sub
