Option Explicit

Private Type ConfigSettings
    StartDate As Date
    DailyTarget As Double
    SickDayHours As Double
    WorkdayStartHour As Integer
    WorkdayEndHour As Integer
End Type

Public Function LoadConfig() As ConfigSettings
    Dim wsConfig As Worksheet
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    If wsConfig Is Nothing Then
        Call InitializeConfig
        Set wsConfig = ThisWorkbook.Sheets("Config")
    End If
    On Error GoTo 0
    
    With LoadConfig
        .StartDate = wsConfig.Range("StartDate").Value
        .DailyTarget = wsConfig.Range("DailyTarget").Value
        .SickDayHours = wsConfig.Range("SickDayHours").Value
        .WorkdayStartHour = wsConfig.Range("WorkdayStartHour").Value
        .WorkdayEndHour = wsConfig.Range("WorkdayEndHour").Value
    End With
End Function

Private Sub InitializeConfig()
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets.Add
    wsConfig.Name = "Config"
    
    With wsConfig
        .Range("A1").Value = "Configuration Settings"
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Start Date"
        .Range("B3").Value = DateSerial(2024, 1, 1)
        .Range("B3").Name = "StartDate"
        
        .Range("A4").Value = "Daily Target Hours"
        .Range("B4").Value = 6.5
        .Range("B4").Name = "DailyTarget"
        
        .Range("A5").Value = "Sick Day Hours"
        .Range("B5").Value = 7.5
        .Range("B5").Name = "SickDayHours"
        
        .Range("A6").Value = "Workday Start Hour"
        .Range("B6").Value = 9
        .Range("B6").Name = "WorkdayStartHour"
        
        .Range("A7").Value = "Workday End Hour"
        .Range("B7").Value = 17
        .Range("B7").Name = "WorkdayEndHour"
        
        .Columns("A:B").AutoFit
    End With
End Sub

Public Function GetConfig(configName As String) As Variant
    Dim config As ConfigSettings
    config = LoadConfig
    
    Select Case configName
        Case "StartDate": GetConfig = config.StartDate
        Case "DailyTarget": GetConfig = config.DailyTarget
        Case "SickDayHours": GetConfig = config.SickDayHours
        Case "WorkdayStartHour": GetConfig = config.WorkdayStartHour
        Case "WorkdayEndHour": GetConfig = config.WorkdayEndHour
        Case Else: GetConfig = Null
    End Select
End Function
