Option Explicit

Sub CalculateProductivityMetrics()
    ' Declare all variables at the top
    Dim wsOutput As Worksheet, wsOutputNE As Worksheet, wsDashboard As Worksheet
    Dim dict As Object, monthDict As Object, teamMembers As Object
    Dim arrOutput As Variant, arrOutputNE As Variant
    Dim lastRowOutput As Long, lastRowOutputNE As Long
    Dim i As Long, j As Long
    Dim weekRow As Long, monthRow As Long
    Dim key As Variant, personName As String, workDate As Date
    Dim dailyHours As Double, weeklyHours As Double
    Dim achievedTarget As Long, totalPossible As Long
    Dim startDate As Date, endDate As Date
    Dim weekStartDate As Date, weekEndDate As Date
    Dim currentWeek As Integer, currentMonth As Integer, currentYear As Integer
    Dim weekKey As String, monthKey As String
    Dim lastUpdateCell As Range
    Dim startTime As Double, endTime As Double
    
    ' Start timer for performance measurement
    startTime = Timer
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Set references to worksheets
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("Output")
    Set wsOutputNE = ThisWorkbook.Sheets("OutputNE")
    On Error GoTo 0
    
    If wsOutput Is Nothing Or wsOutputNE Is Nothing Then
        MsgBox "Required sheets not found. Please ensure you have 'Output' and 'OutputNE' sheets.", vbExclamation
        GoTo CleanUp
    End If
    
    ' Load data into arrays for faster processing
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row
    lastRowOutputNE = wsOutputNE.Cells(wsOutputNE.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there's any data to process
    If lastRowOutput <= 1 And lastRowOutputNE <= 1 Then
        MsgBox "No data found in the input sheets.", vbInformation
        GoTo CleanUp
    End If
    
    ' Load data into arrays
    If lastRowOutput > 1 Then arrOutput = wsOutput.Range("A2:G" & lastRowOutput).Value
    If lastRowOutputNE > 1 Then arrOutputNE = wsOutputNE.Range("A2:D" & lastRowOutputNE).Value
    
    ' Check if dashboard exists, if not create it
    On Error Resume Next
    Set wsDashboard = ThisWorkbook.Sheets("ProductivityDashboard")
    On Error GoTo 0

    If wsDashboard Is Nothing Then
        ' Create new dashboard if it doesn't exist
        Set wsDashboard = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDashboard.Name = "ProductivityDashboard"
        
        ' Set up dashboard headers
        With wsDashboard
            .Range("A1").Value = "Productivity Dashboard"
            .Range("A1").Font.Bold = True
            .Range("A1").Font.Size = 14
            
            ' Add last updated timestamp
            .Range("G1").Value = "Last Updated: " & Now()
            .Range("G1").Font.Italic = True
            .Range("G1").HorizontalAlignment = xlRight
            
            ' Weekly Summary
            .Range("A3").Value = "Week Start"
            .Range("B3").Value = "Week End"
            .Range("C3").Value = "Team Members"
            .Range("D3").Value = "Achieved Target"
            .Range("E3").Value = "Total Possible"
            .Range("F3").Value = "Productivity %"
            
            ' Monthly Summary
            .Range("H3").Value = "Month"
            .Range("I3").Value = "Team Members"
            .Range("J3").Value = "Achieved Target"
            .Range("K3").Value = "Total Possible"
            .Range("L3").Value = "Productivity %"
            
            ' Format headers
            .Range("A3:F3").Interior.Color = RGB(200, 220, 255)
            .Range("H3:L3").Interior.Color = RGB(200, 220, 255)
            .Range("A3:L3").Font.Bold = True
            .Columns("A:L").AutoFit
        End With
    Else
        ' Update the last updated timestamp
        Set lastUpdateCell = wsDashboard.Range("G1")
        If lastUpdateCell.Value = "" Then
            Set lastUpdateCell = wsDashboard.Range("G1")
        End If
        lastUpdateCell.Value = "Last Updated: " & Now()
    End If
    
    ' Create dictionaries for data processing
    Set dict = CreateObject("Scripting.Dictionary")
    Set monthDict = CreateObject("Scripting.Dictionary")
    
    ' Process Output sheet from array (productive hours from entries)
    If Not IsEmpty(arrOutput) Then
        For i = LBound(arrOutput, 1) To UBound(arrOutput, 1)
            If Not IsEmpty(arrOutput(i, 1)) And Not IsEmpty(arrOutput(i, 2)) Then
                personName = arrOutput(i, 2)  ' Column B has names
                workDate = arrOutput(i, 1)    ' Column A has dates
                dailyHours = IIf(IsNumeric(arrOutput(i, 7)), arrOutput(i, 7), 0) ' Column G has hours
                
                ' Create unique key for person and date
                key = personName & "|" & Format(workDate, "yyyy-mm-dd")
                
                ' Add to dictionary
                If dict.Exists(key) Then
                    dict(key) = dict(key) + dailyHours
                Else
                    dict.Add key, dailyHours
                End If
            End If
        Next i
    End If
    
    ' Process OutputNE sheet from array (non-entry hours)
    If Not IsEmpty(arrOutputNE) Then
        For i = LBound(arrOutputNE, 1) To UBound(arrOutputNE, 1)
            If Not IsEmpty(arrOutputNE(i, 1)) And Not IsEmpty(arrOutputNE(i, 2)) Then
                personName = arrOutputNE(i, 2)  ' Column B has names
                workDate = arrOutputNE(i, 1)    ' Column A has dates
                dailyHours = IIf(IsNumeric(arrOutputNE(i, 4)), arrOutputNE(i, 4), 0) ' Column D has hours
                
                ' Create unique key for person and date
                key = personName & "|" & Format(workDate, "yyyy-mm-dd")
                
                ' Add to dictionary
                If dict.Exists(key) Then
                    dict(key) = dict(key) + dailyHours
                Else
                    dict.Add key, dailyHours
                End If
            End If
        Next i
    End If
    
    ' Find date range
    startDate = Date
    endDate = DateSerial(1900, 1, 1)
    For Each key In dict.Keys
        workDate = CDate(Split(key, "|")(1))
        If workDate < startDate Then startDate = workDate
        If workDate > endDate Then endDate = workDate
    Next key
    
    ' Initialize row counters
    weekRow = 4
    monthRow = 4
    
    ' Set up date variables
    currentWeek = DatePart("ww", startDate)
    currentMonth = Month(startDate)
    currentYear = Year(startDate)
    weekStartDate = startDate - Weekday(startDate, vbSunday) + 1
    weekEndDate = weekStartDate + 6
    
    ' Initialize counters
    achievedTarget = 0
    totalPossible = 0
    
    ' Process each week
    Do While weekStartDate <= endDate
        Set teamMembers = CreateObject("Scripting.Dictionary")
        achievedTarget = 0
        totalPossible = 0
        
        ' Check each day in the week
        For i = 0 To 6
            workDate = weekStartDate + i
            If workDate > endDate Then Exit For
            
            ' Check each person for this date
            For Each key In dict.Keys
                If CDate(Split(key, "|")(1)) = workDate Then
                    personName = Split(key, "|")(0)
                    dailyHours = dict(key)
                    
                    ' Add to weekly total for this person
                    If teamMembers.Exists(personName) Then
                        teamMembers(personName) = teamMembers(personName) + dailyHours
                    Else
                        teamMembers.Add personName, dailyHours
                    End If
                End If
            Next key
        Next i
        
        ' Check who achieved weekly target
        For Each key In teamMembers.Keys
            weeklyHours = teamMembers(key)
            totalPossible = totalPossible + 1
            If weeklyHours >= 32.5 Then achievedTarget = achievedTarget + 1
        Next key
        
        ' Add to weekly summary
        If totalPossible > 0 Then
            wsDashboard.Cells(weekRow, 1).Value = weekStartDate
            wsDashboard.Cells(weekRow, 2).Value = weekEndDate
            wsDashboard.Cells(weekRow, 3).Value = teamMembers.Count
            wsDashboard.Cells(weekRow, 4).Value = achievedTarget
            wsDashboard.Cells(weekRow, 5).Value = totalPossible
            wsDashboard.Cells(weekRow, 6).Value = Format(achievedTarget / totalPossible, "0.00%")
            weekRow = weekRow + 1
        End If
        
        ' Add to monthly summary
        monthKey = Format(weekStartDate, "yyyy-mm")
        If monthDict.Exists(monthKey) Then
            monthDict(monthKey) = Array(monthDict(monthKey)(0) + achievedTarget, _
                                       monthDict(monthKey)(1) + totalPossible)
        Else
            monthDict.Add monthKey, Array(achievedTarget, totalPossible)
        End If
        
        ' Move to next week
        weekStartDate = weekStartDate + 7
        weekEndDate = weekStartDate + 6
    Loop
    
    ' Process monthly metrics
    For Each key In monthDict.Keys
        achievedTarget = monthDict(key)(0)
        totalPossible = monthDict(key)(1)
        
        wsDashboard.Cells(monthRow, 8).Value = Format(CDate(key & "-01"), "yyyy-mmm")
        wsDashboard.Cells(monthRow, 9).Value = "N/A" ' Team members count not tracked per month
        wsDashboard.Cells(monthRow, 10).Value = achievedTarget
        wsDashboard.Cells(monthRow, 11).Value = totalPossible
        wsDashboard.Cells(monthRow, 12).Value = Format(achievedTarget / totalPossible, "0.00%")
        monthRow = monthRow + 1
    Next key
    
    ' Format the dashboard
    With wsDashboard
        ' Add borders to data
        .Range("A3:L" & Application.Max(weekRow, monthRow)).Borders.LineStyle = xlContinuous
        
        ' Format percentages
        If weekRow > 4 Then .Range("F4:F" & weekRow - 1).NumberFormat = "0.00%"
        If monthRow > 4 Then .Range("L4:L" & monthRow - 1).NumberFormat = "0.00%"
        
        ' Auto-fit columns and select top
        .Range("A3:L3").EntireColumn.AutoFit
        .Range("A1").Select
        
        ' Freeze panes for better navigation
        .Activate
        .Range("A4").Select
        ActiveWindow.FreezePanes = True
    End With
    
    ' Calculate and show execution time
    endTime = Timer
    Dim execTime As String
    execTime = Format((endTime - startTime), "0.00") & " seconds"
    
    MsgBox "Productivity metrics have been calculated successfully!" & vbNewLine & _
           "Execution time: " & execTime, vbInformation, "Process Complete"
           
CleanUp:
    ' Restore application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub