Option Explicit

' Version 4.2: Added logic to clear both Sick (Col 16) and Away (Col 17) fields
'              before inputting new data to prevent double-counting.
' Stored and run FROM the Destination Workbook (.xlsm)
Sub ProcessAwayTime_WithDetailedLogging()

    ' --- 1. SETUP & DECLARATIONS ---
    Dim sourceWB As Workbook, destWB As Workbook
    Dim sourceWS As Worksheet, destWS As Worksheet, logWS As Worksheet
    Dim sourceRange As Range
    Dim cell As Range
    Dim sourceLastRow As Long, logRow As Long
    Dim matchRow As Variant

    ' Data variables
    Dim personName As String, payCategory As String
    Dim entryDate As Date
    Dim hours As Double
    Dim oldValue As Variant
    Dim oldValStr As String ' String to hold the old value for logging
    
    ' Destination variables
    Dim sheetNameYY As String, sheetNameYYYY As String
    
    Set destWB = ThisWorkbook

    ' --- 2. OPTIMIZATION ---
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' --- 3. SETUP THE LOG SHEET ---
    On Error Resume Next
    Set logWS = destWB.Worksheets("Macro Log")
    On Error GoTo 0

    If logWS Is Nothing Then
        Set logWS = destWB.Worksheets.Add(After:=destWB.Sheets(destWB.Sheets.Count))
        logWS.name = "Macro Log"
    End If

    logWS.Cells.Clear
    logWS.Range("A1:G1").Value = Array("Status", "Name", "Date", "Hours", "Category", "Target Sheet", "Details")
    logRow = 2

    ' --- 4. GET SOURCE FILE ---
    MsgBox "Please select the SOURCE workbook that contains the master list of away time.", vbInformation, "Select Source File"
    Dim sourcePath As String
    sourcePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", 1, "Select the Source Workbook")
    If sourcePath = "False" Then GoTo CleanUp
    Set sourceWB = Workbooks.Open(sourcePath)

    On Error Resume Next
    Set sourceWS = sourceWB.Worksheets(InputBox("Enter the name of the sheet containing the away time data:", "Source Sheet Name"))
    On Error GoTo 0
    If sourceWS Is Nothing Then
        MsgBox "The sheet name you entered was not found. Aborting.", vbCritical, "Error"
        WriteToLog logWS, logRow, "Fatal Error", "", Now(), 0, "", "", "Source sheet not found. Macro stopped."
        GoTo CleanUp
    End If

    ' --- 5. PROCESSING LOOP ---
    sourceLastRow = sourceWS.Cells(sourceWS.Rows.Count, "A").End(xlUp).row
    Set sourceRange = sourceWS.Range("A2:H" & sourceLastRow)

    For Each cell In sourceRange.Rows
        Set destWS = Nothing
        Dim targetCol As Integer: targetCol = 0
        
        personName = Trim(cell.Columns("A").Value)
        payCategory = Trim(cell.Columns("G").Value)
        
        If IsDate(cell.Columns("F").Value) And IsNumeric(cell.Columns("H").Value) And personName <> "" Then
            entryDate = cell.Columns("F").Value
            hours = cell.Columns("H").Value
            
            sheetNameYY = "Non-Entry Hrs " & Format(entryDate, "M-D-YY")
            sheetNameYYYY = "Non-Entry Hrs " & Format(entryDate, "M-D-YYYY")

            On Error Resume Next
            Set destWS = destWB.Worksheets(sheetNameYY)
            If destWS Is Nothing Then Set destWS = destWB.Worksheets(sheetNameYYYY)
            On Error GoTo 0

            If Not destWS Is Nothing Then
                Select Case UCase(payCategory)
                    Case "SICK": targetCol = 16 ' Column P
                    Case "PERSONAL", "VACATION", "BEREAVEMENT", "FLOAT", "MY COMMUNITY", "STUDY": targetCol = 17 ' Column Q
                    Case Else
                        WriteToLog logWS, logRow, "Failed - Category", personName, entryDate, hours, payCategory, destWS.name, "Pay category is not recognized."
                End Select

                If targetCol > 0 Then
                    matchRow = Application.Match(personName, destWS.Columns("A"), 0)
                    
                    If Not IsError(matchRow) Then
                        ' --- *** MODIFIED LOGIC STARTS HERE (v4.2) *** ---
                        
                        ' 1. Store the original value from the target column for more detailed logging.
                        oldValue = destWS.Cells(matchRow, targetCol).Value
                        
                        ' Convert the old value to a string for precise logging ("Empty" vs "0").
                        If IsEmpty(oldValue) Then
                            oldValStr = "Empty"
                        Else
                            oldValStr = CStr(oldValue)
                        End If
                        
                        ' 2. NEW: Clear both the Sick (Col 16) and Away (Col 17) fields first.
                        ' This prevents double-counting if an entry was changed (e.g., from Away to Sick).
                        destWS.Cells(matchRow, 16).ClearContents ' Clear Sick Hours field
                        destWS.Cells(matchRow, 17).ClearContents ' Clear Away Hours field

                        ' 3. Write the new value from the source sheet to the correct column.
                        destWS.Cells(matchRow, targetCol).Value = hours
                        
                        ' 4. Update the log message to be more descriptive of the new process.
                        WriteToLog logWS, logRow, "Success", personName, entryDate, hours, payCategory, destWS.name, "Cleared Sick/Away, then wrote value to " & destWS.Cells(matchRow, targetCol).Address & ". Old value in that cell was: " & oldValStr
                        
                        ' --- *** MODIFIED LOGIC ENDS HERE *** ---
                    Else
                        WriteToLog logWS, logRow, "Failed - Name", personName, entryDate, hours, payCategory, destWS.name, "Name not found in Column A."
                    End If
                End If
            Else
                WriteToLog logWS, logRow, "Failed - Sheet", personName, entryDate, hours, payCategory, sheetNameYY & " or " & sheetNameYYYY, "The required dated sheet does not exist."
            End If
        Else
            WriteToLog logWS, logRow, "Failed - Data", personName, CDate(0), 0, payCategory, "N/A", "Row skipped due to invalid/missing date, hours, or name."
        End If
    Next cell

    logWS.Columns("A:G").AutoFit

    ' --- 6. CLEANUP & FINAL REPORT ---
CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    If Not sourceWB Is Nothing Then sourceWB.Close SaveChanges:=False
    
    If Not destWB Is Nothing Then destWB.Save

    MsgBox "Processing complete!" & vbCrLf & vbCrLf & "A detailed report has been generated on the 'Macro Log' sheet.", vbInformation, "Macro Finished"

End Sub


Private Sub WriteToLog(ByVal ws As Worksheet, ByRef row As Long, _
                       ByVal status As String, ByVal name As String, _
                       ByVal entryDate As Date, ByVal hours As Double, _
                       ByVal category As String, ByVal targetSheet As String, _
                       ByVal details As String)
    With ws
        .Cells(row, "A").Value = status
        .Cells(row, "B").Value = name
        If CStr(entryDate) <> CStr(CDate(0)) Then .Cells(row, "C").Value = entryDate
        If hours > 0 Then .Cells(row, "D").Value = hours
        .Cells(row, "E").Value = category
        .Cells(row, "F").Value = targetSheet
        .Cells(row, "G").Value = details
    End With
    row = row + 1
End Sub
