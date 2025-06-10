Option Explicit

' Version 4.3: Added user selection for the destination workbook.
' The macro can be stored in any .xlsm file and can process any source/destination files.
Sub ProcessAwayTime_SelectDestinationFile()

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
    
    ' --- *** MODIFIED START *** ---
    ' 'Set destWB = ThisWorkbook' <-- This line is removed. We will now prompt the user.
    ' --- *** MODIFIED END *** ---

    ' --- 2. OPTIMIZATION ---
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' --- *** NEW SECTION START: GET DESTINATION FILE *** ---
    ' This new section prompts the user to select the target/destination workbook.
    Dim destPath As String
    Dim wb As Workbook
    
    MsgBox "Please select the DESTINATION workbook where the hours will be recorded.", vbInformation, "Select Destination File"
    destPath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", 1, "Select the Destination Workbook")
    
    ' Exit if the user cancels the selection
    If destPath = "False" Then
        MsgBox "No destination file selected. Macro is aborting.", vbExclamation
        GoTo CleanUp
    End If
    
    ' Check if the selected workbook is already open to avoid errors.
    ' If it is, set our object to the open workbook. If not, open it.
    Dim destFilename As String
    destFilename = Mid(destPath, InStrRev(destPath, "\") + 1)
    
    Set destWB = Nothing
    For Each wb In Application.Workbooks
        If wb.name = destFilename Then
            Set destWB = wb
            Exit For
        End If
    Next wb
    
    ' If the workbook wasn't found in the open workbooks collection, open it.
    If destWB Is Nothing Then
        Set destWB = Workbooks.Open(destPath)
    End If
    ' --- *** NEW SECTION END *** ---

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
                        ' 1. Store the original value from the target column for more detailed logging.
                        oldValue = destWS.Cells(matchRow, targetCol).Value
                        If IsEmpty(oldValue) Then oldValStr = "Empty" Else oldValStr = CStr(oldValue)
                        
                        ' 2. Clear both the Sick (Col 16) and Away (Col 17) fields first.
                        destWS.Cells(matchRow, 16).ClearContents ' Clear Sick Hours field
                        destWS.Cells(matchRow, 17).ClearContents ' Clear Away Hours field

                        ' 3. Write the new value from the source sheet to the correct column.
                        destWS.Cells(matchRow, targetCol).Value = hours
                        
                        ' 4. Update the log message to be more descriptive.
                        WriteToLog logWS, logRow, "Success", personName, entryDate, hours, payCategory, destWS.name, "Cleared Sick/Away, then wrote value to " & destWS.Cells(matchRow, targetCol).Address & ". Old value in that cell was: " & oldValStr
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
    
    ' --- *** MODIFIED START *** ---
    ' We now save the destination workbook, which may be different from ThisWorkbook
    If Not destWB Is Nothing Then destWB.Save
    ' --- *** MODIFIED END *** ---

    MsgBox "Processing complete!" & vbCrLf & vbCrLf & "A detailed report has been generated on the 'Macro Log' sheet in '" & destFilename & "'.", vbInformation, "Macro Finished"

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
