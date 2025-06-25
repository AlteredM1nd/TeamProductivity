' AT THE VERY TOP OF THE MODULE, BEFORE ANY SUBS OR FUNCTIONS:
Option Explicit  '<-- ENSURE THIS LINE IS PRESENT

Sub CreateMonthlySheetsFromTemplates()

    Dim targetMonth As Integer
    Dim targetYear As Integer
    Dim inputMonth As String
    Dim inputYear As String
    Dim firstDayOfMonth As Date
    Dim lastDayOfMonth As Date
    Dim currentDate As Date
    Dim personalSheetName As String
    Dim nonEntrySheetName As String
    Dim ws As Worksheet ' Will represent the newly copied sheet
    Dim sheetCreatedCount As Long ' Variable to count created sheets

    ' --- DEFINE TEMPLATE SHEET NAMES ---
    Const TEMPLATE_PERSONAL_NAME As String = "Personal Entry" '<<< REVERTED TO ORIGINAL
    Const TEMPLATE_NON_ENTRY_NAME As String = "Non-Entry Hrs"  '<<< REVERTED TO ORIGINAL

    ' --- DEFINE CELLS FOR DATE INSERTION ---
    Const PERSONAL_ENTRY_DATE_CELL As String = "A2"
    Const NON_ENTRY_DATE_CELL As String = "A1"

    Dim wsPersonalTemplate As Worksheet
    Dim wsNonEntryTemplate As Worksheet
    Dim currentSheetType As String ' Track which type of sheet we're currently processing

    ' --- Get User Input ---
    On Error GoTo InvalidInput
    inputMonth = InputBox("Enter the month number (e.g., 6 for June):", "Input Month")
    If inputMonth = "" Then Exit Sub ' User pressed Cancel
    If Not IsNumeric(inputMonth) Then
        MsgBox "Please enter a numeric value for the month.", vbExclamation
        Exit Sub
    End If
    targetMonth = CInt(inputMonth)

    inputYear = InputBox("Enter the year (e.g., 2025):", "Input Year")
    If inputYear = "" Then Exit Sub ' User pressed Cancel
    If Not IsNumeric(inputYear) Then
        MsgBox "Please enter a numeric value for the year.", vbExclamation
        Exit Sub
    End If
    targetYear = CInt(inputYear)
    On Error GoTo 0 ' Reset error handling

    ' --- Validate Input ---
    If targetMonth < 1 Or targetMonth > 12 Then
        MsgBox "Invalid month number. Please enter a number between 1 and 12.", vbExclamation
        Exit Sub
    End If
    If targetYear < 1900 Or targetYear > 2999 Then
        MsgBox "Invalid year. Please enter a sensible year (e.g., 2000-2050).", vbExclamation
        Exit Sub
    End If

    ' --- Check if workbook is protected ---
    If ThisWorkbook.ProtectStructure Then
        MsgBox "Workbook structure is protected. Please unprotect the workbook before running this macro.", vbCritical
        Exit Sub
    End If

    ' --- Get references to template sheets and check if they exist ---
    On Error Resume Next ' Temporarily ignore errors for checking
    Set wsPersonalTemplate = Nothing
    Set wsNonEntryTemplate = Nothing

    Set wsPersonalTemplate = ThisWorkbook.Sheets(TEMPLATE_PERSONAL_NAME)
    If wsPersonalTemplate Is Nothing Then
        MsgBox "Template sheet '" & TEMPLATE_PERSONAL_NAME & "' not found!" & vbCrLf & _
               "Please ensure a sheet with this exact name exists in the workbook.", vbCritical
        Exit Sub
    End If

    Set wsNonEntryTemplate = ThisWorkbook.Sheets(TEMPLATE_NON_ENTRY_NAME)
    If wsNonEntryTemplate Is Nothing Then
        MsgBox "Template sheet '" & TEMPLATE_NON_ENTRY_NAME & "' not found!" & vbCrLf & _
               "Please ensure a sheet with this exact name exists in the workbook.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0 ' Restore default error handling

    ' --- Determine First and Last Day of the Month ---
    firstDayOfMonth = DateSerial(targetYear, targetMonth, 1)
    lastDayOfMonth = DateSerial(targetYear, targetMonth + 1, 0)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Disable calculation to speed up and avoid calc warnings
    sheetCreatedCount = 0 ' Initialize count ONCE before the loop

    Debug.Print "Starting sheet creation process. Initial count: " & sheetCreatedCount ' DEBUG

    ' --- Loop Through Each Day of the Month IN FORWARD ORDER ---
    For currentDate = firstDayOfMonth To lastDayOfMonth Step 1
        If Weekday(currentDate, vbMonday) >= 1 And Weekday(currentDate, vbMonday) <= 5 Then 'Only process weekdays
            
            ' --- 1. Personal Entry (created first for the date, so it appears to the left of Non-Entry Hrs) ---
            personalSheetName = "Personal Entry " & Format(currentDate, "M-D-YY")
            If Not SheetExists(personalSheetName, ThisWorkbook) Then
                currentSheetType = "Personal" ' Set current type
                
                ' Check if the sheet name is too long (Excel limit is 31 characters)
                If Len(personalSheetName) > 31 Then
                    MsgBox "Sheet name too long: " & personalSheetName & vbCrLf & "Length: " & Len(personalSheetName) & " characters (max 31)", vbCritical
                    GoTo NextDate
                End If
                
                Debug.Print "About to create new sheet with name: " & personalSheetName & " and copy content from: " & wsPersonalTemplate.Name
                
                ' Create a new worksheet with the correct name
                On Error GoTo SheetRenameError
                Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                ws.name = personalSheetName
                Debug.Print "Successfully created new sheet: " & ws.Name
                
                ' Copy all content from the template to the new sheet
                wsPersonalTemplate.Cells.Copy
                ws.Cells.PasteSpecial Paste:=xlPasteAll
                Application.CutCopyMode = False ' Clear clipboard
                
                ' Set the date
                ws.Range(PERSONAL_ENTRY_DATE_CELL).Value = currentDate
                Debug.Print "Successfully copied content and set date in cell " & PERSONAL_ENTRY_DATE_CELL
                
                ' Force Excel to commit the changes
                DoEvents
                On Error GoTo 0
                
                sheetCreatedCount = sheetCreatedCount + 1
                Debug.Print "SUCCESSFULLY CREATED: " & personalSheetName & ". Count is now: " & sheetCreatedCount ' DEBUG
            Else
                Debug.Print "SKIPPED (already exists): " & personalSheetName ' DEBUG
            End If

            ' --- 2. Non-Entry Hrs (created second for the date, so it appears to the right of Personal Entry) ---
            nonEntrySheetName = "Non-Entry Hrs " & Format(currentDate, "M-D-YY")
            If Not SheetExists(nonEntrySheetName, ThisWorkbook) Then
                currentSheetType = "Non-Entry" ' Set current type
                
                ' Check if the sheet name is too long (Excel limit is 31 characters)
                If Len(nonEntrySheetName) > 31 Then
                    MsgBox "Sheet name too long: " & nonEntrySheetName & vbCrLf & "Length: " & Len(nonEntrySheetName) & " characters (max 31)", vbCritical
                    GoTo NextDate
                End If
                
                Debug.Print "About to create new sheet with name: " & nonEntrySheetName & " and copy content from: " & wsNonEntryTemplate.Name
                
                ' Create a new worksheet with the correct name
                On Error GoTo SheetRenameError
                Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                ws.name = nonEntrySheetName
                Debug.Print "Successfully created new sheet: " & ws.Name
                
                ' Copy all content from the template to the new sheet
                wsNonEntryTemplate.Cells.Copy
                ws.Cells.PasteSpecial Paste:=xlPasteAll
                Application.CutCopyMode = False ' Clear clipboard
                
                ' Set the date
                ws.Range(NON_ENTRY_DATE_CELL).Value = currentDate
                Debug.Print "Successfully copied content and set date in cell " & NON_ENTRY_DATE_CELL
                
                ' Force Excel to commit the changes
                DoEvents
                On Error GoTo 0
                
                sheetCreatedCount = sheetCreatedCount + 1
                Debug.Print "SUCCESSFULLY CREATED: " & nonEntrySheetName & ". Count is now: " & sheetCreatedCount ' DEBUG
            Else
                Debug.Print "SKIPPED (already exists): " & nonEntrySheetName ' DEBUG
            End If
            
NextDate:
        End If
    Next currentDate

    Application.Calculation = xlCalculationAutomatic ' Re-enable calculation
    Application.ScreenUpdating = True
    
    ' Force final save and commit all changes
    DoEvents
    Application.Calculate
    
    ' Save the workbook to ensure all changes are committed
    On Error Resume Next
    ThisWorkbook.Save
    Application.Wait Now + TimeValue("00:00:01") ' Wait 1 second to ensure save completes
    On Error GoTo 0
    
    Debug.Print "Sheet creation process finished. Final count before MsgBox: " & sheetCreatedCount ' DEBUG
    If sheetCreatedCount = 0 Then
        MsgBox "No new sheets were created. All sheets for " & Format(firstDayOfMonth, "MMMM YYYY") & " may already exist.", vbInformation
    Else
        MsgBox sheetCreatedCount & " new sheets created and dates updated for " & Format(firstDayOfMonth, "MMMM YYYY") & ".", vbInformation
    End If
    
    Exit Sub

InvalidInput:
    MsgBox "Invalid input. Please enter numeric values for month and year.", vbExclamation
    Application.Calculation = xlCalculationAutomatic ' Re-enable calculation
    Application.ScreenUpdating = True ' Ensure screen updating is re-enabled on error
    Exit Sub

SheetRenameError:
    Dim attemptedName As String
    If currentSheetType = "Personal" Then
        attemptedName = personalSheetName
    Else
        attemptedName = nonEntrySheetName
    End If
    
    Debug.Print "RENAME ERROR: " & Err.Description & " when trying to rename to: " & attemptedName
    
    MsgBox "Error during sheet renaming." & vbCrLf & _
           "Current sheet name: " & ws.Name & vbCrLf & _
           "Attempted name: " & attemptedName & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "The sheet will be deleted.", vbCritical
    
    ' Clean up the problematic sheet if it was created
    On Error Resume Next
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Debug.Print "Deleted problematic sheet"
    End If
    On Error GoTo 0
    Resume NextDate

SheetCreateError:
    MsgBox "Error creating or renaming sheet: " & vbCrLf & _
           "Sheet name: " & personalSheetName & " or " & nonEntrySheetName & vbCrLf & _
           "Possible causes: Sheet name too long, invalid characters, or workbook protection.", vbCritical
    Application.Calculation = xlCalculationAutomatic ' Re-enable calculation
    Application.ScreenUpdating = True
    Exit Sub

End Sub

' Helper function to check if a sheet exists
Function SheetExists(sheetName As String, Optional ByVal wb As Workbook) As Boolean
    Dim sht As Object
    If wb Is Nothing Then
        Set wb = ThisWorkbook ' Default to ThisWorkbook if no workbook specified
    End If
    On Error Resume Next ' If sheet doesn't exist, sht will remain Nothing
    Set sht = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing ' If sht is an object (sheet exists), Not Nothing is True. Correct.
                                     ' If sht is Nothing (sheet doesn't exist), Not Nothing is False. Correct.
End Function
