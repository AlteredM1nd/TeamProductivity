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
    Const TEMPLATE_PERSONAL_NAME As String = "Personal Entry" '<<< CHANGE IF YOUR TEMPLATE NAME IS DIFFERENT
    Const TEMPLATE_NON_ENTRY_NAME As String = "Non-Entry Hrs"  '<<< CHANGE IF YOUR TEMPLATE NAME IS DIFFERENT

    ' --- DEFINE CELLS FOR DATE INSERTION ---
    Const PERSONAL_ENTRY_DATE_CELL As String = "A2"
    Const NON_ENTRY_DATE_CELL As String = "A1"

    Dim wsPersonalTemplate As Worksheet
    Dim wsNonEntryTemplate As Worksheet

    ' --- Get User Input ---
    On Error GoTo InvalidInput
    inputMonth = InputBox("Enter the month number (e.g., 6 for June):", "Input Month")
    If inputMonth = "" Then Exit Sub ' User pressed Cancel
    targetMonth = CInt(inputMonth)

    inputYear = InputBox("Enter the year (e.g., 2025):", "Input Year")
    If inputYear = "" Then Exit Sub ' User pressed Cancel
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
    sheetCreatedCount = 0 ' Initialize count ONCE before the loop

    Debug.Print "Starting sheet creation process. Initial count: " & sheetCreatedCount ' DEBUG

    ' --- Loop Through Each Day of the Month IN REVERSE ---
    For currentDate = lastDayOfMonth To firstDayOfMonth Step -1
        If Weekday(currentDate, vbMonday) >= 1 And Weekday(currentDate, vbMonday) <= 5 Then 'Only process weekdays
            
            ' --- 1. Non-Entry Hrs (copied first for the date, so it appears to the right of Personal Entry) ---
            nonEntrySheetName = "Non-Entry Hrs " & Format(currentDate, "M-D-YY")
            If Not SheetExists(nonEntrySheetName, ThisWorkbook) Then
                wsNonEntryTemplate.Copy Before:=ThisWorkbook.Sheets(1)
                Set ws = ActiveSheet
                ws.name = nonEntrySheetName
                ws.Range(NON_ENTRY_DATE_CELL).Value = currentDate
                ' Optional: ws.Range(NON_ENTRY_DATE_CELL).NumberFormat = "m/d/yyyy"
                sheetCreatedCount = sheetCreatedCount + 1
                Debug.Print "CREATED: " & nonEntrySheetName & ". Count is now: " & sheetCreatedCount ' DEBUG
            Else
                Debug.Print "SKIPPED (already exists): " & nonEntrySheetName ' DEBUG
            End If

            ' --- 2. Personal Entry (copied second for the date, so it appears to the left of Non-Entry Hrs) ---
            personalSheetName = "Personal Entry " & Format(currentDate, "M-D-YY")
            If Not SheetExists(personalSheetName, ThisWorkbook) Then
                wsPersonalTemplate.Copy Before:=ThisWorkbook.Sheets(1)
                Set ws = ActiveSheet
                ws.name = personalSheetName
                ws.Range(PERSONAL_ENTRY_DATE_CELL).Value = currentDate
                ' Optional: ws.Range(PERSONAL_ENTRY_DATE_CELL).NumberFormat = "m/d/yyyy"
                sheetCreatedCount = sheetCreatedCount + 1
                Debug.Print "CREATED: " & personalSheetName & ". Count is now: " & sheetCreatedCount ' DEBUG
            Else
                Debug.Print "SKIPPED (already exists): " & personalSheetName ' DEBUG
            End If
            
        End If
    Next currentDate

    Application.ScreenUpdating = True
    Debug.Print "Sheet creation process finished. Final count before MsgBox: " & sheetCreatedCount ' DEBUG
    MsgBox sheetCreatedCount & " new sheets created and dates updated for " & Format(firstDayOfMonth, "MMMM YYYY") & ".", vbInformation
    
    Exit Sub

InvalidInput:
    MsgBox "Invalid input. Please enter numeric values for month and year.", vbExclamation
    Application.ScreenUpdating = True ' Ensure screen updating is re-enabled on error

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
