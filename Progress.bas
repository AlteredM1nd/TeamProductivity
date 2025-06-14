Option Explicit

Private Type ProgressInfo
    Phase As String
    Current As Long
    Total As Long
    StartTime As Date
    Description As String
End Type

Private progressData As ProgressInfo

Public Sub InitProgress(phase As String, total As Long, Optional description As String = "")
    With progressData
        .Phase = phase
        .Current = 0
        .Total = total
        .StartTime = Now
        .Description = description
    End With
    UpdateStatusBar
End Sub

Public Sub UpdateProgress(Optional increment As Long = 1)
    progressData.Current = progressData.Current + increment
    UpdateStatusBar
End Sub

Private Sub UpdateStatusBar()
    Dim percentComplete As Double
    Dim timeElapsed As Double
    Dim estimatedTotal As Double
    Dim remainingTime As Double
    
    With progressData
        percentComplete = .Current / .Total
        timeElapsed = (Now - .StartTime) * 86400 ' Convert to seconds
        
        If .Current > 0 Then
            estimatedTotal = timeElapsed / (.Current / .Total)
            remainingTime = estimatedTotal - timeElapsed
            
            Application.StatusBar = .Phase & ": " & _
                                  Format(percentComplete, "0.0%") & " - " & _
                                  .Description & " - " & _
                                  "Remaining: " & FormatTimespan(remainingTime)
        Else
            Application.StatusBar = .Phase & ": Starting... - " & .Description
        End If
    End With
    DoEvents
End Sub

Public Sub EndProgress()
    Application.StatusBar = False
End Sub

Private Function FormatTimespan(seconds As Double) As String
    Dim hours As Long
    Dim minutes As Long
    
    hours = Int(seconds / 3600)
    minutes = Int((seconds Mod 3600) / 60)
    
    If hours > 0 Then
        FormatTimespan = hours & "h " & minutes & "m"
    Else
        FormatTimespan = minutes & "m"
    End If
End Function

Public Sub LogProgress()
    Dim wsStatus As Worksheet
    
    On Error Resume Next
    Set wsStatus = ThisWorkbook.Sheets("Status")
    If wsStatus Is Nothing Then
        Set wsStatus = ThisWorkbook.Sheets.Add
        wsStatus.Name = "Status"
        wsStatus.Range("A1:E1").Value = Array("Timestamp", "Phase", "Progress", "Time Elapsed", "Description")
    End If
    On Error GoTo 0
    
    With progressData
        wsStatus.Cells(wsStatus.Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 5).Value = _
            Array(Now, .Phase, Format(.Current / .Total, "0.0%"), _
                  FormatTimespan((Now - .StartTime) * 86400), .Description)
    End With
End Sub
