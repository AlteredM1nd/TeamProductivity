Option Explicit

Private Type BackupInfo
    Path As String
    Timestamp As Date
    Description As String
End Type

Public Sub CreateBackup(Optional description As String = "")
    Dim backupInfo As BackupInfo
    
    ' Create Backups folder if it doesn't exist
    Dim backupFolder As String
    backupFolder = ThisWorkbook.Path & "\Backups"
    
    On Error Resume Next
    MkDir backupFolder
    On Error GoTo 0
    
    ' Generate backup filename with timestamp
    With backupInfo
        .Timestamp = Now
        .Description = description
        .Path = backupFolder & "\" & _
                Format(.Timestamp, "yyyy-mm-dd_hhmmss") & "_backup.xlsm"
    End With
    
    ' Save backup
    ThisWorkbook.SaveCopyAs backupInfo.Path
    
    ' Log backup details
    LogBackup backupInfo
End Sub

Private Sub LogBackup(backupInfo As BackupInfo)
    Dim wsBackupLog As Worksheet
    
    On Error Resume Next
    Set wsBackupLog = ThisWorkbook.Sheets("BackupLog")
    If wsBackupLog Is Nothing Then
        Set wsBackupLog = ThisWorkbook.Sheets.Add
        wsBackupLog.Name = "BackupLog"
        wsBackupLog.Range("A1:D1").Value = Array("Timestamp", "Backup Path", "Description", "File Size (KB)")
    End If
    On Error GoTo 0
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim backupFile As Object
    Set backupFile = fso.GetFile(backupInfo.Path)
    
    wsBackupLog.Cells(wsBackupLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 4).Value = _
        Array(backupInfo.Timestamp, backupInfo.Path, backupInfo.Description, _
              Format(backupFile.Size / 1024, "#,##0.0"))
              
    Set backupFile = Nothing
    Set fso = Nothing
End Sub

Public Sub CleanupOldBackups(Optional daysToKeep As Integer = 30)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim backupFolder As String
    backupFolder = ThisWorkbook.Path & "\Backups"
    
    If Not fso.FolderExists(backupFolder) Then Exit Sub
    
    Dim folder As Object
    Set folder = fso.GetFolder(backupFolder)
    
    Dim file As Object
    Dim cutoffDate As Date
    cutoffDate = Date - daysToKeep
    
    For Each file In folder.Files
        If file.DateCreated < cutoffDate Then
            file.Delete
        End If
    Next file
    
    Set folder = Nothing
    Set fso = Nothing
End Sub
