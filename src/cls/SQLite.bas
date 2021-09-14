Attribute VB_Name = "SQLite"



Private m_SQLitePath As String
Private m_SQLiteArg As String
Private m_ShellArgPath As String
Private m_ProjectPath As String
Private m_Error As String




'----------------------------------------------------------------------------------
'   Type:        SubProcedure
'   Name:        Run
'   Parameters:  Void
'   Retval:      Process
'   Purpose:     Launches the SQLite Editor 'SQLiteDatabaseBrowserPortable.exe'
'----------------------------------------------------------------------------------
Public Sub Run()
    On Error GoTo ErrorHandler:
    m_ProjectPath = CurrentProject.Path
    m_SQLitePath = Replace(m_ProjectPath, "accdb\models", "sqlite\gui\SQLiteDatabaseBrowserPortable.exe")
    m_SQLiteArg = Replace(m_ProjectPath, "accdb\models", "sqlite\gui\models\Data.db")
    m_ShellArgPath = m_SQLitePath & " " & m_SQLiteArg
    vPid = Shell(m_ShellArgPath, 3)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProcessError
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProcessError()
    If Err.Number <> 0 Then
        m_Error = "Source:      " & Err.Source _
            & vbCrLf & "Number:     " & Err.Number _
            & vbCrLf & "Issue:      " & Err.Description
    End If
    MessageFactory.ShowError (m_Error)
    Err.Clear
End Sub



