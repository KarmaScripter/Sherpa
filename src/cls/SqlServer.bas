Attribute VB_Name = "SqlServer"
Option Compare Database



Public pid As Variant
Private m_CompactPath As String
Private m_CompactArg As String
Private m_ShellArgPath As String
Private m_Error As String



'----------------------------------------------------------------------------------
'   Type        SubProcedure
'   Name        Calculate
'   Parameters  Void
'   Purpose     Launches the Windows 10 calculator 'calc.exe'
'----------------------------------------------------------------------------------
Public Sub RunCompact()
    On Error GoTo ErrorHandler:
    m_CompactPath = Replace(CurrentProject.Path, "accdb\models", "sqlce\gui\CompactView.exe")
    m_CompactArg = " " & Replace(CurrentProject.Path, "accdb\models", "sqlce\gui\Data.sdf")
    m_ShellArgPath = m_CompactPath & m_CompactArg
    vPID = Shell(mShellArgPath, 3)
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
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
End Sub



