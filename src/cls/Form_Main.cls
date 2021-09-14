VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database




Private m_Error As String







'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Caption = "Main Menu"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub






'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        GuidanceButton_Click()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub GuidanceButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close ObjectType:=acForm, ObjectName:="Main"
    DoCmd.OpenForm FormName:="Guidance", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub







'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        MaintenanceButton_Click()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MaintenanceButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close ObjectType:=acForm, ObjectName:="Main"
    DoCmd.OpenForm FormName:="Maintenance", WindowMode:=acDialog
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





