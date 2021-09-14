VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Guidance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private mError As String







'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Caption = "Guidance Menu"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            AdviceOfAllowanceButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AdviceOfAllowanceButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, Me.Name
    Application.FollowHyperlink CurrentProject.path & "\etc\docs\rmds\AdviceOfAllowance21.pdf"
    DoCmd.OpenForm FormName:="Main", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            FundsControlManualButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundsControlManualButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, Me.Name
    Application.FollowHyperlink CurrentProject.path & "\etc\docs\rmds\FundsControlManual.pdf"
    DoCmd.OpenForm FormName:="Main", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            MoneyAndFinanceButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MoneyAndFinanceButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, Me.Name
    Application.FollowHyperlink CurrentProject.path & "\etc\docs\rmds\CFR31.pdf"
    DoCmd.OpenForm FormName:="Main", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ObjectClassManualButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ObjectClassManualButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, Me.Name
    Application.FollowHyperlink CurrentProject.path & "\etc\docs\rmds\ObjectClassManual.pdf"
    DoCmd.OpenForm FormName:="Main", WindowMode:=acDialog
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





