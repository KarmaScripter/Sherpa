VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Payroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


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
    Me.Caption = "Payroll Menu"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DivisionPayrollButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PayrollUtilization", _
        WindowMode:=acDialog, _
        OpenArgs:="[BocCode] ='10' AND [BudgetLevel] = '7'"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FteButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="FullTimeAllocations", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayrollAccrualButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PayrollAccrual", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayrollAuthorityButton__AuthorityReportButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PayrollAllocations", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayrollObligationButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PayrollObligations", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub TimeAttendanceButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PeoplePlusHours", WindowMode:=acDialog
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





