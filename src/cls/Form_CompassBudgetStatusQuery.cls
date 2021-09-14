VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CompassBudgetStatusQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_Error As String
Private m_Caller As String



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            Form_Open()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    If Not IsNull(Me.OpenArgs) Then
        m_Caller = CStr(Me.OpenArgs)
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ExitButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
    Me.Caption = "CDW Compass Budget Status Query"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ExitButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.BudgetDetailedSpendingButton.Visible = Not Me.BudgetDetailedSpendingButton.Visible
    Me.CompassAdvancedQueryButton.Visible = Not Me.CompassAdvancedQueryButton.Visible
    Me.CompassReprogrammingQueryButton.Visible = Not Me.CompassReprogrammingQueryButton.Visible
    Me.CompassTravelQueryButton.Visible = Not Me.CompassTravelQueryButton.Visible
    Me.PayrollAccrualsButton.Visible = Not Me.PayrollAccrualsButton.Visible
    Me.BudgetStatusQueryButton.Visible = Not Me.BudgetStatusQueryButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ExitButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    Me.Visible = False
    If Not m_Caller & "" = "" Then
        DoCmd.OpenForm FormName:=m_Caller, WindowMode:=acDialog
    Else
        DoCmd.OpenForm FormName:="Main", WindowMode:=acDialog
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ExitButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
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








