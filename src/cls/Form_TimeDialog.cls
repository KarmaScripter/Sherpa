VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TimeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Public Args As BudgetArgs
Private m_StartDate As String
Private m_EndDate As String
Private m_FundName As String
Private m_BocName As String
Private pAnd As String
Private m_Error As String





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    Set Args = New BudgetArgs
    SetComboBoxColors
    pAnd = " AND "
    ClearComboBoxValues
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Allocations.FundName FROM Allocations;"
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Allocations.BocName FROM Allocations;"
    Me.StartDateTextBox.Value = #10/1/2020#
    Me.EndDateTextBox.Value = Date
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        BocNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.FundName = Me.FundNameComboBox.Value
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Allocations.FundName" _
        & " FROM Allocations WHERE" _
        & " Allocations.FundName = '" & Me.FundNameComboBox.Value & "'"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProgramProjectNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.BocName = Me.BocNameComboBox.Value
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub StartDateTextBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    m_StartDate = vbNullString
    m_StartDate = Me.StartDateTextBox.text
    Args.StartDate = Me.StartDateTextBox.text
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub EndDateTextBox_Change()
    On Error GoTo ErrorHandler:
    m_EndDate = vbNullString
    m_EndDate = Me.EndDateTextBox.text
    Args.EndDate = Me.EndDateTextBox.text
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AcceptButton_Click()
    Me.Visible = Not Me.Visible
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ClearTextBoxValues
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetComboBoxColors
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me!BocNameComboBox.ForeColor = RGB(255, 255, 255)
    Me!BocNameComboBox.BackColor = RGB(33, 33, 33)
    Me!BocNameComboBox.BorderColor = RGB(68, 114, 196)
    Me!FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me!FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me!FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me!StartDateTextBox.ForeColor = RGB(255, 255, 255)
    Me!StartDateTextBox.BackColor = RGB(33, 33, 33)
    Me!StartDateTextBox.BorderColor = RGB(68, 114, 196)
    Me!EndDateTextBox.ForeColor = RGB(255, 255, 255)
    Me!EndDateTextBox.BackColor = RGB(33, 33, 33)
    Me!EndDateTextBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetButtonVisibility
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.ExitButton.SetFocus
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.AcceptButton.Visible = Not Me.AcceptButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearComboBoxValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.BocNameComboBox.Value = vbNullString
    Me.FundNameComboBox.Value = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearTextBoxValues()
    On Error GoTo ErrorHandler:
    Me.StartDateTextBox.Value = vbNullString
    Me.EndDateTextBox.Value = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearFilterValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    m_EndDate = vbNullString
    m_StartDate = vbNullString
    m_FundName = vbNullString
    m_BocName = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExitButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    Set Args = Nothing
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExitButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub StartDateTextBox_Enter()
    On Error GoTo ErrorHandler:
    m_StartDate = vbNullString
    Dim calendar As New Form_BudgetCalendar
    DoCmd.OpenForm FormName:="BudgetCalendar", WindowMode:=acDialog, OpenArgs:=Me.StartDateTextBox
    Set calendar = Forms("BudgetCalendar")
    Me.StartDateTextBox.Value = calendar.txtCalendarDate
    m_StartDate = Me.StartDateTextBox.Value
    Args.StartDate = m_StartDate
    DoCmd.Close ObjectType:=acForm, ObjectName:="BudgetCalendar", Save:=acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExitButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub EndDateTextBox_Enter()
    On Error GoTo ErrorHandler:
    Dim calendar As New Form_BudgetCalendar
    DoCmd.OpenForm FormName:="BudgetCalendar", WindowMode:=acDialog, _
        OpenArgs:=Me.EndDateTextBox
    Set calendar = Forms("BudgetCalendar")
    Me.EndDateTextBox.Value = calendar.txtCalendarDate
    m_EndDate = Me.EndDateTextBox.Value
    Args.EndDate = m_EndDate
    DoCmd.Close ObjectType:=acForm, ObjectName:="BudgetCalendar", _
        Save:=acSaveNo
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





