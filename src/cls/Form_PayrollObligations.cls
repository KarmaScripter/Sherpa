VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PayrollObligations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As PayrollArgs
Public m_Dialog As Form_PayrollDialog
Private FundNameFilter As String
Private DivisionNameFilter As String
Private PayPeriodFilter As String
Private WorkCodeFilter As String
Private ProgramProjectNameFilter As String
Private StartDateFilter As String
Private EndDateFilter As String
Public pAnd As String
Public calendar As Form_BudgetCalendar
Private m_Error As String
Private m_Notification As String




'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        Form_Open
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler::
    Set m_Dialog = New Form_PayrollDialog
    DoCmd.OpenForm FormName:="PayrollDialog", _
        WindowMode:=acDialog
    Set m_Dialog = Forms("PayrollDialog")
    pAnd = " AND "
    Set Args = New PayrollArgs
    Set Args = m_Dialog.Args
    If Not Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not StartDateFilter & "" = "" And _
        Not EndDateFilter & "" = "" Then
        FundNameFilter = "[FundName] = '" & Args.FundName & "'"
        DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
        StartDateFilter = "[StartDate] = '" & Args.StartDate & "'"
        EndDateFilter = "[EndDate] = '" & Args.EndDate & "'"
        Me.Filter = DivisionNameFilter & pAnd & FundNameFilter & _
            pAnd & StartDateFilter & pAnd & EndDateFilter
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        Not StartDateFilter & "" = "" And _
        Not EndDateFilter & "" = "" Then
        DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
        StartDateFilter = "[StartDate] = '" & Args.StartDate & "'"
        EndDateFilter = "[EndDate] = '" & Args.EndDate & "'"
        Me.Filter = DivisionNameFilter & pAnd & _
            StartDateFilter & pAnd & EndDateFilter
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        StartDateFilter & "" = "" And _
        Not EndDateFilter & "" = "" Then
        DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
        EndDateFilter = "[EndDate] = '" & Args.EndDate & "'"
        Me.Filter = DivisionNameFilter & pAnd & FundNameFilter & _
            pAnd & StartDateFilter & pAnd & EndDateFilter
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        StartDateFilter & "" = "" And _
        EndDateFilter & "" = "" Then
        DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
        Me.Filter = DivisionNameFilter
    End If
    If Not Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        StartDateFilter & "" = "" And _
        EndDateFilter & "" = "" Then
        FundNameFilter = "[FundName] = '" & Args.FundName & "'"
        DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
        Me.Filter = DivisionNameFilter & pAnd & FundNameFilter
    End If
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.FundName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.ProgramProjectName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollObligations.PayPeriod" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollObligations.WorkCode" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.RecordSource = "SELECT * FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    DoCmd.Close ObjectType:=acForm, ObjectName:=m_Dialog.Name, Save:=acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Section(acHeader).AutoHeight = False
    Me.Section(acHeader).Height = 2
    Me.Section(acDetail).AutoHeight = False
    Me.Section(acDetail).Height = 4
    Me.Section(acFooter).AutoHeight = False
    Me.Section(acFooter).Height = 0.5
    InitButtonVisibility
    pAnd = " AND "
    SetComboBoxColors
    ClearComboBoxValues
    ClearFilterValues
    SetDivisionIcon
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    FundNameFilter = vbNullString
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetFundNameFilter
    Me.RecordSource = "SELECT * FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.Requery
    Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollObligations.WorkCode" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.ProgramProjectName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollObligations.PayPeriod" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFundNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetFundNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not FundNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = PayPeriodFilter & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = PayPeriodFilter & pAnd & WorkCodeFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = PayPeriodFilter & pAnd & WorkCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = WorkCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            Me.Filter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = WorkCodeFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = PayPeriodFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        PayrollQueryButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayrollQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "PayrollAccrualQuery"
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
Private Sub ProgramProjectNameComboBox_Change()
    On Error GoTo ErrorHandler:
    ProgramProjectNameFilter = vbNullString
    Me.Filter = vbNullString
    ProgramProjectNameFilter = "[AccountCode] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetProgramProjectNameFilter
    Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollObligations.WorkCode" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.FundName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetProgramProjectNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetProgramProjectNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not ProgramProjectNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = PayPeriodFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = WorkCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = WorkCodeFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = PayPeriodFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        PayPeriodComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayPeriodComboBox_Change()
    On Error GoTo ErrorHandler:
    PayPeriodFilter = vbNullString
    Me.Filter = vbNullString
    PayPeriodFilter = "[PayPeriod] = '" & Me.PayPeriodComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetPayPeriodFilter
    Me.RecordSource = "SELECT * FROM PayrollObligations WHERE " & Me.Filter
    Me.Requery
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.FundName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.ProgramProjectName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollObligations.WorkCode" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFundNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetPayPeriodFilter() As String
    On Error GoTo ErrorHandler:
    If Not PayPeriodFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetPayPeriodFilter = PayPeriodFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetPayPeriodFilter = WorkCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetPayPeriodFilter = FundNameFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetPayPeriodFilter = WorkCodeFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        WorkCodeComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub WorkCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    WorkCodeFilter = vbNullString
    Me.Filter = vbNullString
    WorkCodeFilter = "[WorkCode] = '" & Me.WorkCodeComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetWorkCodeFilter
    Me.RecordSource = "SELECT * FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.Requery
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.FundName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations..ProgramProjectName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollObligations.PayPeriod" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetWorkCodeFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetWorkCodeFilter() As String
    On Error GoTo ErrorHandler:
    If Not WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetWorkCodeFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetWorkCodeFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetWorkCodeFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetWorkCodeFilter = FundNameFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetWorkCodeFilter = PayPeriodFilter _
                & pAnd & FundNameFilter _
                & pAnd & WorkCodeFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetComboBoxColors
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.WorkCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me.WorkCodeComboBox.BackColor = RGB(33, 33, 33)
    Me.WorkCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me.ProgramProjectNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.ProgramProjectNameComboBox.BackColor = RGB(33, 33, 33)
    Me.ProgramProjectNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.PayPeriodComboBox.ForeColor = RGB(255, 255, 255)
    Me.PayPeriodComboBox.BackColor = RGB(33, 33, 33)
    Me.PayPeriodComboBox.BorderColor = RGB(68, 114, 196)
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
    Me.FundNameComboBox.Value = vbNullString
    Me.WorkCodeComboBox.Value = vbNullString
    Me.ProgramProjectNameComboBox.Value = vbNullString
    Me.PayPeriodComboBox.Value = vbNullString
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
    Me.AddButton.Visible = Not Me.AddButton.Visible
    Me.FirstButton.Visible = Not Me.FirstButton.Visible
    Me.PreviousButton.Visible = Not Me.PreviousButton.Visible
    Me.NextButton.Visible = Not Me.NextButton.Visible
    Me.LastButton.Visible = Not Me.LastButton.Visible
    Me.EditButton.Visible = Not Me.EditButton.Visible
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.DeleteButton.Visible = Not Me.DeleteButton.Visible
    Me.CalculatorButton.Visible = Not Me.CalculatorButton.Visible
    Me.ExcelButton.Visible = Not Me.ExcelButton.Visible
    Me.UndoButton.Visible = Not Me.UndoButton.Visible
    Me.SaveButton.Visible = Not Me.SaveButton.Visible
    Me.DataButton.Visible = Not Me.DataButton.Visible
    Me.PayrollQueryButton.Visible = Not Me.PayrollQueryButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Timer()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Timer()
    On Error GoTo ErrorHandler:
    HideButtons
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Hides buttons
'   Parameters:  Void
'   Purpose:     Toggles the toolbar button on/off
'---------------------------------------------------------------------------------
Private Sub HideButtons()
    On Error GoTo ErrorHandler:
    Me.AddButton.Visible = False
    Me.DataButton.Visible = False
    Me.FirstButton.Visible = False
    Me.PreviousButton.Visible = False
    Me.NextButton.Visible = False
    Me.LastButton.Visible = False
    Me.EditButton.Visible = False
    Me.RefreshButton.Visible = False
    Me.DeleteButton.Visible = False
    Me.CalculatorButton.Visible = False
    Me.ExcelButton.Visible = False
    Me.UndoButton.Visible = False
    Me.SaveButton.Visible = False
    Me.PayrollQueryButton.Visible = False
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        InitButtonVisibility
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub InitButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.AddButton.Visible = False
    Me.FirstButton.Visible = False
    Me.PreviousButton.Visible = False
    Me.NextButton.Visible = False
    Me.LastButton.Visible = False
    Me.EditButton.Visible = False
    Me.RefreshButton.Visible = False
    Me.DeleteButton.Visible = False
    Me.CalculatorButton.Visible = False
    Me.ExcelButton.Visible = False
    Me.UndoButton.Visible = False
    Me.SaveButton.Visible = False
    Me.DataButton.Visible = False
    Me.PayrollQueryButton.Visible = False
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
    WorkCodeFilter = vbNullString
    ProgramProjectNameFilter = vbNullString
    FundNameFilter = vbNullString
    PayPeriodFilter = vbNullString
    Me.Filter = DivisionNameFilter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearFilterValues
    ClearComboBoxValues
    Me.Filter = DivisionNameFilter
    Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollObligations.WorkCode" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.ProgramProjectName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.FundName" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollObligations.PayPeriod" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.RecordSource = "SELECT * FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        CalculatorButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        MenuButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetDivisionIcon
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetDivisionIcon()
    On Error GoTo ErrorHandler:::
    If Not Args.RcCode = vbNullString Then
        Select Case Args.RcCode
            Case "06A"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\ORA.png"
            Case "06B"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\LCARD.png"
            Case "06C"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\MSD.png"
            Case "06D"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\ORC.png"
            Case "06F"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\EJ.png"
            Case "06G"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\WCF.png"
            Case "06H"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\LSASD.png"
            Case "06J"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\ARD.png"
            Case "06K"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\WD.png"
            Case "06L"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\SEMD.png"
            Case "06M"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\ECAD.png"
            Case "06N"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\WSA.png"
            Case "06R"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\MSR.png"
            Case "06X"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\XA.png"
        End Select
    Else
        Me.DivisionIcon.Picture = _
            CurrentProject.Path & "\etc\png\AppIcons\interface\ui\Reports.png"
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetColumnNames
'   Parameters:  Void
'   Retval:      Collection
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetColumnNames() As String()
    On Error GoTo ErrorHandler:
    Dim m_Fields As Collection
    Dim field As DAO.field
    Dim m_Data As DAO.Recordset
    Set m_Data = Me.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim m_Array() As String
    j = m_Data.Fields.count - 1
    ReDim m_Array(j)
    For i = LBound(m_Array()) To UBound(m_Array())
        If Not m_Data.Fields(i).Name & "" = "" Then
            m_Array(i) = m_Data.Fields(i).Name
        End If
    Next i
    GetColumnNames = m_Array()
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetReportData
'   Parameters:  Void
'   Retval:      DAO Recordset
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetReportData() As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim m_Data As DAO.Recordset
    Set m_Data = Me.Recordset
    m_Data.Filter = Me.Filter
    Set GetReportData = m_Data
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExcelButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExcelButton_Click()
    On Error GoTo ErrorHandler:
    Dim p_Fields() As String
    p_Fields() = GetColumnNames
    Dim p_BudgetPath As BudgetPath
    Dim p_Excel As Excel.Application
    Dim p_Allocations As Excel.Workbook
    Dim p_Worksheet As Excel.Worksheet
    Dim p_List As Excel.ListObject
    Dim p_Range As Excel.Range
    Dim p_Cell As Object
    Dim p_Header As Excel.Range
    Dim p_Start As Excel.Range
    Dim p_End As Excel.Range
    Dim field As DAO.field
    Dim p_Data As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Set p_BudgetPath = New BudgetPath
    Set p_Excel = CreateObject("Excel.Application")
    Set p_Allocations = p_Excel.Workbooks.Open(p_BudgetPath.ReportTemplate)
    p_Allocations.Worksheets(2).Visible = False
    Set p_Worksheet = p_Allocations.Worksheets(1)
    p_Worksheet.Name = "Payroll Obligations"
    p_Worksheet.Cells.HorizontalAlignment = xlHAlignLeft
    p_Worksheet.Cells.Font.Name = "Source Code Pro"
    p_Worksheet.Cells.Font.Size = 8
    Set p_Start = p_Worksheet.Cells(1, 1)
    Set p_End = p_Worksheet.Cells(1, UBound(p_Fields) - 1)
    Set p_Header = p_Worksheet.Range(p_Start, p_End)
    p_Header.HorizontalAlignment = xlHAlignLeft
    p_Header.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    p_Header.Font.Name = "Source Code Pro"
    p_Header.Font.Color = vbBlack
    p_Header.Font.Bold = True
    p_Header.Font.Size = 8
    For i = LBound(p_Fields) To UBound(p_Fields)
        p_Header.Cells(i + 1).Value = p_Fields(i)
    Next i
    Set p_Data = Me.Recordset
    p_Data.Filter = Me.Filter
    p_Worksheet.Cells(2, 1).CopyFromRecordset p_Data
    p_Header.Font.Color = vbBlack
    p_Excel.WindowState = xlMaximized
    p_Excel.Visible = True
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SaveButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UndoButton_Click()
    On Error GoTo ErrorHandler:
    If Me.Dirty Then
        DoCmd.RunCommand acCmdUndo
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SaveButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SaveButton_Click()
    On Error GoTo ErrorHandler:
    If Me.Dirty Then
        DoCmd.RunCommand acCmdSave
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DeleteButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DeleteButton_Click()
    On Error GoTo ErrorHandler:
    If Me.Dirty Then
        DoCmd.RunCommand acCmdDelete
    End If
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





