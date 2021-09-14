VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PayrollAccrual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Public Args As PayrollArgs
Public m_Dialog As Form_PayPeriodDialog
Private FundNameFilter As String
Private DivisionNameFilter As String
Private AccountCodeFilter As String
Private BocNameFilter As String
Private WorkCodeFilter As String
Private PayPeriodFilter As String
Private FiscalYearFilter As String
Private pAnd As String
Private m_Error As String
Private m_Notification As String





'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Open
'   Parameters:  Integer
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:::
    pAnd = " AND "
    Me.Filter = vbNullString
    Set m_Dialog = New Form_PayPeriodDialog
    Set Args = New PayrollArgs
    DoCmd.OpenForm FormName:="PayPeriodDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("PayPeriodDialog")
    Set Args = m_Dialog.Args
    SetDivisionIcon
    If Not Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Not Args.PayPeriod & "" = "" Then
            DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
            FundNameFilter = "[FundName] = '" & Args.FundName & "'"
            FiscalYearFilter = "[BFY] = '" & Args.BFY & "'"
            PayPeriodFilter = "[PayPeriod] = '" & Args.PayPeriod & "'"
            Me.Filter = DivisionNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & PayPeriodFilter
            Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.AccountCode" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.WorkCode" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.FundName" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.PayPeriod" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.RecordSource = "SELECT * FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Args.PayPeriod & "" = "" Then
            DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
            FiscalYearFilter = "[BFY] = '" & Args.BFY & "'"
            Me.Filter = DivisionNameFilter & _
                pAnd & FiscalYearFilter
            Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.AccountCode" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.WorkCode" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.FundName" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.PayPeriod" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.RecordSource = "SELECT * FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Not Args.PayPeriod & "" = "" Then
            DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
            FiscalYearFilter = "[BFY] = '" & Args.BFY & "'"
            PayPeriodFilter = "[PayPeriod] = '" & Args.PayPeriod & "'"
            Me.Filter = DivisionNameFilter & _
                pAnd & FiscalYearFilter & pAnd & PayPeriodFilter
            Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.AccountCode" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.WorkCode" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.FundName" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.PayPeriod" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.RecordSource = "SELECT * FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
    End If
    If Not Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Args.PayPeriod & "" = "" Then
            DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
            FundNameFilter = "[FundName] = '" & Args.FundName & "'"
            FiscalYearFilter = "[BFY] = '" & Args.BFY & "'"
            PayPeriodFilter = "[PayPeriod] = '" & Args.PayPeriod & "'"
            Me.Filter = DivisionNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
            Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.AccountCode" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.WorkCode" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.FundName" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.PayPeriod" _
                & " FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
            Me.RecordSource = "SELECT * FROM PayrollAccruals" _
                & " WHERE " & Me.Filter
    End If
    DoCmd.Close ObjectType:=acForm, ObjectName:="PayPeriodDialog", Save:=acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        Form_Load
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Section(acHeader).AutoHeight = False
    Me.Section(acHeader).Height = 2
    Me.Section(acDetail).AutoHeight = False
    Me.Section(acDetail).Height = 4
    Me.Section(acFooter).AutoHeight = False
    Me.Section(acFooter).Height = 0.5
    ClearComboBoxValues
    SetComboBoxColors
    SetButtonVisibility
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        Form_Close
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    ClearFilterValues
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Sub-Procedure
'   Name        SetButtonVisibility
'   Parameters  Void
'   Retval      Void
'   Purpose
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
    Me.CompassSpendingQueryButton.Visible = Not Me.CompassSpendingQueryButton.Visible
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
    Me.CompassSpendingQueryButton.Visible = False
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        SetComboBoxColors
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.WorkCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me.WorkCodeComboBox.BackColor = RGB(33, 33, 33)
    Me.WorkCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me.AccountCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me.AccountCodeComboBox.BackColor = RGB(33, 33, 33)
    Me.AccountCodeComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        ClearComboBoxValues
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.Value = vbNullString
    Me.WorkCodeComboBox.Value = vbNullString
    Me.AccountCodeComboBox.Value = vbNullString
    Me.PayPeriodComboBox.Value = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        ClearFilterValues
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    FundNameFilter = vbNullString
    WorkCodeFilter = vbNullString
    AccountCodeFilter = vbNullString
    PayPeriodFilter = vbNullString
    FiscalYearFilter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        WorkCodeComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub WorkCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    WorkCodeFilter = vbNullString
    Me.Filter = vbNullString
    Me.AccountCodeComboBox.Value = vbNullString
    WorkCodeFilter = "[WorkCode] = '" & Me.WorkCodeComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetWorkCodeFilter
    Me.RecordSource = "SELECT * FROM PayrollAccruals WHERE " & Me.Filter
    Me.Requery
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.AccountCode" _
        & " FROM PayrollAccruals" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetWorkCodeFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetWorkCodeFilter() As String
    On Error GoTo ErrorHandler:
    If Not WorkCodeFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter _
            & pAnd & AccountCodeFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & FundNameFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & FundNameFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & FundNameFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not WorkCodeFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        AccountCodeComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub AccountCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    AccountCodeFilter = vbNullString
    AccountCodeFilter = "[AccountCode] = '" & Me.AccountCodeComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetAccountCodeFilter
    Me.RecordSource = "SELECT * FROM PayrollAccruals WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetAccountCodeFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetAccountCodeFilter() As String
    On Error GoTo ErrorHandler:
    If Not AccountCodeFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
            & pAnd & WorkCodeFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & FundNameFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & FundNameFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FundNameFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & FundNameFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        FundNameComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler::
    Me.Filter = vbNullString
    FundNameFilter = vbNullString
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetFundNameFilter
    Me.RecordSource = "SELECT * FROM PayrollAccruals WHERE " & Me.Filter
    Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.WorkCode" _
        & " FROM PayrollAccruals" _
        & " WHERE " & Me.Filter
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.AccountCode" _
        & " FROM PayrollAccruals" _
        & " WHERE " & Me.Filter
    Me.Requery
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
        WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
            & pAnd & WorkCodeFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString And _
        PayPeriodFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & AccountCodeFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetFiscalYearFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetFiscalYearFilter() As String
    On Error GoTo ErrorHandler:
    If Not PayPeriodFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString Then
            GetFiscalYearFilter = PayPeriodFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString Then
            GetFiscalYearFilter = PayPeriodFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString Then
            GetFiscalYearFilter = PayPeriodFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FundNameFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        WorkCodeFilter = vbNullString Then
            GetFiscalYearFilter = PayPeriodFilter _
            & pAnd & FundNameFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString Then
            GetFiscalYearFilter = PayPeriodFilter _
                & pAnd & FundNameFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString Then
            GetFiscalYearFilter = PayPeriodFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString Then
            GetFiscalYearFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not PayPeriodFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not WorkCodeFilter = vbNullString Then
            GetFiscalYearFilter = PayPeriodFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & WorkCodeFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        CalculatorButton_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CompassSpendingQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetDetailedSpendingQuery"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        MenuButton_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        OriginalActionDateTextBox_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub EndDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.EndDateTextBox
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        RefreshButton_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    Me.Filter = DivisionNameFilter
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.AccountCode" _
        & " FROM PayrollAccruals" _
        & " WHERE " & Me.Filter
    Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.WorkCode" _
        & " FROM PayrollAccruals" _
        & " WHERE " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.FundName" _
        & " FROM PayrollAccruals" _
        & " WHERE " & Me.Filter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.PayPeriod" _
        & " FROM PayrollAccruals" _
        & " WHERE " & Me.Filter
    Me.RecordSource = "SELECT * FROM PayrollAccruals" _
        & " WHERE" & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            PayPeriodComboBox_Change
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayPeriodComboBox_Change()
    On Error GoTo ErrorHandler:
    PayPeriodFilter = ""
    Me.Filter = ""
    PayPeriodFilter = "[PayPeriod] = '" & Me.PayPeriodComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetPayPeriodFilter
    Me.RecordSource = "SELECT * FROM PayrollAccruals WHERE " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.FundName FROM PayrollAccruals WHERE " & Me.Filter
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.AccountCode FROM PayrollAccruals WHERE " & Me.Filter
    Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PayrollAccruals.WorkCode FROM PayrollAccruals WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetPayPeriodFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetPayPeriodFilter() As String
    On Error GoTo ErrorHandler:
    If Not PayPeriodFilter = "" And _
        Not FundNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & FundNameFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        FundNameFilter = "" And _
        WorkCodeFilter = "" And _
        AccountCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter
    End If
    If Not PayPeriodFilter = "" And _
        Not FundNameFilter = "" And _
        WorkCodeFilter = "" And _
        AccountCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & FundNameFilter
    End If
    If Not PayPeriodFilter = "" And _
        Not FundNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        AccountCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & FundNameFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        FundNameFilter = "" And _
        WorkCodeFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        FundNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        FundNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        AccountCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        Not FundNameFilter = "" And _
        WorkCodeFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            StartDateTextBox_Click
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub StartDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.StartDateTextBox
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
    On Error GoTo ErrorHandler::
    If Not Args.RcCode & "" = "" Then
        Me.Caption = "Division Payroll Accruals"
        Me.DivisionTitle.Caption = Args.DivisionName
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
    Dim m_Fields() As String
    m_Fields() = GetColumnNames
    Dim m_BudgetPath As BudgetPath
    Dim m_Excel As Excel.Application
    Dim m_Allocations As Excel.Workbook
    Dim m_Worksheet As Excel.Worksheet
    Dim m_List As Excel.ListObject
    Dim m_Range As Excel.Range
    Dim m_Cell As Object
    Dim m_Header As Excel.Range
    Dim m_Start As Excel.Range
    Dim m_End As Excel.Range
    Dim field As DAO.field
    Dim m_Data As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Set m_BudgetPath = New BudgetPath
    Set m_Excel = CreateObject("Excel.Application")
    Set m_Allocations = m_Excel.Workbooks.Open(m_BudgetPath.ReportTemplate)
    m_Allocations.Worksheets(2).Visible = False
    Set m_Worksheet = m_Allocations.Worksheets(1)
    m_Worksheet.Name = "Payroll Accruals"
    m_Worksheet.Cells.HorizontalAlignment = xlHAlignLeft
    m_Worksheet.Cells.Font.Name = "Source Code Pro"
    m_Worksheet.Cells.Font.Size = 8
    Set m_Start = m_Worksheet.Cells(1, 1)
    Set m_End = m_Worksheet.Cells(1, UBound(m_Fields) - 1)
    Set m_Header = m_Worksheet.Range(m_Start, m_End)
    m_Header.HorizontalAlignment = xlHAlignLeft
    m_Header.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    m_Header.Font.Name = "Source Code Pro"
    m_Header.Font.Color = vbBlack
    m_Header.Font.Bold = True
    m_Header.Font.Size = 8
    For i = LBound(m_Fields) To UBound(m_Fields)
        m_Header.Cells(i + 1).Value = m_Fields(i)
    Next i
    Set m_Data = Me.Recordset
    m_Data.Filter = Me.Filter
    m_Worksheet.Cells(2, 1).CopyFromRecordset m_Data
    m_Header.Font.Color = vbBlack
    m_Excel.WindowState = xlMaximized
    m_Excel.Visible = True
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




