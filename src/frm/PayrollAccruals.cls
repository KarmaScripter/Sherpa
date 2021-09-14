Option Compare Database
Option Explicit



Public Args As PayrollArgs
Public mDialog As Form_PayPeriodDialog
Private FundNameFilter As String
Private DivisionNameFilter As String
Private AccountCodeFilter As String
Private BocNameFilter As String
Private WorkCodeFilter As String
Private PayPeriodFilter As String
Private FiscalYearFilter As String
Private pAnd As String
Private mError As String
Private mNotification As String




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
    Set mDialog = New Form_PayPeriodDialog
    Set Args = New PayrollArgs
    DoCmd.OpenForm FormName:="PayPeriodDialog", WindowMode:=acDialog
    Set mDialog = Forms("PayPeriodDialog")
    Set Args = mDialog.Args
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member:     Form_Open()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: Form_Close()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: SetButtonVisibility()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: Form_Timer()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:      FinanceObligations" _
            & vbCrLf & "Member:     HideButtons()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: ClearComboBoxValues()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: ClearFilterValues()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   AllocationForm" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   AllocationForm" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: WorkCodeComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: GetWorkCodeFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: AccountCodeComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: GetAccountCodeFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: FundNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: GetFundNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: GetFiscalYearFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: CalculatorButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: CompassSpendingQueryButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: MenuButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: OriginalActionDateTextBox_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member: RefreshButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:      PeoplePlusHours" _
            & vbCrLf & "Member:     PayPeriodComboBox_Change()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:      PayrollAccrual" _
            & vbCrLf & "Member:     GetPayPeriodFilter()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:      PayrollAccrual" _
            & vbCrLf & "Member:     StartDateTextBox_Click()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\ORA.png"
            Case "06B"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\LCARD.png"
            Case "06C"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\MSD.png"
            Case "06D"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\ORC.png"
            Case "06F"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\EJ.png"
            Case "06G"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\WCF.png"
            Case "06H"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\LSASD.png"
            Case "06J"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\ARD.png"
            Case "06K"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\WD.png"
            Case "06L"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\SEMD.png"
            Case "06M"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\ECAD.png"
            Case "06N"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\WSA.png"
            Case "06R"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\MSR.png"
            Case "06X"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\XA.png"
        End Select
    Else
        Me.DivisionIcon.Picture = _
            CurrentProject.path & "\etc\png\AppIcons\interface\ui\Reports.png"
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      FinanceObligations" _
            & vbCrLf & "Member:     SetDivisionIcon()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Dim mFields As Collection
    Dim field As DAO.field
    Dim mData As DAO.Recordset
    Set mData = Me.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim mArray() As String
    j = mData.Fields.count - 1
    ReDim mArray(j)
    For i = LBound(mArray()) To UBound(mArray())
        If Not mData.Fields(i).Name & "" = "" Then
            mArray(i) = mData.Fields(i).Name
        End If
    Next i
    GetColumnNames = mArray()
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:       PayrollAccrual" _
            & vbCrLf & "Member:     GetColumnNames" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    Dim mData As DAO.Recordset
    Set mData = Me.Recordset
    mData.Filter = Me.Filter
    Set GetReportData = mData
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member:     GetReportData" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Dim mFields() As String
    mFields() = GetColumnNames
    Dim mBudgetPath As BudgetPath
    Dim mExcel As Excel.Application
    Dim mAllocations As Excel.Workbook
    Dim mWorksheet As Excel.Worksheet
    Dim mList As Excel.ListObject
    Dim mRange As Excel.Range
    Dim mCell As Object
    Dim mHeader As Excel.Range
    Dim mStart As Excel.Range
    Dim mEnd As Excel.Range
    Dim field As DAO.field
    Dim mData As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Set mBudgetPath = New BudgetPath
    Set mExcel = CreateObject("Excel.Application")
    Set mAllocations = mExcel.Workbooks.Open(mBudgetPath.ReportTemplate)
    mAllocations.Worksheets(2).Visible = False
    Set mWorksheet = mAllocations.Worksheets(1)
    mWorksheet.Name = "Payroll Accruals"
    mWorksheet.Cells.HorizontalAlignment = xlHAlignLeft
    mWorksheet.Cells.Font.Name = "Source Code Pro"
    mWorksheet.Cells.Font.Size = 8
    Set mStart = mWorksheet.Cells(1, 1)
    Set mEnd = mWorksheet.Cells(1, UBound(mFields) - 1)
    Set mHeader = mWorksheet.Range(mStart, mEnd)
    mHeader.HorizontalAlignment = xlHAlignLeft
    mHeader.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    mHeader.Font.Name = "Source Code Pro"
    mHeader.Font.Color = vbBlack
    mHeader.Font.Bold = True
    mHeader.Font.Size = 8
    For i = LBound(mFields) To UBound(mFields)
        mHeader.Cells(i + 1).Value = mFields(i)
    Next i
    Set mData = Me.Recordset
    mData.Filter = Me.Filter
    mWorksheet.Cells(2, 1).CopyFromRecordset mData
    mHeader.Font.Color = vbBlack
    mExcel.WindowState = xlMaximized
    mExcel.Visible = True
ErrorHandler:
    If Err.Number <> 0 Then
        mError = "Source:   PayrollAccrual" _
            & vbCrLf & "Member:     ExcelButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
        Set mExcel = Nothing
        Set mAllocations = Nothing
        Set mWorksheet = Nothing
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

