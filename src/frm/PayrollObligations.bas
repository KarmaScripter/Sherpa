Option Compare Database
Option Explicit



Public Args As PayrollArgs
Public mDialog As Form_PayrollDialog
Private FundNameFilter As String
Private DivisionNameFilter As String
Private PayPeriodFilter As String
Private WorkCodeFilter As String
Private ProgramProjectNameFilter As String
Private StartDateFilter As String
Private EndDateFilter As String
Public pAnd As String
Public calendar As Form_BudgetCalendar
Private mError As String
Private mNotification As String





'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        Form_Open
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler::
    Set mDialog = New Form_PayrollDialog
    DoCmd.OpenForm FormName:="PayrollDialog", _
        WindowMode:=acDialog
    Set mDialog = Forms("PayrollDialog")
    pAnd = " AND "
    Set Args = New PayrollArgs
    Set Args = mDialog.Args
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
    DoCmd.Close ObjectType:=acForm, ObjectName:=mDialog.Name, Save:=acSaveNo
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member:     Form_Open()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member: FundNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member: GetFundNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        Dim msg As String
        msg = "Source:  PayrollObligations" _
            & vbCrLf & vbCrLf & "Description:  " & Err.Description
            Err.Clear
        MessageFactory.ShowError (msg)
    End If
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member: ProgramProjectNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member:     GetProgramProjectFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member: PayPeriodComboBox_Change()" _
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member: GetPayPeriodFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member: Form_Timer()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   GetWorkCodeFilter" _
            & vbCrLf & "Member: Form_Timer()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member: Form_Timer()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member: Form_Timer()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   PayrollObligations" _
            & vbCrLf & "Member: Form_Timer()" _
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
        mError = "Source:   PayrollObligations" _
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
    Me.PayrollQueryButton.Visible = False
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      PayrollObligations" _
            & vbCrLf & "Member:     HideButtons()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:      PayrollObligations" _
            & vbCrLf & "Member:     InitButtonVisibility()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:      PayrollObligations" _
            & vbCrLf & "Member:     ClearFilterValues()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:      PayrollObligations" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:      PayrollObligations" _
            & vbCrLf & "Member:     CalculatorButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:      PayrollObligations" _
            & vbCrLf & "Member:     MenuButton_Click()" _
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
    On Error GoTo ErrorHandler:::
    If Not Args.RcCode = vbNullString Then
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
        mError = "Source:      PayrollObligations" _
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
        mError = "Source:       PayrollObligations" _
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
        mError = "Source:   PayrollObligations" _
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
    mWorksheet.Name = "Payroll Obligations"
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
        mError = "Source:   PayrollObligations" _
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



