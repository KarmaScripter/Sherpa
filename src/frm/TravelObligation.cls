Option Compare Database
Option Explicit



Public mDialog As Form_DivisionSelector
Public Args As AccountArgs
Private LastNameFilter As String
Private AccountCodeFilter As String
Private BocNameFilter As String
Private FocNameFilter As String
Private DivisionNameFilter As String
Private FiscalYearFilter As String
Private pAnd As String
Private mError As String
Private mNotification As String




'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    pAnd = " AND "
    Set mDialog = New Form_DivisionSelector
    Set Args = New AccountArgs
    DoCmd.OpenForm FormName:="DivisionSelector", WindowMode:=acDialog
    Set mDialog = Forms("DivisionSelector")
    Set Args = mDialog.Args
    Me.FiscalYearOptionGroup.Value = 2
    If Not Args.DivisionName & "" = "" Then
        DivisionNameFilter = "[RcName] = '" & Args.DivisionName & "'"
        FiscalYearFilter = "[BFY] = '2021'"
        Me.Filter = DivisionNameFilter
        Me.RecordSource = "SELECT * FROM TravelObligations" _
            & " WHERE " & Me.Filter
        Me.Requery
        SetDivisionIcon
    Else
        Me.RecordSource = "SELECT * FROM TravelObligations" _
            & " WHERE " & Me.Filter
        Me.Requery
    End If
    DoCmd.Close ObjectType:=acForm, ObjectName:="DivisionSelector", Save:=acSaveNo
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
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
    SetButtonVisibility
    ResetOptionGroupValues
    ClearComboBoxValues
    SetComboBoxColors
    ClearFilterValues
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT TravelObligations.AccountCode" _
        & " FROM TravelObligations" _
        & " WHERE " & Me.Filter
    Me.FocNameComboBox.RowSource = "SELECT DISTINCT TravelObligations.FocName" _
        & " FROM TravelObligations" _
        & " WHERE " & Me.Filter
    Me.LastNameComboBox.RowSource = "SELECT DISTINCT TravelObligations.LastName" _
        & " FROM TravelObligations" _
        & " WHERE " & Me.Filter
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member:     Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
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
        mError = "Source:   TravelObligations" _
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
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member:     Form_Timer()" _
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
'   Retval:      Void
'   Purpose:
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
        mError = "Source:   TravelObligations" _
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
    Me.LastNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.LastNameComboBox.BackColor = RGB(33, 33, 33)
    Me.LastNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.FocNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FocNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FocNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.AccountCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me.AccountCodeComboBox.BackColor = RGB(33, 33, 33)
    Me.AccountCodeComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
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
    Me.LastNameComboBox.Value = vbNullString
    Me.FocNameComboBox.Value = vbNullString
    Me.AccountCodeComboBox.Value = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
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
    LastNameFilter = vbNullString
    FocNameFilter = vbNullString
    AccountCodeFilter = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member: ClearFilterValues()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ResetOptionGroupValues
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
    On Error GoTo ErrorHandler:
     Me.FiscalYearOptionGroup.Value = 2
     FiscalYearFilter = "[BFY] = '2021'"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member: ResetOptionGroupValues()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        FocNameComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    FocNameFilter = vbNullString
    Me.Filter = vbNullString
    FocNameFilter = "[FocName] = '" & Me.FocNameComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetFocNameFilter
    Me.RecordSource = "SELECT * FROM TravelObligations WHERE " & Me.Filter
    Me.Requery
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT TravelObligations.AccountCode" _
        & " FROM TravelObligations" _
        & " WHERE " & Me.Filter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member:     FocNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetFocNameFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetFocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not FocNameFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        LastNameFilter = vbNullString Then
            GetFocNameFilter = FocNameFilter
    End If
    If Not FocNameFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        LastNameFilter = vbNullString Then
            GetFocNameFilter = FocNameFilter _
            & pAnd & AccountCodeFilter
    End If
    If Not FocNameFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        LastNameFilter = vbNullString Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FocNameFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & LastNameFilter
    End If
    If Not FocNameFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & LastNameFilter
    End If
    If Not FocNameFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & LastNameFilter
    End If
    If Not FocNameFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        LastNameFilter = vbNullString Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FocNameFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & LastNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member: GetFocNameFilter()" _
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
    Me.RecordSource = "SELECT * FROM TravelObligations WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member:     AccountCodeComboBox_Change()" _
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
        FocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        LastNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not FocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        LastNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
            & pAnd & FocNameFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not FocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        LastNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not FocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FocNameFilter _
                & pAnd & LastNameFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        FocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & LastNameFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        FocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & LastNameFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        FocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        LastNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not FocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FocNameFilter _
                & pAnd & LastNameFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        FocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & LastNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member: GetAccountCodeFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        LastNameComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub LastNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    LastNameFilter = vbNullString
    LastNameFilter = "[LastName] = '" & Me.LastNameComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetLastNameFilter
    Me.FocNameComboBox.RowSource = "SELECT DISTINCT TravelObligations.FocName" _
        & " FROM TravelObligations" _
        & " WHERE " & Me.Filter
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT TravelObligations.AccountCode" _
        & " FROM TravelObligations" _
        & " WHERE " & Me.Filter
    Me.RecordSource = "SELECT * FROM TravelObligations" _
        & " WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member:     LastNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetLastNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetLastNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not LastNameFilter = vbNullString And _
        FocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter
    End If
    If Not LastNameFilter = vbNullString And _
        Not FocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter _
            & pAnd & FocNameFilter
    End If
    If Not LastNameFilter = vbNullString And _
        Not FocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & FocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not LastNameFilter = vbNullString And _
        Not FocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & FocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not LastNameFilter = vbNullString And _
        FocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not LastNameFilter = vbNullString And _
        FocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not LastNameFilter = vbNullString And _
        FocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not LastNameFilter = vbNullString And _
        Not FocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & FocNameFilter _
                & pAnd & AccountCodeFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member:     GetLastNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        FiscalYearOptionGroup_AfterUpdate
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    FiscalYearFilter = vbNullString
    Select Case FiscalYearOptionGroup.Value
        Case 1
            FiscalYearFilter = "[BFY] = '2020'"
            Me.Filter = DivisionNameFilter & pAnd & GetFiscalYearFilter
        Case 2
            FiscalYearFilter = "[BFY] = '2021'"
            Me.Filter = DivisionNameFilter & pAnd & GetFiscalYearFilter
    End Select
    Me.RecordSource = "SELECT * FROM TravelObligations WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member: FiscalYearOptionGroup_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetFiscalYearFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetFiscalYearFilter() As String
    On Error GoTo ErrorHandler:
    If Not FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        LastNameFilter = vbNullString And _
        FocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        LastNameFilter = vbNullString And _
        FocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        Not LastNameFilter = vbNullString And _
        FocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & LastNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        Not LastNameFilter = vbNullString And _
        FocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
            & pAnd & LastNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        Not LastNameFilter = vbNullString And _
        Not FocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & LastNameFilter _
                & pAnd & FocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        Not LastNameFilter = vbNullString And _
        Not FocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & LastNameFilter _
                & pAnd & FocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString And _
        LastNameFilter = vbNullString And _
        Not FocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString And _
        LastNameFilter = vbNullString And _
        Not FocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FocNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
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
        mError = "Source:   TravelObligations" _
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
        mError = "Source:   TravelObligations" _
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
        mError = "Source:   TravelObligations" _
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
Private Sub OriginalActionDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.OriginalActionDateTextBox
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
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
    ResetOptionGroupValues
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT TravelObligations.AccountCode" _
        & " FROM TravelObligations" _
        & " WHERE " & DivisionNameFilter & pAnd & FiscalYearFilter & ";"
    Me.FocNameComboBox.RowSource = "SELECT DISTINCT TravelObligations.FocName" _
        & " FROM TravelObligations" _
        & " WHERE " & DivisionNameFilter & pAnd & FiscalYearFilter & ";"
    Me.LastNameComboBox.RowSource = "SELECT DISTINCT TravelObligations.LastName" _
        & " FROM TravelObligations" _
        & " WHERE " & DivisionNameFilter & pAnd & FiscalYearFilter & ";"
    Me.RecordSource = "SELECT * FROM TravelObligations" _
        & " WHERE " & DivisionNameFilter & pAnd & FiscalYearFilter & ";"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member: RefreshButton_Click()" _
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
Private Sub ReturnDateTextBox_Enter()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.OriginalActionDateTextBox
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member: ReturnDateTextBox_Enter()" _
            & vbCrLf & "Descript:   " & Err.Description
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
    On Error GoTo ErrorHandler:
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
        mError = "Source:   TravelObligations" _
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
        mError = "Source:       TravelObligations" _
            & vbCrLf & "Member:     GetColumnNames" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetColumnNames
'   Parameters:  Void
'   Retval:      Collection
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetReportData() As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim mSQL As String
    Dim mData As DAO.Recordset
    mSQL = ""
    Set mData = CurrentDb.OpenRecordset(mSQL)
    Set GetReportData = mData
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:       TravelObligations" _
            & vbCrLf & "Member:     GetReportData()" _
            & vbCrLf & "Descript:   " & Err.Description
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
    mWorksheet.Name = "Travel Obligations"
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
        mError = "Source:   TravelObligations" _
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
        mError = "Source:   TravelObligations" _
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
        mError = "Source:   TravelObligations" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



