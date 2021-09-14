Option Compare Database
Option Explicit


Private PRC As AccountArgs
Private FiscalYearFilter As String
Private FundNameFilter As String
Private AccountCodeFilter As String
Private pAnd As String



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
    Me.FirstButton.Visible = Not Me.FirstButton.Visible
    Me.PreviousButton.Visible = Not Me.PreviousButton.Visible
    Me.NextButton.Visible = Not Me.NextButton.Visible
    Me.LastButton.Visible = Not Me.LastButton.Visible
    Me.EditButton.Visible = Not Me.EditButton.Visible
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.CalculatorButton.Visible = Not Me.CalculatorButton.Visible
    Me.SaveButton.Visible = Not Me.SaveButton.Visible
    Me.BudgetStatusQueryButton.Visible = Not Me.BudgetStatusQueryButton.Visible
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    Me.FundNameComboBox.Value = vbNullString
    Me.ProgramProjectNameComboBox.Value = vbNullString
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilters()
    FiscalYearFilter = vbNullString
    FundNameFilter = vbNullString
    Me.Filter = vbNullString
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
        Me.FiscalYearOptionGroup.Value = vbNullString
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BudgetStatusQueryButton_Click()
    DoCmd.OpenForm "CompassBudgetStatusQuery"
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    SetButtonVisibility
    Set PRC = New AccountArgs
    ClearComboBoxValues
    Me.FiscalYearOptionGroup.Value = 2
    FiscalYearFilter = "[BFY] = '2021'"
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode IN ('21', '28');"
        
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.FundName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode IN ('21', '28');"
    
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.AccountCode" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode IN ('21', '28');"
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    PRC.BFY = vbNullString
    FiscalYearFilter = vbNullString
    Me.Filter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            PRC.BFY = "2020"
            FiscalYearFilter = "[BFY] = '2020'"
        Case 2
            PRC.BFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
    End Select
    Me.Filter = GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM StatusOfFunds WHERE" _
        & " StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode IN ('21', '28')" _
        & pAnd & Me.Filter
    Me.Requery
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFiscalYearFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetFiscalYearFilter() As String
    If Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
End Function







'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    pAnd = " AND "
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & " StatusOfFunds.BocCode IN ('21', '28');"
        
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.FundName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & " StatusOfFunds.BocCode IN ('21', '28');"
    
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.AccountCode" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & " StatusOfFunds.BocCode IN ('21', '28');"
        
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FundNameComboBox_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_AfterUpdate()
    Me.Filter = vbNullString
    FundNameFilter = vbNullString
    PRC.FundName = Me.FundNameComboBox.Value
    FundNameFilter = "[FundName] = '" & Me!FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.AccountCode" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode IN ('21', '28')" _
        & pAnd & Me.Filter
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode IN ('21', '28')" _
        & pAnd & Me.Filter
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFundNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetFundNameFilter() As String
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not AccountCodeFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProgramProjectNameComboBox_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProgramProjectNameComboBox_AfterUpdate()
    Me.Filter = vbNullString
    AccountCodeFilter = vbNullString
    PRC.AccountCode = Me!ProgramProjectNameComboBox.Value
    AccountCodeFilter = "[AccountCode] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    Me.Filter = GetAccountCodeFilter
    Me.RecordSource = "SELECT * From StatusOfFunds WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND StatusOfFunds.BocCode IN ('21', '28')" _
        & pAnd & Me.Filter
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetAccountCodeFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetAccountCodeFilter() As String
    If Not AccountCodeFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetAccountCodeFilter = FundNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetAccountCodeFilter = FundNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetAccountCodeFilter = FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not AccountCodeFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetAccountCodeFilter = FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
End Function




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    Calculator.Calculate
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    SetButtonVisibility
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    ClearComboBoxValues
    ResetOptionGroupValues
    ClearFilters
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode IN ('21', '28');"
    Me.FiscalYearOptionGroup.Value = 2
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
        mError = "Source:   TravelReconcilliation" _
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
        mError = "Source:   TravelReconcilliation" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


