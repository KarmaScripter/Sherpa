Option Explicit

Private PRC As AccountArgs
Private FiscalYearFilter As String
Private FundNameFilter As String
Private AccountCodeFilter As String
Private BocNameFilter As String
Private pAnd As String

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
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
    Me.BudgetStatusQueryButton.Visible = Not Me.BudgetStatusQueryButton.Visible
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    Me.FundNameComboBox.Value = ""
    Me.AccountCodeComboBox.Value = ""
    Me.BocNameComboBox.Value = ""
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilters()
    BocNameFilter = ""
    FiscalYearFilter = ""
    FundNameFilter = ""
    BocNameFilter = ""
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
        Me.FiscalYearOptionGroup.Value = ""
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
Private Sub Form_Load()
    SetButtonVisibility
    pAnd = " AND "
    Set PRC = New AccountArgs
    ClearComboBoxValues
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode NOT IN ('10', '17');"
        
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.FundName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode NOT IN ('10', '17');"
    
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.AccountCode" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode NOT IN ('10', '17');"
        
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.BocName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode NOT IN ('10', '17');"
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    PRC.BFY = ""
    FiscalYearFilter = ""
    Me.Filter = ""
    'Set the PRC parameters
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            PRC.BFY = "2020"
            FiscalYearFilter = "[BFY] IN ('2019', '2020')"
        Case 2
            PRC.BFY = "2021"
            FiscalYearFilter = "[BFY] IN ('2021', '2022')"
    End Select
    Me.Filter = GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM StatusOfFunds WHERE" _
        & " StatusOfFunds.BudgetLevel = '7'" _
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
    If Not FiscalYearFilter = "" And _
        BocNameFilter = "" And _
        FundNameFilter = "" And _
        AccountCodeFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = "" And _
        Not BocNameFilter = "" And _
        FundNameFilter = "" And _
        AccountCodeFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter = "" And _
        Not BocNameFilter = "" And _
        Not FundNameFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FiscalYearFilter = "" And _
        BocNameFilter = "" And _
        Not FundNameFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FiscalYearFilter = "" And _
        Not BocNameFilter = "" And _
        FundNameFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FiscalYearFilter = "" And _
        BocNameFilter = "" And _
        Not FundNameFilter = "" And _
        AccountCodeFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = "" And _
        BocNameFilter = "" And _
        FundNameFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FundNameComboBox_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_AfterUpdate()
    Me.Filter = ""
    FundNameFilter = ""
    PRC.FundName = Me!FundNameComboBox.SelText
    FundNameFilter = "[FundName] = '" & Me!FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.AccountCode" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter
        
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode NOT IN ('10', '17')" _
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
    If Not FundNameFilter = "" And _
        BocNameFilter = "" And _
        FiscalYearFilter = "" And _
        AccountCodeFilter = "" Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = "" And _
        Not BocNameFilter = "" And _
        FiscalYearFilter = "" And _
        AccountCodeFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not FundNameFilter = "" And _
        Not BocNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & BocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter = "" And _
        BocNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter = "" And _
        Not BocNameFilter = "" And _
        FiscalYearFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & BocNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter = "" And _
        BocNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        AccountCodeFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = "" And _
        BocNameFilter = "" And _
        FiscalYearFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AccountCodeComboBox_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AccountCodeComboBox_AfterUpdate()
    Me.Filter = ""
    AccountCodeFilter = ""
    PRC.AccountCode = Me!AccountCodeComboBox.Value
    AccountCodeFilter = "[AccountCode] = '" & Me.AccountCodeComboBox.Value & "'"
    Me.Filter = GetAccountCodeFilter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.BocName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & "StatusOfFunds.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter
        
    Me.RecordSource = "SELECT * From StatusOfFunds WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND StatusOfFunds.BocCode NOT IN ('10', '17')" _
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
    If Not AccountCodeFilter = "" And _
        BocNameFilter = "" And _
        FiscalYearFilter = "" And _
        AccountCodeFilter = "" Then
            GetAccountCodeFilter = AccountCodeFilter
    End If
    If Not AccountCodeFilter = "" And _
        Not BocNameFilter = "" And _
        FiscalYearFilter = "" And _
        AccountCodeFilter = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & BocNameFilter
    End If
    If Not AccountCodeFilter = "" And _
        Not BocNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        Not FundNameFilter = "" Then
            GetAccountCodeFilter = FundNameFilter _
                & pAnd & BudgetLevelFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not AccountCodeFilter = "" And _
        BocNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetAccountCodeFilter = FundNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not AccountCodeFilter = "" And _
        Not BocNameFilter = "" And _
        FiscalYearFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetAccountCodeFilter = FundNameFilter _
                & pAnd & BudgetLevelFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not AccountCodeFilter = "" And _
        BocNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        FundNameFilter = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not AccountCodeFilter = "" And _
        BocNameFilter = "" And _
        FiscalYearFilter = "" And _
        Not FundNameFilter = "" Then
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
Private Sub BocNameComboBox_AfterUpdate()
    Me.Filter = ""
    BocNameFilter = ""
    PRC.BocName = Me!BocNameComboBox.Value
    BocNameFilter = "[BocName] = '" & Me.BocNameComboBox.Value & "'"
    Me.Filter = GetBocNameFilter
    Me.RecordSource = "SELECT * From StatusOfFunds WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND StatusOfFunds.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter() As String
    If Not BocNameFilter = "" And _
        FundNameFilter = "" And _
        FiscalYearFilter = "" And _
        AccountCodeFilter = "" Then
            GetBocNameFilter = BocNameFilter
    End If
    If Not BocNameFilter = "" And _
        Not AccountCodeFilter = "" And _
        FiscalYearFilter = "" And _
        FundNameFilter = "" Then
            GetBocNameFilter = AccountCodeFilter _
                & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = "" And _
        Not AccountCodeFilter = "" And _
        Not FiscalYearFilter = "" And _
        Not FundNameFilter = "" Then
            GetBocNameFilter = FundNameFilter _
                & pAnd & BocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not BocNameFilter = "" And _
        FundNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not BocNameFilter = "" And _
        Not FundNameFilter = "" And _
        FiscalYearFilter = "" And _
        Not AccountCodeFilter = "" Then
            GetBocNameFilter = FundNameFilter _
                & pAnd & BocNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not BocNameFilter = "" And _
        AccountCodeFilter = "" And _
        Not FiscalYearFilter = "" And _
        FundNameFilter = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter = "" And _
        AccountCodeFilter = "" And _
        FiscalYearFilter = "" And _
        Not FundNameFilter = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
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
        mError = "Source:   SystemAllocation" _
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
        mError = "Source:   SystemAllocation" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



