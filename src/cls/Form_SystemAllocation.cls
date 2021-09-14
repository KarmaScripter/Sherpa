VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SystemAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Args As AccountArgs
Private FiscalYearFilter As String
Private FundNameFilter As String
Private AccountCodeFilter As String
Private BocNameFilter As String
Private pAnd As String
Private m_Error As String

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
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
    Me.BudgetStatusQueryButton.Visible = Not Me.BudgetStatusQueryButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.Value = ""
    Me.AccountCodeComboBox.Value = ""
    Me.BocNameComboBox.Value = ""
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilters()
    On Error GoTo ErrorHandler:
    BocNameFilter = ""
    FiscalYearFilter = ""
    FundNameFilter = ""
    BocNameFilter = ""
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
    On Error GoTo ErrorHandler:
        Me.FiscalYearOptionGroup.Value = ""
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BudgetStatusQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "CompassBudgetStatusQuery"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
    pAnd = " AND "
    Set Args = New AccountArgs
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
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.BFY = ""
    FiscalYearFilter = ""
    Me.Filter = ""
    'Set the PRC parameters
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            Args.BFY = "2020"
            FiscalYearFilter = "[BFY] IN ('2019', '2020')"
        Case 2
            Args.BFY = "2021"
            FiscalYearFilter = "[BFY] IN ('2021', '2022')"
    End Select
    Me.Filter = GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM StatusOfFunds WHERE" _
        & " StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFiscalYearFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetFiscalYearFilter() As String
    On Error GoTo ErrorHandler:
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
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FundNameComboBox_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me.Filter = ""
    FundNameFilter = ""
    Args.FundName = Me!FundNameComboBox.SelText
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
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AccountCodeComboBox_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AccountCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me.Filter = ""
    AccountCodeFilter = ""
    Args.AccountCode = Me!AccountCodeComboBox.Value
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
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetAccountCodeFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetAccountCodeFilter() As String
    On Error GoTo ErrorHandler:
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
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocNameComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me.Filter = ""
    BocNameFilter = ""
    Args.BocName = Me!BocNameComboBox.Value
    BocNameFilter = "[BocName] = '" & Me.BocNameComboBox.Value & "'"
    Me.Filter = GetBocNameFilter
    Me.RecordSource = "SELECT * From StatusOfFunds WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND StatusOfFunds.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter() As String
    On Error GoTo ErrorHandler:
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
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
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
'   Name:        AfterUpdate
'   Parameters:  Void
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
'   Type:        Sub-Procedure
'   Name:        AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ResetOptionGroupValues
    ClearFilters
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




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        EditButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:     Opens editor form based on current record data
'---------------------------------------------------------------------------------
Private Sub EditButton_Click()
    On Error GoTo ErrorHandler:
    Dim mDataSet As DAO.Database
    Set mDataSet = CurrentDb
    Dim Account As DAO.Recordset
    Dim ID As Integer
    Me.AllocationSubForm.Controls("PrcIdTextBoxt").SetFocus
    ID = CInt(Me!PrcIdTextBox.Value)
    Dim frm As Form_AccountEditor
    Set frm = New Form_AccountEditor
    DoCmd.OpenForm FormName:="AccountEditor", OpenArgs:=str(ID)
    Set frm = Forms("AccountEditor")
    Forms("AccountEditor").Tag = Me.Name
    Forms("AccountEditor").RecordSource = Me.RecordSource
    Forms("AccountEditor").BudgetLevelTextBox.SetFocus
    Forms("AccountEditor").BudgetLevelTextBox.Value = Me.AllocationSubForm.Controls("BudgetLevelTextBox").Value
    Forms("AccountEditor").RpioCodeTextBox.SetFocus
    Forms("AccountEditor").RpioCodeTextBox.Value = Me.AllocationSubForm.Controls("RpioCodeTextBox").Value
    Forms("AccountEditor").BfyTextBox.SetFocus
    Forms("AccountEditor").BfyTextBox.Value = Me.AllocationSubForm.Controls("BfyTextBox").Value
    Forms("AccountEditor").AhCodeTextBox.SetFocus
    Forms("AccountEditor").AhCodeTextBox.Value = Me.AllocationSubForm.Controls("AhCodeTextBox").Value
    Forms("AccountEditor").FundCodeTextBox.SetFocus
    Forms("AccountEditor").FundCodeTextBox.Value = Me.AllocationSubForm.Controls("FundCodeTextBox").Value
    Forms("AccountEditor").OrgCodeTextBox.SetFocus
    Forms("AccountEditor").OrgCodeTextBox.Value = Me.AllocationSubForm.Controls("OrgCodeTextBox").Value
    Forms("AccountEditor").AccountCodeTextBox.SetFocus
    Forms("AccountEditor").AccountCodeTextBox.Value = Me.AllocationSubForm.Controls("AccountCodeTextBox").Value
    Forms("AccountEditor").BocCodeTextBox.SetFocus
    Forms("AccountEditor").BocCodeTextBox.Value = Me.AllocationSubForm.Controls("BocCodeTextBox").Value
    Forms("AccountEditor").RcCodeTextBox.SetFocus
    Forms("AccountEditor").RcCodeTextBox.Value = Me.AllocationSubForm.Controls("RcCodeTextBox").Value
    Forms("AccountEditor").OriginalAmountTextBox.SetFocus
    Forms("AccountEditor").OriginalAmountTextBox.Value = Me.AllocationSubForm.Controls("AmountTextBox").Value
    DoCmd.Close ObjectType:=acForm, _
        ObjectName:="SystemAllocation", Save:=acSaveYes
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateBudgetLevel
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateBudgetLevel()
    On Error GoTo ErrorHandler:
    Me.AllocationSubForm.Controls("PrcIdTextBox").SetFocus
    Args.ID = CInt(Me!PrcIdTextBox.text)
    Me.AllocationSubForm.Controls("BudgetLevelTextBox").SetFocus
    Args.BudgetLevel = Me.AllocationSubForm.Controls("BudgetLevelTextBox").text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.BudgetLevel = '" _
        & Args.BudgetLevel & "'" & _
        " WHERE Allocations.PrcId = " & Args.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateBfy
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateBfy()
    On Error GoTo ErrorHandler:
    Me.AllocationSubForm.Controls("PrcIdTextBox").SetFocus
    Args.ID = CInt(Me.AllocationSubForm.Controls("PrcIdTextBox").text)
    Me.AllocationSubForm.Controls("BfyTextBox").SetFocus
    Args.BFY = Me.AllocationSubForm.Controls("BfyTextBox").text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.BFY = '" _
        & Args.BFY & "'" & _
        " WHERE Allocations.PrcId = " & Args.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateRpioCode
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateRpioCode()
    On Error GoTo ErrorHandler:
    Me.AllocationSubForm.Controls("PrcIdTextBox").SetFocus
    Args.ID = CInt(Me.AllocationSubForm.Controls("PrcIdTextBox").text)
    Me.AllocationSubForm.Controls("RpioCodeTextBox").SetFocus
    Args.RpioCode = Me.AllocationSubForm.Controls("RpioCodeTextBox").text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.RPIO = '" _
        & Args.RpioCode & "'" & _
        " WHERE Allocations.PrcId = " & Args.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateAhCode
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateAhCode()
    On Error GoTo ErrorHandler:
    Me.AllocationSubForm.Controls("PrcIdTextBox").SetFocus
    Args.ID = CInt(Me.AllocationSubForm.Controls("PrcIdTextBox").text)
    Me.AllocationSubForm.Controls("AhCodeTextBox").SetFocus
    Args.AhCode = Me.AllocationSubForm.Controls("AhCodeTextBox").text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.AhCode = '" _
        & Args.AhCode & "'" & _
        " WHERE Allocations.PrcId = " & Args.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateOrgCode
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateOrgCode()
    On Error GoTo ErrorHandler:
    Me.AllocationSubForm.Controls("PrcIdTextBox").SetFocus
    Args.ID = CInt(Me.AllocationSubForm.Controls("PrcIdTextBox").text)
    Me.AllocationSubForm.Controls("OrgCodeTextBox").SetFocus
    Args.OrgCode = Me.AllocationSubForm.Controls("OrgCodeTextBox").text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.OrgCode = '" _
        & Args.OrgCode & "'" & _
        " WHERE Allocations.PrcId = " & Args.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateAccountCode
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateAccountCode()
    On Error GoTo ErrorHandler:
    Me.AllocationSubForm.Controls("PrcIdTextBox").SetFocus
    Args.ID = CInt(Me.AllocationSubForm.Controls("PrcIdTextBox").text)
    Me.AllocationSubForm.Controls("AccountCodeTextBox").SetFocus
    Args.AccountCode = Me.AllocationSubForm.Controls("AccountCodeTextBox").text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.AccountCode = '" _
        & Args.AccountCode & "'" & _
        " WHERE Allocations.PrcId = " & Args.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateBocCode
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateBocCode()
    On Error GoTo ErrorHandler:
    Me.AllocationSubForm.Controls("PrcIdTextBox").SetFocus
    Args.ID = CInt(Me.AllocationSubForm.Controls("PrcIdTextBox").text)
    Me.AllocationSubForm.Controls("BocCodeTextBox").SetFocus
    Args.BocCode = Me.AllocationSubForm.Controls("BocCodeTextBox").text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.BocCode = '" _
        & Args.BocCode & "'" & _
        " WHERE Allocations.PrcId = " & Args.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateRcCode
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateRcCode()
    On Error GoTo ErrorHandler:
    Me.AllocationSubForm.Controls("PrcIdTextBox").SetFocus
    Args.ID = CInt(Me.AllocationSubForm.Controls("PrcIdTextBox").text)
    Me.AllocationSubForm.Controls("RcCodeTextBox").SetFocus
    Args.RcCode = Me.AllocationSubForm.Controls("RcCodeTextBox").text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.RcCode = '" _
        & Args.RcCode & "'" & _
        " WHERE Allocations.PrcId = " & Args.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateAmount
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateAmount()
    On Error GoTo ErrorHandler:
    Me.AllocationSubForm.Controls("PrcIdTextBox").SetFocus
    Args.ID = CInt(Me.AllocationSubForm.Controls("PrcIdTextBox").text)
    Me.AllocationSubForm.Controls("AmountTextBox").SetFocus
    Dim Change As Double
    Change = CDbl(Me.AllocationSubForm.Controls("AmountTextBox").text)
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.Amount = " _
        & Change & _
        " WHERE Allocations.PrcId = " & Args.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
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





