VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FullTimeAllocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit



Private pRC As AccountArgs
Private m_Dialog As Form_AccountEditor
Private FiscalYearFilter As String
Private BudgetLevelFilter As String
Private FundNameFilter As String
Private ProgramProjectNameFilter As String
Private pAnd As String
Private m_DataSet As DAO.Recordset
Private m_OriginalAmount As Double
Private m_Error As String





'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Open
'   Parameters:  Integer
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set m_Dialog = New Form_AccountEditor
    Set pRC = New AccountArgs
    Me.RecordSource = "SELECT * FROM Allocations" _
        & " WHERE Allocations.BocCode = '17'" _
        & " AND Allocations.BudgetLevel = '7'"
ErrorHandler:
    ProcessError
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
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
    pAnd = " AND "
    SetComboBoxColors
    ResetOptionGroupValues
    ClearComboBoxes
    Set m_DataSet = Me.Recordset
    HideButtons
    BudgetLevelFilter = "[BudgetLevel] = '7'"
    FiscalYearFilter = "[BFY] = '2021'"
    Me.Filter = BudgetLevelFilter & " AND " & FiscalYearFilter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Allocations.FundName" _
        & " FROM Allocations" _
        & " WHERE Allocations.BocCode = '17'" _
        & pAnd & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT Allocations.ProgramProjectName" _
        & " FROM Allocations" _
        & " WHERE Allocations.BocCode = '17'" _
        & pAnd & Me.Filter
    Me.BudgetLevelOptionGroup.DefaultValue = 1
    Me.FiscalYearOptionGroup.DefaultValue = 2
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetComboBoxColors
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.ProgramProjectNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.ProgramProjectNameComboBox.BackColor = RGB(33, 33, 33)
    Me.ProgramProjectNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetButtonVisibility
'   Parameters:  Void
'   Purpose:     Toggles the toolbar button on/off
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.FirstButton.Visible = Not Me.FirstButton.Visible
    Me.PreviousButton.Visible = Not Me.PreviousButton.Visible
    Me.NextButton.Visible = Not Me.NextButton.Visible
    Me.LastButton.Visible = Not Me.LastButton.Visible
    Me.EditButton.Visible = Not Me.EditButton.Visible
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.CalculatorButton.Visible = Not Me.CalculatorButton.Visible
    Me.ExcelButton.Visible = Not Me.ExcelButton.Visible
    Me.SaveButton.Visible = Not Me.SaveButton.Visible
    Me.CompassStatusOfFundsButton.Visible = Not Me.CompassStatusOfFundsButton.Visible
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
    Me.FirstButton.Visible = False
    Me.PreviousButton.Visible = False
    Me.NextButton.Visible = False
    Me.LastButton.Visible = False
    Me.EditButton.Visible = False
    Me.RefreshButton.Visible = False
    Me.CalculatorButton.Visible = False
    Me.ExcelButton.Visible = False
    Me.SaveButton.Visible = False
    Me.CompassStatusOfFundsButton.Visible = False
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearFilters
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilters()
    On Error GoTo ErrorHandler:
    BudgetLevelFilter = vbNullString
    FiscalYearFilter = vbNullString
    FundNameFilter = vbNullString
    ProgramProjectNameFilter = vbNullString
    Me.Filter = vbNullString
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
'   Name:        ClearComboBoxes
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxes()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.Value = vbNullString
    Me.ProgramProjectNameComboBox.Value = vbNullString
    ClearFilters
ErrorHandler:
    ProcessError
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
    pRC.BFY = "2021"
    Me.BudgetLevelOptionGroup.Value = 1
    BudgetLevelFilter = "[BudgetLevel] = '7'"
    pRC.BudgetLevel = "7"
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
    pRC.BFY = vbNullString
    FiscalYearFilter = vbNullString
    Me.Filter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            pRC.BFY = "2020"
            FiscalYearFilter = "[BFY] = '2020'"
        Case 2
            pRC.BFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
    End Select
    Me.Filter = GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Allocations" _
        & " WHERE Allocations.BocCode = '17' AND " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        BudgetLevelOptionGroup_AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BudgetLevelOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    pRC.BudgetLevel = vbNullString
    BudgetLevelFilter = vbNullString
    Me.Filter = vbNullString
    Select Case Me.BudgetLevelOptionGroup.Value
        Case 1
            pRC.BudgetLevel = "7"
            BudgetLevelFilter = "[BudgetLevel] = '7'"
        Case 2
            pRC.BudgetLevel = "8"
            BudgetLevelFilter = "[BudgetLevel] = '8'"
    End Select
    Me.Filter = GetBudgetLevelFilter
    Me.RecordSource = "SELECT * FROM Allocations WHERE Allocations.BocCode = '17' AND " & Me.Filter
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
Private Function GetFiscalYearFilter()
    On Error GoTo ErrorHandler:
    If Not FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BudgetLevelFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BudgetLevelFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BudgetLevelFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BudgetLevelFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    FundNameFilter = vbNullString
    pRC.FundName = vbNullString
    Me.Filter = vbNullString
    pRC.FundName = Me.FundNameComboBox.Value
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    Me.RecordSource = "SELECT * FROM Allocations WHERE Allocations.BocCode = '17' AND " & Me.Filter
    Me.Requery
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT ProgramProjectName FROM Allocations" _
        & " WHERE Allocations.BocCode = '17' AND " & Me.Filter
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
Private Function GetFundNameFilter()
    On Error GoTo ErrorHandler:
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BudgetLevelFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BudgetLevelFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = BudgetLevelFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            Me.Filter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = BudgetLevelFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProgramProjectNameComboBox_Change
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProgramProjectNameComboBox_Change()
    On Error GoTo ErrorHandler:
    pRC.ProgramProjectName = vbNullString
    ProgramProjectNameFilter = vbNullString
    Me.Filter = vbNullString
    pRC.ProgramProjectName = Me.ProgramProjectNameComboBox.Value
    ProgramProjectNameFilter = "[ProgramProjectName] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    Me.Filter = GetProgramProjectNameFilter
    Me.RecordSource = "SELECT * FROM Allocations WHERE Allocations.BocCode = '17' AND " & Me.Filter
    Me.Requery
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
Private Function GetProgramProjectNameFilter()
    On Error GoTo ErrorHandler:
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & BudgetLevelFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & BudgetLevelFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = BudgetLevelFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BudgetLevelFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = BudgetLevelFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BudgetLevelFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetProgramProjectNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBudgetLevelFilter() As String
    On Error GoTo ErrorHandler:
    If Not BudgetLevelFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBudgetLevelFilter = BudgetLevelFilter
    End If
    If Not BudgetLevelFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = BudgetLevelFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BudgetLevelFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = BudgetLevelFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BudgetLevelFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = FiscalYearFilter _
                & pAnd & BudgetLevelFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BudgetLevelFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = BudgetLevelFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BudgetLevelFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = BudgetLevelFilter _
                & pAnd & FundNameFilter
    End If
    If Not BudgetLevelFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = BudgetLevelFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BudgetLevelFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = BudgetLevelFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BudgetLevelFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = BudgetLevelFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If BudgetLevelFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = ProgramProjectNameFilter
    End If
    If BudgetLevelFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = FiscalYearFilter
    End If
    If BudgetLevelFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If BudgetLevelFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBudgetLevelFilter = FiscalYearFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Close
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    ClearFilters
    Me.Filter = vbNullString
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
'----------------------------------------------------------------------------------
Private Sub CompassStatusOfFundsButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "CompassBudgetStatusQuery"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        EditButton_Click
'   Parameters:  Void
'   Purpose:     Opens editor form based on current record data
'---------------------------------------------------------------------------------
Private Sub EditButton_Click()
    On Error GoTo ErrorHandler:
    Dim m_DataSet As DAO.Database
    Set m_DataSet = CurrentDb
    Dim Account As DAO.Recordset
    Dim ID As Integer
    Me.CurrentIdTextBox.SetFocus
    ID = CInt(Me!CurrentIdTextBox.Value)
    Dim frm As Form_AccountEditor
    Set frm = New Form_AccountEditor
    DoCmd.OpenForm FormName:="AccountEditor", OpenArgs:=str(ID)
    Set frm = Forms("AccountEditor")
    Forms("AccountEditor").Tag = Me.Name
    Forms("AccountEditor").RecordSource = Me.RecordSource
    Forms("AccountEditor").BudgetLevelTextBox.SetFocus
    Forms("AccountEditor").BudgetLevelTextBox.Value = Me.BudgetLevelTextBox.Value
    Forms("AccountEditor").RpioCodeTextBox.SetFocus
    Forms("AccountEditor").RpioCodeTextBox.Value = Me.RpioCodeTextBox.Value
    Forms("AccountEditor").BfyTextBox.SetFocus
    Forms("AccountEditor").BfyTextBox.Value = Me.BfyTextBox.Value
    Forms("AccountEditor").AhCodeTextBox.SetFocus
    Forms("AccountEditor").AhCodeTextBox.Value = Me.AhCodeTextBox.Value
    Forms("AccountEditor").FundCodeTextBox.SetFocus
    Forms("AccountEditor").FundCodeTextBox.Value = Me.FundCodeTextBox.Value
    Forms("AccountEditor").OrgCodeTextBox.SetFocus
    Forms("AccountEditor").OrgCodeTextBox.Value = Me.OrgCodeTextBox.Value
    Forms("AccountEditor").AccountCodeTextBox.SetFocus
    Forms("AccountEditor").AccountCodeTextBox.Value = Me.AccountCodeTextBox.Value
    Forms("AccountEditor").BocCodeTextBox.SetFocus
    Forms("AccountEditor").BocCodeTextBox.Value = Me.BocCodeTextBox.Value
    Forms("AccountEditor").RcCodeTextBox.SetFocus
    Forms("AccountEditor").RcCodeTextBox.Value = Me.RcCodeTextBox.Value
    Forms("AccountEditor").OriginalAmountTextBox.SetFocus
    Forms("AccountEditor").OriginalAmountTextBox.Value = Me.AmountTextBox.Value
    DoCmd.Close ObjectType:=acForm, _
        ObjectName:=Me.Name, Save:=acSaveYes
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        CalculatorButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    Calculator.Calculate
ErrorHandler:
    ProcessError
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        MenuButton_Click
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
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxes
    ResetOptionGroupValues
    ClearFilters
    BudgetLevelFilter = "[BudgetLevel] = '7'"
    FiscalYearFilter = "[BFY] = '2021'"
    Me.Filter = BudgetLevelFilter & " AND " & FiscalYearFilter
    Me.RecordSource = "SELECT * FROM Allocations WHERE Allocations.BocCode = '17' AND " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Allocations.FundName" _
        & " FROM Allocations" _
        & " WHERE Allocations.BocCode = '17'" _
        & pAnd & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT Allocations.ProgramProjectName" _
        & " FROM Allocations" _
        & " WHERE Allocations.BocCode = '17'" _
        & pAnd & Me.Filter
    Me.BudgetLevelOptionGroup.Value = 1
    Me.FiscalYearOptionGroup.Value = 2
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        UpdateBudgetLevel
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateBudgetLevel()
    On Error GoTo ErrorHandler:
    Me.PrcIdTextBox.SetFocus
    pRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.BudgetLevelTextBox.SetFocus
    pRC.BudgetLevel = Me.BudgetLevelTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.BudgetLevel = '" _
        & pRC.BudgetLevel & "'" & _
        " WHERE Allocations.PrcId = " & pRC.ID
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
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateBfy()
    On Error GoTo ErrorHandler:
    Me.PrcIdTextBox.SetFocus
    pRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.BfyTextBox.SetFocus
    pRC.BFY = Me.BfyTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.BFY = '" _
        & pRC.BFY & "'" & _
        " WHERE Allocations.PrcId = " & pRC.ID
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
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateRpioCode()
    On Error GoTo ErrorHandler:
    Me.PrcIdTextBox.SetFocus
    pRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.RpioCodeTextBox.SetFocus
    pRC.RpioCode = Me.RpioCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.RPIO = '" _
        & pRC.RpioCode & "'" & _
        " WHERE Allocations.PrcId = " & pRC.ID
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
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateAhCode()
    On Error GoTo ErrorHandler:
    Me.PrcIdTextBox.SetFocus
    pRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.AhCodeTextBox.SetFocus
    pRC.AhCode = Me.AhCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.AhCode = '" _
        & pRC.AhCode & "'" & _
        " WHERE Allocations.PrcId = " & pRC.ID
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
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateOrgCode()
    On Error GoTo ErrorHandler:
    Me.PrcIdTextBox.SetFocus
    pRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.OrgCodeTextBox.SetFocus
    pRC.OrgCode = Me.OrgCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.OrgCode = '" _
        & pRC.OrgCode & "'" & _
        " WHERE Allocations.PrcId = " & pRC.ID
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
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateAccountCode()
    On Error GoTo ErrorHandler:
    Me.PrcIdTextBox.SetFocus
    pRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.AccountCodeTextBox.SetFocus
    pRC.AccountCode = Me.AccountCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.AccountCode = '" _
        & pRC.AccountCode & "'" & _
        " WHERE Allocations.PrcId = " & pRC.ID
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
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateBocCode()
    On Error GoTo ErrorHandler:
    Me.PrcIdTextBox.SetFocus
    pRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.BocCodeTextBox.SetFocus
    pRC.BocCode = Me.BocCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.BocCode = '" _
        & pRC.BocCode & "'" & _
        " WHERE Allocations.PrcId = " & pRC.ID
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
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateRcCode()
    On Error GoTo ErrorHandler:
    Me.PrcIdTextBox.SetFocus
    pRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.RcCodeTextBox.SetFocus
    pRC.RcCode = Me.RcCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.RcCode = '" _
        & pRC.RcCode & "'" & _
        " WHERE Allocations.PrcId = " & pRC.ID
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
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UpdateAmount()
    On Error GoTo ErrorHandler:
    Me.PrcIdTextBox.SetFocus
    pRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.AmountTextBox.SetFocus
    Dim Change As Double
    Change = CDbl(Me.AmountTextBox.text)
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.Amount = " _
        & Change & _
        " WHERE Allocations.PrcId = " & pRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
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
Private Function GetReportData() As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim m_SQL As String
    Dim m_Data As DAO.Recordset
    m_SQL = "SELECT FullTimeEquivalents.BudgetLevel AS [Level], FullTimeEquivalents.RPIO, FullTimeEquivalents.BFY," _
        & " FullTimeEquivalents.AhCode AS AH, FullTimeEquivalents.FundCode AS Fund," _
        & " FullTimeEquivalents.OrgCode AS ORG, FullTimeEquivalents.AccountCode AS Account," _
        & " FullTimeEquivalents.BocCode AS BOC, FullTimeEquivalents.Amount AS Amount" _
        & " FROM FullTimeEquivalents" _
        & " WHERE FullTimeEquivalents.Amount <> 0" & pAnd & Me.Filter
    Set m_Data = CurrentDb.OpenRecordset(m_SQL)
    Set GetReportData = m_Data
ErrorHandler:
    ProcessError
    Exit Function
End Function





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
    Dim i As Integer
    Dim j As Integer
    Dim m_Array() As String
    Set m_Data = GetReportData
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
'   Type:        Event Sub-Procedure
'   Name:        ExcelButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExcelButton_Click()
    On Error GoTo ErrorHandler:
    Dim m_BudgetPath As BudgetPath
    Dim m_Excel As Excel.Application
    Dim m_Allocations As Excel.Workbook
    Dim m_Worksheet As Excel.Worksheet
    Dim m_List As Excel.ListObject
    Dim m_Range As Excel.Range
    Dim m_Cell As Variant
    Dim m_Header As Excel.Range
    Dim m_Start As Excel.Range
    Dim m_End As Excel.Range
    Dim field As DAO.field
    Dim m_Data As DAO.Recordset
    Dim m_Fields() As String
    Set m_Data = GetReportData
    m_Fields() = GetColumnNames
    Dim i As Integer
    Dim j As Integer
    Set m_BudgetPath = New BudgetPath
    Set m_Excel = CreateObject("Excel.Application")
    Set m_Allocations = m_Excel.Workbooks.Open(m_BudgetPath.ReportTemplate)
    m_Allocations.Worksheets(2).Visible = False
    Set m_Worksheet = m_Allocations.Worksheets(1)
    m_Worksheet.Name = "Allocations"
    m_Worksheet.Cells.HorizontalAlignment = xlHAlignLeft
    m_Worksheet.Cells.Font.Name = "Source Code Pro"
    m_Worksheet.Cells.Font.Size = 8
    Set m_Start = m_Worksheet.Cells(1, 1)
    Set m_End = m_Worksheet.Cells(1, UBound(m_Fields) + 1)
    Set m_Header = m_Worksheet.Range(m_Start, m_End)
    m_Header.Interior.Color = RGB(217, 217, 217)
    m_Header.HorizontalAlignment = xlHAlignLeft
    m_Header.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    m_Header.Borders(xlEdgeLeft).Color = RGB(133, 133, 133)
    m_Header.Borders(xlEdgeRight).Color = RGB(133, 133, 133)
    m_Header.Borders(xlEdgeTop).Color = RGB(133, 133, 133)
    m_Header.Font.Name = "Source Code Pro"
    m_Header.Font.Color = vbBlack
    m_Header.Font.Bold = True
    m_Header.Font.Size = 8
    For i = LBound(m_Fields) To UBound(m_Fields)
        j = i
        m_Header.Cells(1, i + 1).Value = m_Fields(j)
    Next i
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





