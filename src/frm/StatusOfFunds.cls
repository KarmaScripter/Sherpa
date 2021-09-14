Option Compare Database
Option Explicit

Public PRC As AccountArgs
Private FiscalYearFilter As String
Private FundNameFilter As String
Private ProgramProjectNameFilter As String
Private BocNameFilter As String
Private pAnd As String
Private mError As String
Private mNotification As String



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Open
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Me.RecordSource = "SELECT *" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND StatusOfFunds.BocCode NOT IN ('10', '17', '21', '28');"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     Form_Open(Cancel As Integer)" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    pAnd = " AND "
    SetButtonVisibility
    SetComboBoxColors
    ResetOptionGroupValues
    ClearComboBoxes
    Set PRC = New AccountArgs
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.FundName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND BocCode NOT IN ('10', '17', '21', '28');"
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramProjectName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND BocCode NOT IN ('10', '17', '21', '28');"
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.BocName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND BocCode NOT IN ('10', '17', '21', '28');"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     Form_Load()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
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
    Me.CompassStatusOfFundsButton.Visible = Not Me.CompassStatusOfFundsButton.Visible
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
    Me.BocNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.BocNameComboBox.BackColor = RGB(33, 33, 33)
    Me.BocNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
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
    Me.RecordSource = "SELECT * FROM StatusOfFunds WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   StatusOfFunds" _
            & vbCrLf & "Member:      FiscalYearOptionGroup_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    BocNameFilter = vbNullString
    FiscalYearFilter = vbNullString
    FundNameFilter = vbNullString
    ProgramProjectNameFilter = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     ClearFilters()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    Me.BocNameComboBox.Value = vbNullString
    ClearFilters
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     ClearComboBoxes()" _
            & vbCrLf & "Descript: " & Err.Description
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
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     ResetOptionGroupValues()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     GetFiscalYearFilter()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    Me.Filter = vbNullString
    PRC.FundName = Me.FundNameComboBox.Value
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramProjectName FROM StatusOfFunds WHERE" _
        & " StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.BocName FROM StatusOfFunds WHERE" _
        & " StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
    Me.RecordSource = "SELECT * FROM StatusOfFunds WHERE StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     FundNameComboBox_Change()" _
            & vbCrLf & "Descript: " & Err.Description
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
Private Function GetFundNameFilter()
    On Error GoTo ErrorHandler:
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            Me.Filter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     GetFundNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ProgramProjectName = vbNullString
    ProgramProjectNameFilter = vbNullString
    Me.Filter = vbNullString
    PRC.ProgramProjectName = Me.ProgramProjectNameComboBox.Value
    ProgramProjectNameFilter = "[ProgramProjectName] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    Me.Filter = GetProgramProjectNameFilter
    Me.RecordSource = "SELECT * FROM StatusOfFunds WHERE StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
    Me.Requery
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.BocName FROM StatusOfFunds WHERE" _
        & " StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     ProgramProjectNameComboBox_Change()" _
            & vbCrLf & "Descript: " & Err.Description
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
Private Function GetProgramProjectNameFilter()
    On Error GoTo ErrorHandler:
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     GetProgramProjectNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        BocNameComboBox_Change
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    PRC.BocName = vbNullString
    BocNameFilter = vbNullString
    Me.Filter = vbNullString
    PRC.BocName = Me.BocNameComboBox.Value
    BocNameFilter = "[BocName] = '" & Me.BocNameComboBox.Value & "'"
    Me.Filter = GetBocNameFilter
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     BocNameComboBox_Change()" _
            & vbCrLf & "Descript: " & Err.Description
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
Private Function GetBocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     GetBocNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


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
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND BocCode NOT IN ('10', '17', '21', '28');"
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        CalculatorButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     CalculatorButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
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
'----------------------------------------------------------------------------------
Private Sub CompassStatusOfFundsButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "CompassBudgetStatusQuery"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      StatusOfFunds" _
            & vbCrLf & "Member:     CompassStatusOfFundsButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


