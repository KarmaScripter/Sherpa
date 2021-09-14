
Option Compare Database
Option Explicit

Private PRC As AccountArgs
Private FiscalYearFilter As String
Private BocNameFilter As String
Private FundNameFilter As String
Private ProgramProjectNameFilter As String
Private pAnd As String
Private mDataSet As DAO.Recordset
Private mOriginalAmount As Double
Private mError As String
Private mNotification As String


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Me.RecordSource = "SELECT * FROM OperatingPlans;"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
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
    Set PRC = New AccountArgs
    Set mDataSet = Me.Recordset
    HideButtons
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT OperatingPlans.FundName" _
        & " FROM OperatingPlans;"
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT OperatingPlans.ProgramProjectName" _
        & " FROM OperatingPlans;"
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT OperatingPlans.BocName" _
        & " FROM OperatingPlans;"
    Me.Filter = vbNullString
    Me.FilterOn = True
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     Form_Load()" _
            & vbCrLf & "Descript: " & Err.Description
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
        mError = "Source:   OperatingPlans" _
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
    Me.CompassStatusOfFundsButton.Visible = False
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     HideButtons()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetButtonVisibility
'   Parameters:  Void
'   Retval:      Void
'   Purpose:     Toggles the toolbar button on/off
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
    Me.CompassStatusOfFundsButton.Visible = Not Me.CompassStatusOfFundsButton.Visible
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     SetButtonVisibility()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
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
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     ClearFilters()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
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
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     ClearComboBoxes()" _
            & vbCrLf & "Descript: " & Err.Description
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
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     ResetOptionGroupValues()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.BFY = vbNullString
    FiscalYearFilter = vbNullString
    Me.Filter = vbNullString
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
    Me.RecordSource = "SELECT * FROM OperatingPlans WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     FiscalYearOptionGroup_AfterUpdate()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
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
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     GetFiscalYearFilter()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
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
    Me.RecordSource = "SELECT * FROM OperatingPlans WHERE " & Me.Filter
    Me.Requery
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT OperatingPlans.ProgramProjectName" _
        & " FROM OperatingPlans WHERE" _
        & " OperatingPlans.BudgetLevel = '7' AND " & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT OperatingPlans.BocName" _
        & " FROM OperatingPlans" _
        & " WHERE" & Me.Filter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     FundNameComboBox_Change()" _
            & vbCrLf & "Descript: " & Err.Description
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
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     GetFundNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
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
    Me.RecordSource = "SELECT * FROM OperatingPlans WHERE " & Me.Filter
    Me.Requery
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT OperatingPlans.BocName FROM OperatingPlans" _
        & " WHERE" & Me.Filter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     ProgramProjectNameComboBox_Change()" _
            & vbCrLf & "Descript: " & Err.Description
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
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     GetProgramProjectNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        BocNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
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
    Me.RecordSource = "SELECT * FROM OperatingPlans WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     BocNameComboBox_Change()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetBocNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter()
    On Error GoTo ErrorHandler:
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBocNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBocNameFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetBocNameFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetBocNameFilter = ProgramProjectNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetBocNameFilter = FundNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBocNameFilter = ProgramProjectNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetBocNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & BocNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     GetBocNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function

'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    ClearFilters
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:    Form_Close()" _
            & vbCrLf & "Descript: " & Err.Description
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
'----------------------------------------------------------------------------------
Private Sub CompassStatusOfFundsButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetStatusQuery"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     CompassStatusOfFundsButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Dim mDataSet As DAO.Database
    Set mDataSet = CurrentDb
    Dim Account As DAO.Recordset
    Dim ID As Integer
    Me.CurrentIdTextBox.SetFocus
    ID = CInt(Me!CurrentIdTextBox.text)
    Me.RecordSource = "SELECT * FROM OperatingPlans WHERE OppId = " & ID
    Dim frm As Form
    DoCmd.OpenForm ("AccountEditor")
    Form_AccountEditor.Tag = Me.Name
    Form_AccountEditor.RecordSource = Me.RecordSource
    Form_AccountEditor.IdTextBox.Value = ID
    Form_AccountEditor.BudgetLevelTextBox.Value = Me.BudgetLevelTextBox.Value
    Form_AccountEditor.RpioCodeTextBox.Value = Me.RpioCodeTextBox.Value
    Form_AccountEditor.BfyTextBox.Value = Me.BfyTextBox.Value
    Form_AccountEditor.AhCodeTextBox.Value = Me.AhCodeTextBox.Value
    Form_AccountEditor.FundCodeTextBox.Value = Me.FundCodeTextBox.Value
    Form_AccountEditor.OrgCodeTextBox.Value = Me.OrgCodeTextBox.Value
    Form_AccountEditor.AccountCodeTextBox.Value = Me.AccountCodeTextBox.Value
    Form_AccountEditor.BocCodeTextBox.Value = Me.BocCodeTextBox.Value
    Form_AccountEditor.RcCodeTextBox.Value = Me.RcCodeTextBox.Value
    Form_AccountEditor.OriginalAmountTextBox.Value = Me.AmountTextBox.Value
    DoCmd.Close acForm, "OperatingPlans", acSaveNo
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     EditButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
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
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     CalculatorButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
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
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     MenuButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Me.Filter = vbNullString
    ClearFilters
    Me.RecordSource = "SELECT * FROM OperatingPlans;"
    Set PRC = New AccountArgs
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.BudgetLevelTextBox.SetFocus
    PRC.BudgetLevel = Me.BudgetLevelTextBox.text
    Dim SQL As String
    SQL = "UPDATE OperatingPlans SET OperatingPlans.BudgetLevel = '" _
        & PRC.BudgetLevel & "'" & _
        " WHERE OperatingPlans.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     UpdateBudgetLevel()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.BfyTextBox.SetFocus
    PRC.BFY = Me.BfyTextBox.text
    Dim SQL As String
    SQL = "UPDATE OperatingPlans SET OperatingPlans.BFY = '" _
        & PRC.BFY & "'" & _
        " WHERE OperatingPlans.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     UpdateBfy()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.RpioCodeTextBox.SetFocus
    PRC.RpioCode = Me.RpioCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE OperatingPlans SET OperatingPlans.RPIO = '" _
        & PRC.RpioCode & "'" & _
        " WHERE OperatingPlans.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     UpdateRpioCode()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.AhCodeTextBox.SetFocus
    PRC.AhCode = Me.AhCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE OperatingPlans SET OperatingPlans.AhCode = '" _
        & PRC.AhCode & "'" & _
        " WHERE OperatingPlans.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     UpdateAhCode()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.OrgCodeTextBox.SetFocus
    PRC.OrgCode = Me.OrgCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE OperatingPlans SET OperatingPlans.OrgCode = '" _
        & PRC.OrgCode & "'" & _
        " WHERE OperatingPlans.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     UpdateOrgCode()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.AccountCodeTextBox.SetFocus
    PRC.AccountCode = Me.AccountCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE OperatingPlans SET OperatingPlans.AccountCode = '" _
        & PRC.AccountCode & "'" & _
        " WHERE OperatingPlans.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     UpdateAccountCode()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.BocCodeTextBox.SetFocus
    PRC.BocCode = Me.BocCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE OperatingPlans SET OperatingPlans.BocCode = '" _
        & PRC.BocCode & "'" & _
        " WHERE OperatingPlans.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     UpdateBocCode()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.RcCodeTextBox.SetFocus
    PRC.RcCode = Me.RcCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE OperatingPlans SET OperatingPlans.RcCode = '" _
        & PRC.RcCode & "'" & _
        " WHERE OperatingPlans.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     UpdateRcCode()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.ID = CInt(Me!PrcIdTextBox.text)
    Me.AmountTextBox.SetFocus
    Dim Change As Double
    Change = CDbl(Me.AmountTextBox.text)
    Dim SQL As String
    SQL = "UPDATE OperatingPlans SET OperatingPlans.Amount = " _
        & Change & _
        " WHERE OperatingPlans.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:     OperatingPlans" _
            & vbCrLf & "Member:     UpdateAmount()" _
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
Private Function GetReportData() As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim mSQL As String
    Dim mData As DAO.Recordset
    mSQL = "SELECT Allocations.BudgetLevel AS [Level], Allocations.RPIO, Allocations.BFY," _
        & " Allocations.AhCode AS AH, Allocations.FundCode AS Fund," _
        & " Allocations.OrgCode AS ORG, Allocations.AccountCode AS Account, Allocations.BocCode AS BOC," _
        & " Allocations.Amount AS Amount" _
        & " FROM Allocations" _
        & " WHERE Allocations.Amount <> 0" & pAnd & Me.Filter
    Set mData = CurrentDb.OpenRecordset(mSQL)
    Set GetReportData = mData
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:       OperatingPlans" _
            & vbCrLf & "Member:     GetReportData()" _
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
Private Function GetColumnNames() As String()
    On Error GoTo ErrorHandler:
    Dim mFields As Collection
    Dim field As DAO.field
    Dim mData As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim mArray() As String
    Set mData = GetReportData
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
        mError = "Source:       OperatingPlans" _
            & vbCrLf & "Member:     GetColumnNames" _
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
    Dim mBudgetPath As BudgetPath
    Dim mExcel As Excel.Application
    Dim mAllocations As Excel.Workbook
    Dim mWorksheet As Excel.Worksheet
    Dim mList As Excel.ListObject
    Dim mRange As Excel.Range
    Dim mCell As Variant
    Dim mHeader As Excel.Range
    Dim mStart As Excel.Range
    Dim mEnd As Excel.Range
    Dim field As DAO.field
    Dim mData As DAO.Recordset
    Dim mFields() As String
    Set mData = GetReportData
    mFields() = GetColumnNames
    Dim i As Integer
    Dim j As Integer
    Set mBudgetPath = New BudgetPath
    Set mExcel = CreateObject("Excel.Application")
    Set mAllocations = mExcel.Workbooks.Open(mBudgetPath.ReportTemplate)
    mAllocations.Worksheets(2).Visible = False
    Set mWorksheet = mAllocations.Worksheets(1)
    mWorksheet.Name = "Allocations"
    mWorksheet.Cells.HorizontalAlignment = xlHAlignLeft
    mWorksheet.Cells.Font.Name = "Source Code Pro"
    mWorksheet.Cells.Font.Size = 8
    Set mStart = mWorksheet.Cells(1, 1)
    Set mEnd = mWorksheet.Cells(1, UBound(mFields) + 1)
    Set mHeader = mWorksheet.Range(mStart, mEnd)
    mHeader.Interior.Color = RGB(217, 217, 217)
    mHeader.HorizontalAlignment = xlHAlignLeft
    mHeader.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    mHeader.Borders(xlEdgeLeft).Color = RGB(133, 133, 133)
    mHeader.Borders(xlEdgeRight).Color = RGB(133, 133, 133)
    mHeader.Borders(xlEdgeTop).Color = RGB(133, 133, 133)
    mHeader.Font.Name = "Source Code Pro"
    mHeader.Font.Color = vbBlack
    mHeader.Font.Bold = True
    mHeader.Font.Size = 8
    For i = LBound(mFields) To UBound(mFields)
        j = i
        mHeader.Cells(1, i + 1).Value = mFields(j)
    Next i
    mWorksheet.Cells(2, 1).CopyFromRecordset mData
    mHeader.Font.Color = vbBlack
    mExcel.WindowState = xlMaximized
    mExcel.Visible = True
ErrorHandler:
    If Err.Number <> 0 Then
        mError = "Source:   OperatingPlans" _
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



