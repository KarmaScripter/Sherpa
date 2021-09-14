VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TravelObligations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Public m_Dialog As Form_DivisionSelector
Public Args As AccountArgs
Private LastNameFilter As String
Private AccountCodeFilter As String
Private BocNameFilter As String
Private FocNameFilter As String
Private DivisionNameFilter As String
Private FiscalYearFilter As String
Private pAnd As String
Private m_Error As String
Private m_Notification As String






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
    Set m_Dialog = New Form_DivisionSelector
    Set Args = New AccountArgs
    DoCmd.OpenForm FormName:="DivisionSelector", WindowMode:=acDialog
    Set m_Dialog = Forms("DivisionSelector")
    Set Args = m_Dialog.Args
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
    ProcessError
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
    Me.LastNameComboBox.Value = vbNullString
    Me.FocNameComboBox.Value = vbNullString
    Me.AccountCodeComboBox.Value = vbNullString
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
    LastNameFilter = vbNullString
    FocNameFilter = vbNullString
    AccountCodeFilter = vbNullString
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
ErrorHandler:
    ProcessError
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
    ProcessError
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
    Me.RecordSource = "SELECT * FROM TravelObligations WHERE " & Me.Filter
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
    ProcessError
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
    ProcessError
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
    ProcessError
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
    ProcessError
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
Private Sub OriginalActionDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.OriginalActionDateTextBox
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
Private Sub ReturnDateTextBox_Enter()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.OriginalActionDateTextBox
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
    On Error GoTo ErrorHandler:
    If Not Args.RcCode = vbNullString Then
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
'   Name:        GetColumnNames
'   Parameters:  Void
'   Retval:      Collection
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetReportData() As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim m_SQL As String
    Dim m_Data As DAO.Recordset
    m_SQL = ""
    Set m_Data = CurrentDb.OpenRecordset(m_SQL)
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
    m_Worksheet.Name = "Travel Obligations"
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
        Set m_Excel = Nothing
        Set m_Allocations = Nothing
        Set m_Worksheet = Nothing
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




