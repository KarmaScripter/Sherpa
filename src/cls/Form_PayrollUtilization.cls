VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PayrollUtilization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public m_Dialog As Form_FundDialog
Public Args As AccountArgs
Private FundNameFilter As String
Private ProgramAreaNameFilter As String
Private ProgramProjectNameFilter As String
Private FiscalYearFilter As String
Private pAnd As String
Private m_Error As String
Private m_Notification As String








'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        Form_Open
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set m_Dialog = New Form_FundDialog
    DoCmd.OpenForm FormName:="FundDialog", _
        WindowMode:=acDialog
    Set m_Dialog = Forms("FundDialog")
    pAnd = " AND "
    Set Args = New AccountArgs
    Set Args = m_Dialog.Args
    If Not Args Is Nothing Then
        FundNameFilter = "[FundName] = '" & Args.FundName & "'"
        Me.FiscalYearOptionGroup.Value = 2
        FiscalYearFilter = "[BFY] = '2021'"
        Me.Filter = FiscalYearFilter & pAnd & FundNameFilter
    Else
        Me.FiscalYearOptionGroup.Value = 2
        FiscalYearFilter = "[BFY] = '2021'"
        Me.Filter = FiscalYearFilter
    End If
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramProjectName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & Me.Filter
    Me.ProgramAreaNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramAreaName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.FundName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & Me.Filter
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & Me.Filter
    DoCmd.Close ObjectType:=acForm, ObjectName:="FundDialog", Save:=acSaveNo
ErrorHandler:
    ProcessError
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
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler:
    ClearFilterValues
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
    ClearComboBoxValues
    SetComboBoxColors
    SetButtonVisibility
    SetHeaderCaption
ErrorHandler:
    ProcessError
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
    Me.FirstButton.Visible = Not Me.FirstButton.Visible
    Me.PreviousButton.Visible = Not Me.PreviousButton.Visible
    Me.NextButton.Visible = Not Me.NextButton.Visible
    Me.LastButton.Visible = Not Me.LastButton.Visible
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.CalculatorButton.Visible = Not Me.CalculatorButton.Visible
    Me.ExcelButton.Visible = Not Me.ExcelButton.Visible
    Me.DataButton.Visible = Not Me.DataButton.Visible
ErrorHandler:
    ProcessError
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
    On Error GoTo ErrorHandler::
    Args.BFY = vbNullString
    FiscalYearFilter = vbNullString
    Me.Filter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            Args.BFY = "2020"
            FiscalYearFilter = "[BFY] = '2020'"
        Case 2
            Args.BFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
    End Select
    Me.Filter = GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Allocations WHERE " & Me.Filter
    Me.Requery
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
'---------------------------------------------------------------------------------
Private Sub HideButtons()
    On Error GoTo ErrorHandler:
    Me.DataButton.Visible = False
    Me.FirstButton.Visible = False
    Me.PreviousButton.Visible = False
    Me.NextButton.Visible = False
    Me.LastButton.Visible = False
    Me.RefreshButton.Visible = False
    Me.CalculatorButton.Visible = False
    Me.ExcelButton.Visible = False
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
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.ProgramAreaNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.ProgramAreaNameComboBox.BackColor = RGB(33, 33, 33)
    Me.ProgramAreaNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.ProgramProjectNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.ProgramProjectNameComboBox.BackColor = RGB(33, 33, 33)
    Me.ProgramProjectNameComboBox.BorderColor = RGB(68, 114, 196)
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
    Me.FundNameComboBox.Value = vbNullString
    Me.ProgramAreaNameComboBox.Value = vbNullString
    Me.ProgramProjectNameComboBox.Value = vbNullString
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
    FundNameFilter = vbNullString
    ProgramAreaNameFilter = vbNullString
    FiscalYearFilter = vbNullString
    ProgramProjectNameFilter = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        ProgramAreaNameComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ProgramAreaNameComboBox_Change()
    On Error GoTo ErrorHandler:
    ProgramAreaNameFilter = vbNullString
    Me.Filter = vbNullString
    ProgramAreaNameFilter = "[ProgramAreaName] = '" & Me.ProgramAreaNameComboBox.Value & "'"
    Me.Filter = GetProgramAreaNameFilter
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7'" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetProgramAreaNameFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetProgramAreaNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not ProgramAreaNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramAreaNameFilter = ProgramAreaNameFilter
    End If
    If Not ProgramAreaNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramAreaNameFilter = ProgramAreaNameFilter _
            & pAnd & FiscalYearFilter
    End If
    If Not ProgramAreaNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramAreaNameFilter = ProgramAreaNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramAreaNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramAreaNameFilter = ProgramAreaNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not ProgramAreaNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramAreaNameFilter = ProgramAreaNameFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not ProgramAreaNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramAreaNameFilter = ProgramAreaNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not ProgramAreaNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramAreaNameFilter = ProgramAreaNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramAreaNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramAreaNameFilter = ProgramAreaNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        ProgramProjectNameComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ProgramProjectNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    ProgramProjectNameFilter = vbNullString
    ProgramProjectNameFilter = "[ProgramProjectName] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    If Not IsNull(Me.OpenArgs) Then
        Me.Filter = Me.OpenArgs & pAnd & GetProgramProjectNameFilter
    Else
        Me.Filter = GetProgramProjectNameFilter
    End If
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE " & Me.Filter
    Me.Requery
    SetHeaderCaption
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetProgramProjectNameFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetProgramProjectNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not ProgramProjectNameFilter = "" And _
        Not FundNameFilter = "" And _
        Not ProgramAreaNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramAreaNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        FundNameFilter = "" And _
        ProgramAreaNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        Not FundNameFilter = "" And _
        ProgramAreaNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        Not FundNameFilter = "" And _
        Not ProgramAreaNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramAreaNameFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        FundNameFilter = "" And _
        ProgramAreaNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        FundNameFilter = "" And _
        Not ProgramAreaNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & ProgramAreaNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        FundNameFilter = "" And _
        Not ProgramAreaNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & ProgramAreaNameFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        Not FundNameFilter = "" And _
        ProgramAreaNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
ErrorHandler:
    ProcessError
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
    If Not FiscalYearFilter = vbNullString And _
        ProgramAreaNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not ProgramAreaNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
            & pAnd & ProgramAreaNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not ProgramAreaNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramAreaNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not ProgramAreaNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & ProgramAreaNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        ProgramAreaNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        ProgramAreaNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        ProgramAreaNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not ProgramAreaNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramAreaNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        ProgramAreaNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    ProcessError
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
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    FundNameFilter = vbNullString
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & Me.Filter
    Me.Requery
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramProjectName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7'" _
        & pAnd & Me.Filter
    Me.ProgramAreaNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramAreaName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7'" _
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
    If Not FundNameFilter = vbNullString And _
        ProgramAreaNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not ProgramAreaNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
            & pAnd & ProgramAreaNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not ProgramAreaNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramAreaNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not ProgramAreaNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramAreaNameFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = vbNullString And _
        ProgramAreaNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = vbNullString And _
        ProgramAreaNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = vbNullString And _
        ProgramAreaNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not ProgramAreaNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramAreaNameFilter _
                & pAnd & FiscalYearFilter
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
'   Name        RefreshButton_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    Me.FiscalYearOptionGroup.Value = 2
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramProjectName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7';"
    Me.ProgramAreaNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramAreaName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7';"
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.FundName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7';"
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7';"
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
    On Error GoTo ErrorHandler::
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



Private Sub SetHeaderCaption()
    On Error GoTo ErrorHandler:
    If Not Me.ProgramProjectNameComboBox.Value & "" = "" Then
            Me.ProgramAreaHeader.Caption = DLookup("ProgramAreaName", "ProgramDescriptions", _
                "[ProgramDescriptions.ProgramProjectName] = '" & Me.ProgramProjectNameComboBox.Value & "'")
    Else
        Me.ProgramAreaHeader.Caption = "Payroll Utilization"
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





