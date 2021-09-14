VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PersonnelData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private SectionNameFilter As String
Private JobTitleFilter As String
Private PayGradeFilter As String
Private StepFilter As String
Public Records As Integer
Private pAnd As String
Private m_Error As String
Private m_Notification As String







'----------------------------------------------------------------------------------
'   Type:            Event Sub-Procedure
'   Name:            Form_Load
'   Parameters:      Void
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
    SetButtonVisibility
    ClearFilterValues
    ClearComboBoxValues
    Me.RecordSource = "SELECT * FROM DivisionPersonnel;"
    Me.SectionNameComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.HrOrgName FROM DivisionPersonnel;"
    Me.JobTitleComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.JobTitle FROM DivisionPersonnel;"
    Me.PayGradeComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.Grade FROM DivisionPersonnel" _
        & " WHERE DivisionPersonnel.Grade <> '00';"
    Me.StepComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.Step FROM DivisionPersonnel;"
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
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            JobTitleComboBox_Click
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub JobTitleComboBox_Change()
    On Error GoTo ErrorHandler:
    JobTitleFilter = vbNullString
    Me.Filter = vbNullString
    JobTitleFilter = "[JobTitle] = '" & Me.JobTitleComboBox.Value & "'"
    Me.Filter = GetJobTitleFilter
    Me.RecordSource = "SELECT * FROM DivisionPersonnel WHERE " & Me.Filter
    Me.Requery
    Me.SectionNameComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.HrOrgName FROM DivisionPersonnel WHERE " & Me.Filter
    Me.PayGradeComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.Grade FROM DivisionPersonnel WHERE " & Me.Filter
    Me.StepComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.Step FROM DivisionPersonnel WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetJobTitleFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetJobTitleFilter() As String
    On Error GoTo ErrorHandler:
    If Not JobTitleFilter = vbNullString And _
        Not SectionNameFilter = vbNullString And _
        Not PayGradeFilter = vbNullString And _
        Not StepFilter = vbNullString Then
            GetJobTitleFilter = JobTitleFilter _
                & pAnd & SectionNameFilter _
                & pAnd & PayGradeFilter _
                & pAnd & StepFilter
    End If
    If Not JobTitleFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        PayGradeFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetJobTitleFilter = JobTitleFilter
    End If
    If Not JobTitleFilter = vbNullString And _
        Not SectionNameFilter = vbNullString And _
        PayGradeFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetJobTitleFilter = JobTitleFilter _
                & pAnd & SectionNameFilter
    End If
    If Not JobTitleFilter = vbNullString And _
        Not SectionNameFilter = vbNullString And _
        Not PayGradeFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetJobTitleFilter = JobTitleFilter _
                & pAnd & SectionNameFilter _
                & pAnd & PayGradeFilter
    End If
    If Not JobTitleFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        Not PayGradeFilter = vbNullString And _
        Not StepFilter = vbNullString Then
            GetJobTitleFilter = JobTitleFilter _
                & pAnd & PayGradeFilter _
                & pAnd & StepFilter
    End If
    If Not JobTitleFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        Not PayGradeFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetJobTitleFilter = JobTitleFilter _
                & pAnd & PayGradeFilter
    End If
    If Not JobTitleFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        PayGradeFilter = vbNullString And _
        Not StepFilter = vbNullString Then
            GetJobTitleFilter = JobTitleFilter _
                & pAnd & StepFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            SectionNameComboBox_Click
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SectionNameComboBox_Change()
    On Error GoTo ErrorHandler:
    SectionNameFilter = vbNullString
    Me.Filter = vbNullString
    SectionNameFilter = "[HrOrgName] = '" & Me.SectionNameComboBox.Value & "'"
    Me.Filter = GetSectionNameFilter
    Me.RecordSource = "SELECT * FROM DivisionPersonnel WHERE " & Me.Filter
    Me.Requery
    Me.PayGradeComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.Grade FROM DivisionPersonnel WHERE " & Me.Filter
    Me.StepComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.Step FROM DivisionPersonnel WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetSectionNameFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetSectionNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        Not PayGradeFilter = vbNullString And _
        Not StepFilter = vbNullString Then
            GetSectionNameFilter = SectionNameFilter _
                & pAnd & JobTitleFilter _
                & pAnd & PayGradeFilter _
                & pAnd & StepFilter
    End If
    If Not SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        PayGradeFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetSectionNameFilter = SectionNameFilter
    End If
    If Not SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        PayGradeFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetSectionNameFilter = SectionNameFilter _
                & pAnd & JobTitleFilter
    End If
    If Not SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        Not PayGradeFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetSectionNameFilter = SectionNameFilter _
                & pAnd & JobTitleFilter _
                & pAnd & PayGradeFilter
    End If
    If Not SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        Not PayGradeFilter = vbNullString And _
        Not StepFilter = vbNullString Then
            GetSectionNameFilter = SectionNameFilter _
                & pAnd & PayGradeFilter _
                & pAnd & StepFilter
    End If
    If Not SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        Not PayGradeFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetSectionNameFilter = SectionNameFilter _
                & pAnd & PayGradeFilter
    End If
    If Not SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        PayGradeFilter = vbNullString And _
        Not StepFilter = vbNullString Then
            GetSectionNameFilter = SectionNameFilter _
                & pAnd & StepFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            PayGradeComboBox_Click
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayGradeComboBox_Change()
    On Error GoTo ErrorHandler:
    PayGradeFilter = vbNullString
    Me.Filter = vbNullString
    PayGradeFilter = "[Grade] = '" & Me.PayGradeComboBox.Value & "'"
    Me.Filter = GetPayGradeFilter
    Me.RecordSource = "SELECT * FROM DivisionPersonnel WHERE " & Me.Filter
    Me.Requery
    Me.StepComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.Step FROM DivisionPersonnel WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetPayGradeFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetPayGradeFilter() As String
    On Error GoTo ErrorHandler:
    If Not PayGradeFilter = vbNullString And _
        Not SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        Not StepFilter = vbNullString Then
            GetPayGradeFilter = PayGradeFilter _
                & pAnd & SectionNameFilter _
                & pAnd & JobTitleFilter _
                & pAnd & StepFilter
    End If
    If Not PayGradeFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetPayGradeFilter = PayGradeFilter
    End If
    If Not PayGradeFilter = vbNullString And _
        Not SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetPayGradeFilter = PayGradeFilter _
                & pAnd & SectionNameFilter
    End If
    If Not PayGradeFilter = vbNullString And _
        Not SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetPayGradeFilter = PayGradeFilter _
                & pAnd & SectionNameFilter _
                & pAnd & JobTitleFilter
    End If
    If Not PayGradeFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        Not StepFilter = vbNullString Then
            GetPayGradeFilter = PayGradeFilter _
                & pAnd & JobTitleFilter _
                & pAnd & StepFilter
    End If
    If Not PayGradeFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        StepFilter = vbNullString Then
            GetPayGradeFilter = PayGradeFilter _
                & pAnd & JobTitleFilter
    End If
    If Not PayGradeFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        Not StepFilter = vbNullString Then
            GetPayGradeFilter = PayGradeFilter _
                & pAnd & StepFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            StepComboBox_Click
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub StepComboBox_Change()
    On Error GoTo ErrorHandler:
    StepFilter = vbNullString
    Me.Filter = vbNullString
    StepFilter = "[Step] = '" & Me.StepComboBox.Value & "'"
    Me.Filter = GetStepFilter
    Me.RecordSource = "SELECT * FROM DivisionPersonnel WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetStepFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetStepFilter() As String
    On Error GoTo ErrorHandler:
    If Not StepFilter = vbNullString And _
        Not SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        Not PayGradeFilter = vbNullString Then
            GetStepFilter = StepFilter _
                & pAnd & SectionNameFilter _
                & pAnd & JobTitleFilter _
                & pAnd & PayGradeFilter
    End If
    If Not StepFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        PayGradeFilter = vbNullString Then
            GetStepFilter = StepFilter
    End If
    If Not StepFilter = vbNullString And _
        Not SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        PayGradeFilter = vbNullString Then
            GetStepFilter = StepFilter _
                & pAnd & SectionNameFilter
    End If
    If Not StepFilter = vbNullString And _
        Not SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        PayGradeFilter = vbNullString Then
            GetStepFilter = StepFilter _
                & pAnd & SectionNameFilter _
                & pAnd & JobTitleFilter
    End If
    If Not StepFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        Not PayGradeFilter = vbNullString Then
            GetStepFilter = StepFilter _
                & pAnd & JobTitleFilter _
                & pAnd & PayGradeFilter
    End If
    If Not StepFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        Not JobTitleFilter = vbNullString And _
        PayGradeFilter = vbNullString Then
            GetStepFilter = StepFilter _
                & pAnd & JobTitleFilter
    End If
    If Not StepFilter = vbNullString And _
        SectionNameFilter = vbNullString And _
        JobTitleFilter = vbNullString And _
        Not PayGradeFilter = vbNullString Then
            GetStepFilter = StepFilter _
                & pAnd & PayGradeFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            SetButtonVisibility
'   Parameters:      Void
'   Purpose:         Toggles button visibility ON and OFF
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
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
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    SectionNameFilter = vbNullString
    JobTitleFilter = vbNullString
    PayGradeFilter = vbNullString
    StepFilter = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.SectionNameComboBox.Value = vbNullString
    Me.JobTitleComboBox.Value = vbNullString
    Me.PayGradeComboBox.Value = vbNullString
    Me.StepComboBox.Value = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Event Sub-Procedure
'   Name:            CalculatorButton_Click
'   Parameters:      Void
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
'   Type:            Event Sub-Procedure
'   Name:            GradeEntryDateTextBox_Click
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub GradeEntryDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.GradeEntryDateTextBox
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub HireDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.HireDateTextBox
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Event Sub-Procedure
'   Name:            LastIncreaseDateTextBox_Click
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub LastIncreaseDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.LastIncreaseDateTextBox
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Event Sub-Procedure
'   Name:            MenuButton_Click
'   Parameters:      Void
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
'   Type:            Event Sub-Procedure
'   Name:            WigiDueDateTextBox_Click
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub WigiDueDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.WigiDueDateTextBox
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:            Event Sub-Procedure
'   Name:            RefreshButton_Click
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    Me.RecordSource = "SELECT * FROM DivisionPersonnel;"
    Me.SectionNameComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.HrOrgName FROM DivisionPersonnel;"
    Me.JobTitleComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.JobTitle FROM DivisionPersonnel;"
    Me.PayGradeComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.Grade FROM DivisionPersonnel" _
        & " WHERE DivisionPersonnel.Grade <> '00';"
    Me.StepComboBox.RowSource = "SELECT DISTINCT DivisionPersonnel.Step FROM DivisionPersonnel;"
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




