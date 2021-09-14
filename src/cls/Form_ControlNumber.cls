VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ControlNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Args As BudgetArgs
Public DbData   As DAO.Database
Public DbSet As DAO.Recordset
Private m_Error As String
Private m_Notification As String


'----------------------------------------------------------------------------------
'   Type:       Sub-Procedure
'   Name:       SetComboBoxColors
'   Parameters: Void
'   Purpose:    Provides dark mode settings for the combobox controls
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.FundCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundCodeComboBox.BackColor = RGB(33, 33, 33)
    Me.FundCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me.RcCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me.RcCodeComboBox.BackColor = RGB(33, 33, 33)
    Me.RcCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me.FiscalYearComboBox.ForeColor = RGB(255, 255, 255)
    Me.FiscalYearComboBox.BackColor = RGB(33, 33, 33)
    Me.FiscalYearComboBox.BorderColor = RGB(68, 114, 196)
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
    Me.DeleteButton.Visible = Not Me.DeleteButton.Visible
    Me.CalculatorButton.Visible = Not Me.CalculatorButton.Visible
    Me.ExcelButton.Visible = Not Me.ExcelButton.Visible
    Me.UndoButton.Visible = Not Me.UndoButton.Visible
    Me.SaveButton.Visible = Not Me.SaveButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:       Sub Procedure
'   Name:       ClearComboBoxValues
'   Parameters: Void
'   Purpose:    Sets the selected values of the form's combobox control
'               to the empty string value
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.FundCodeComboBox.Value = ""
    Me.RcCodeComboBox.Value = ""
    Me.FiscalYearComboBox.Value = ""
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:       Event Sub-Procedure
'   Name:       CalcularButton Click
'   Parameters: Void
'   Purpose:    Launches Window 10 Calculator
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:       Event Sub-Procedure
'   Name:       DateIssuedTextBox DblClick
'   Parameters: Integer
'   Purpose:    Opens the BudgetCalendar Form
'---------------------------------------------------------------------------------
Private Sub DateIssuedTextBox_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Me.ActiveControl = Me.DateIssuedTextBox
    DoCmd.OpenForm FormName:="BudgetCalendar", _
        view:=acNormal, _
        OpenArgs:=Me.ActiveControl
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:       Event Sub-Procedure
'   Name:
'   Parameters:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.BFY = Me.FiscalYearComboBox.SelText
    Me.RcCodeComboBox.RowSource = "SELECT DISTINCT Allocations.RcCode" _
        & " FROM Allocations" _
        & " WHERE Allocations.BFY = '" & Args.BFY & "'"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:       Event Sub-Procedure
'   Name:
'   Parameters:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Section(acHeader).AutoHeight = False
    Me.Section(acHeader).Height = 2
    Me.Section(acDetail).AutoHeight = False
    Me.Section(acDetail).Height = 5
    Me.Section(acFooter).AutoHeight = False
    Me.Section(acFooter).Height = 1
    Set Args = New BudgetArgs
    ClearComboBoxValues
    SetComboBoxColors
    SetButtonVisibility
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:       Event Sub-Procedure
'   Name:
'   Parameters:
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
'   Type:       Event Sub-Procedure
'   Name:
'   Parameters:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RcCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.RcCode = Me.RcCodeComboBox.SelText
    Me.FundCodeComboBox.RowSource = "SELECT DISTINCT Allocations.FundCode" _
        & " FROM Allocations" _
        & " WHERE Allocations.BFY = '" & Args.BFY & "'" _
       & " AND Allocations.RcCode = '" & Args.RcCode & "'"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:       Event Sub-Procedure
'   Name:
'   Parameters:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    Me.RecordSource = "SELECT * FROM ControlNumbers;"
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





