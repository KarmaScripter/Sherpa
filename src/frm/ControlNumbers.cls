Option Compare Database
Option Explicit


Private Budget As BudgetArgs
Public DbData   As DAO.Database
Public DbSet As DAO.Recordset
Private mError As String
Private mNotification As String

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
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: SetButtonVisibility" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: ClearComboBoxValues()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: CalculatorButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: DateIssuedTextBox_DblClick()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Budget.BFY = Me.FiscalYearComboBox.SelText
    Me.RcCodeComboBox.RowSource = "SELECT DISTINCT Allocations.RcCode" _
        & " FROM Allocations" _
        & " WHERE Allocations.BFY = '" & Budget.BFY & "'"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: FiscalYearComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Set Budget = New BudgetArgs
    ClearComboBoxValues
    SetComboBoxColors
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: MenuButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Budget.RcCode = Me.RcCodeComboBox.SelText
    Me.FundCodeComboBox.RowSource = "SELECT DISTINCT Allocations.FundCode" _
        & " FROM Allocations" _
        & " WHERE Allocations.BFY = '" & Budget.BFY & "'" _
       & " AND Allocations.RcCode = '" & Budget.RcCode & "'"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: RcCodeComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   ControlNumbers" _
            & vbCrLf & "Member: RefreshButton_Click()" _
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



