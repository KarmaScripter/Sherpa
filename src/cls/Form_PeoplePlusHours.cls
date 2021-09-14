VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PeoplePlusHours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As PayrollArgs
Public m_Dialog As Form_PayPeriodDialog
Private DivisionNameFilter As String
Private LastNameFilter As String
Private PayPeriodFilter As String
Private FundNameFilter As String
Private WorkCodeFilter As String
Private ReportingCodeFilter As String
Public pAnd As String
Private m_Error As String
Private m_Notification As String




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Open
'   Parameters:  Integer
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    pAnd = " AND "
    Me.Filter = vbNullString
    Set m_Dialog = New Form_PayPeriodDialog
    Set Args = New PayrollArgs
    DoCmd.OpenForm FormName:="PayPeriodDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("PayPeriodDialog")
    Set Args = m_Dialog.Args
    If Not Args.DivisionName & "" = "" And _
        Not Args.PayPeriod & "" = "" Then
            DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
            PayPeriodFilter = "[PayPeriod] = '" & Args.PayPeriod & "'"
            Me.PayPeriodComboBox.Value = Args.PayPeriod
            Me.Filter = DivisionNameFilter & pAnd & PayPeriodFilter
            Me.RecordSource = "SELECT *" _
                & " FROM PeoplePlus WHERE " & Me.Filter
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.PayPeriod & "" = "" Then
            DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
            Me.Filter = DivisionNameFilter
            Me.RecordSource = "SELECT *" _
                & " FROM PeoplePlus WHERE " & Me.Filter
    End If
    If Args.DivisionName & "" = "" And _
        Args.PayPeriod & "" = "" Then
            Me.Filter = "SELECT *" _
                & " FROM PeoplePlus;"
            Me.RecordSource = "SELECT *" _
                & " FROM PeoplePlus;"
    End If
    DoCmd.Close ObjectType:=acForm, ObjectName:="PayPeriodDialog", Save:=acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            Form_Load
'   Parameters:      Void
'   Retval:          Void
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
    ClearComboBoxValues
    SetButtonVisibility
    SetDivisionIcon
    Me.Caption = "People Plus Time And Attendance"
    Me.DivisionTitle.Caption = Args.DivisionName
    If Not Args.DivisionName & "" = "" And _
        Not Args.PayPeriod & "" = "" Then
            Me.Filter = DivisionNameFilter & pAnd & PayPeriodFilter
            Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName" _
                & " FROM PeoplePlus WHERE " & Me.Filter
            Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode" _
                & " FROM PeoplePlus WHERE " & Me.Filter
            Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName" _
                & " FROM PeoplePlus WHERE " & Me.Filter
            Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod" _
                & " FROM PeoplePlus WHERE " & Me.Filter
            Me.PayPeriodComboBox.Value = Args.PayPeriod
    Else
            Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName" _
                & " FROM PeoplePlus;"
            Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode" _
                & " FROM PeoplePlus;"
            Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName" _
                & " FROM PeoplePlus;"
            Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod" _
                & " FROM PeoplePlus;"
    End If
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
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            Form_Close
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    DivisionNameFilter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            LastNameComboBox_Change
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub LastNameComboBox_Change()
    On Error GoTo ErrorHandler:
    LastNameFilter = ""
    Me.Filter = ""
    LastNameFilter = "[LastName] = '" & Me.LastNameComboBox.Value & "'"
    If Not Args.DivisionName & "" = "" Then
        Me.Filter = DivisionNameFilter & pAnd & GetLastNameFilter
        Me.RecordSource = "SELECT * FROM PeoplePlus WHERE " & Me.Filter
        Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode FROM PeoplePlus WHERE " & Me.Filter
        Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName FROM PeoplePlus WHERE " & Me.Filter
        Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod FROM PeoplePlus WHERE " & Me.Filter
    Else
        Me.Filter = GetLastNameFilter
        Me.RecordSource = "SELECT * FROM PeoplePlus WHERE " & Me.Filter
        Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode FROM PeoplePlus WHERE " & Me.Filter
        Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName FROM PeoplePlus WHERE " & Me.Filter
        Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod FROM PeoplePlus WHERE " & Me.Filter
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetLastNameFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetLastNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not LastNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        Not PayPeriodFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not LastNameFilter = "" And _
        WorkCodeFilter = "" And _
        PayPeriodFilter = "" And _
        ReportingCodeFilter = "" Then
            GetLastNameFilter = LastNameFilter
    End If
    If Not LastNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        PayPeriodFilter = "" And _
        ReportingCodeFilter = "" Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not LastNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        Not PayPeriodFilter = "" And _
        ReportingCodeFilter = "" Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not LastNameFilter = "" And _
        WorkCodeFilter = "" And _
        PayPeriodFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not LastNameFilter = "" And _
        WorkCodeFilter = "" And _
        Not PayPeriodFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not LastNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        PayPeriodFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not LastNameFilter = "" And _
        WorkCodeFilter = "" And _
        Not PayPeriodFilter = "" And _
        ReportingCodeFilter = "" Then
            GetLastNameFilter = LastNameFilter _
                & pAnd & PayPeriodFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ReportingCodeComboBox_Change
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ReportingCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    ReportingCodeFilter = ""
    Me.Filter = ""
    ReportingCodeFilter = "[ReportingCodeName] = '" & Me.ReportingCodeComboBox.Value & "'"
    If Not Args.DivisionName & "" = "" Then
        Me.Filter = DivisionNameFilter & pAnd & GetReportingCodeFilter
        Me.RecordSource = "SELECT * FROM PeoplePlus WHERE " & Me.Filter
        Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode FROM PeoplePlus WHERE " & Me.Filter
        Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName FROM PeoplePlus WHERE " & Me.Filter
        Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod FROM PeoplePlus WHERE " & Me.Filter
    Else
        Me.Filter = GetReportingCodeFilter
        Me.RecordSource = "SELECT * FROM PeoplePlus WHERE " & Me.Filter
        Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode FROM PeoplePlus WHERE " & Me.Filter
        Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName FROM PeoplePlus WHERE " & Me.Filter
        Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod FROM PeoplePlus WHERE " & Me.Filter
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetReportingCodeFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetReportingCodeFilter() As String
    On Error GoTo ErrorHandler:
    If Not ReportingCodeFilter = "" And _
        Not LastNameFilter = "" And _
        Not PayPeriodFilter = "" And _
        Not WorkCodeFilter = "" Then
            GetReportingCodeFilter = ReportingCodeFilter _
                & pAnd & LastNameFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not ReportingCodeFilter = "" And _
        LastNameFilter = "" And _
        PayPeriodFilter = "" And _
        WorkCodeFilter = "" Then
            GetReportingCodeFilter = ReportingCodeFilter
    End If
    If Not ReportingCodeFilter = "" And _
        Not LastNameFilter = "" And _
        PayPeriodFilter = "" And _
        WorkCodeFilter = "" Then
            GetReportingCodeFilter = ReportingCodeFilter _
                & pAnd & LastNameFilter
    End If
    If Not ReportingCodeFilter = "" And _
        Not LastNameFilter = "" And _
        Not PayPeriodFilter = "" And _
        WorkCodeFilter = "" Then
            GetReportingCodeFilter = ReportingCodeFilter _
                & pAnd & LastNameFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not ReportingCodeFilter = "" And _
        LastNameFilter = "" And _
        PayPeriodFilter = "" And _
        Not WorkCodeFilter = "" Then
            GetReportingCodeFilter = ReportingCodeFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not ReportingCodeFilter = "" And _
        LastNameFilter = "" And _
        Not PayPeriodFilter = "" And _
        Not WorkCodeFilter = "" Then
            GetReportingCodeFilter = ReportingCodeFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not ReportingCodeFilter = "" And _
        Not LastNameFilter = "" And _
        PayPeriodFilter = "" And _
        Not WorkCodeFilter = "" Then
            GetReportingCodeFilter = ReportingCodeFilter _
                & pAnd & LastNameFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not ReportingCodeFilter = "" And _
        LastNameFilter = "" And _
        Not PayPeriodFilter = "" And _
        WorkCodeFilter = "" Then
            GetReportingCodeFilter = ReportingCodeFilter _
                & pAnd & PayPeriodFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            WorkCodeComboBox_Change
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub WorkCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    WorkCodeFilter = ""
    Me.Filter = ""
    WorkCodeFilter = "[WorkCode] = '" & Me.WorkCodeComboBox.Value & "'"
    If Not Args.DivisionName & "" = "" Then
        Me.Filter = DivisionNameFilter & pAnd & GetWorkCodeFilter
        Me.RecordSource = "SELECT * FROM PeoplePlus WHERE " & Me.Filter
        Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName FROM PeoplePlus WHERE " & Me.Filter
        Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName FROM PeoplePlus WHERE " & Me.Filter
        Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod FROM PeoplePlus WHERE " & Me.Filter
    Else
        Me.Filter = GetWorkCodeFilter
        Me.RecordSource = "SELECT * FROM PeoplePlus WHERE " & Me.Filter
        Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName FROM PeoplePlus WHERE " & Me.Filter
        Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName FROM PeoplePlus WHERE " & Me.Filter
        Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod FROM PeoplePlus WHERE " & Me.Filter
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetWorkCodeFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetWorkCodeFilter() As String
    On Error GoTo ErrorHandler:
    If Not WorkCodeFilter = "" And _
        Not LastNameFilter = "" And _
        Not PayPeriodFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & LastNameFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not WorkCodeFilter = "" And _
        LastNameFilter = "" And _
        PayPeriodFilter = "" And _
        ReportingCodeFilter = "" Then
            GetWorkCodeFilter = WorkCodeFilter
    End If
    If Not WorkCodeFilter = "" And _
        Not LastNameFilter = "" And _
        PayPeriodFilter = "" And _
        ReportingCodeFilter = "" Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & LastNameFilter
    End If
    If Not WorkCodeFilter = "" And _
        Not LastNameFilter = "" And _
        Not PayPeriodFilter = "" And _
        ReportingCodeFilter = "" Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & LastNameFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not WorkCodeFilter = "" And _
        LastNameFilter = "" And _
        PayPeriodFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not WorkCodeFilter = "" And _
        LastNameFilter = "" And _
        Not PayPeriodFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & PayPeriodFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not WorkCodeFilter = "" And _
        Not LastNameFilter = "" And _
        PayPeriodFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & LastNameFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not WorkCodeFilter = "" And _
        LastNameFilter = "" And _
        Not PayPeriodFilter = "" And _
        ReportingCodeFilter = "" Then
            GetWorkCodeFilter = WorkCodeFilter _
                & pAnd & PayPeriodFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            PayPeriodComboBox_Change
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayPeriodComboBox_Change()
    On Error GoTo ErrorHandler:
    PayPeriodFilter = ""
    Me.Filter = ""
    PayPeriodFilter = "[PayPeriod] = '" & Me.PayPeriodComboBox.Value & "'"
    If Not Args.DivisionName & "" = "" Then
        Me.Filter = DivisionNameFilter & pAnd & GetPayPeriodFilter
        Me.RecordSource = "SELECT * FROM PeoplePlus WHERE " & Me.Filter
        Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName FROM PeoplePlus WHERE " & Me.Filter
        Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName FROM PeoplePlus WHERE " & Me.Filter
        Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode FROM PeoplePlus WHERE " & Me.Filter
    Else
        Me.Filter = GetPayPeriodFilter
        Me.RecordSource = "SELECT * FROM PeoplePlus WHERE " & Me.Filter
        Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName FROM PeoplePlus WHERE " & Me.Filter
        Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName FROM PeoplePlus WHERE " & Me.Filter
        Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode FROM PeoplePlus WHERE " & Me.Filter
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetPayPeriodFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetPayPeriodFilter() As String
    On Error GoTo ErrorHandler:
    If Not PayPeriodFilter = "" And _
        Not LastNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & LastNameFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        LastNameFilter = "" And _
        WorkCodeFilter = "" And _
        ReportingCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter
    End If
    If Not PayPeriodFilter = "" And _
        Not LastNameFilter = "" And _
        WorkCodeFilter = "" And _
        ReportingCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & LastNameFilter
    End If
    If Not PayPeriodFilter = "" And _
        Not LastNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        ReportingCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & LastNameFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        LastNameFilter = "" And _
        WorkCodeFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        LastNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter _
                & pAnd & ReportingCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        LastNameFilter = "" And _
        Not WorkCodeFilter = "" And _
        ReportingCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & WorkCodeFilter
    End If
    If Not PayPeriodFilter = "" And _
        Not LastNameFilter = "" And _
        WorkCodeFilter = "" And _
        Not ReportingCodeFilter = "" Then
            GetPayPeriodFilter = PayPeriodFilter _
                & pAnd & LastNameFilter _
                & pAnd & ReportingCodeFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            StartDateTextBox_Click
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub StartDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.StartDateTextBox
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            SetButtonVisibility
'   Parameters:      Void
'   Retval:          Void
'   Purpose:         Toggles button visibility ON and OFF
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
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.LastNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.LastNameComboBox.BackColor = RGB(100, 100, 100)
    Me.LastNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.ReportingCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me.ReportingCodeComboBox.BackColor = RGB(33, 33, 33)
    Me.ReportingCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me.WorkCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me.WorkCodeComboBox.BackColor = RGB(33, 33, 33)
    Me.WorkCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me.PayPeriodComboBox.ForeColor = RGB(255, 255, 255)
    Me.PayPeriodComboBox.BackColor = RGB(33, 33, 33)
    Me.PayPeriodComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ClearComboBoxValues
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.WorkCodeComboBox.Value = vbNullString
    Me.LastNameComboBox.Value = vbNullString
    Me.ReportingCodeComboBox.Value = vbNullString
    Me.PayPeriodComboBox.Value = vbNullString
    If Not Args.DivisionName & "" = "" Then
        Me.Filter = DivisionNameFilter
        Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName" _
            & " FROM PeoplePlus WHERE " & Me.Filter
        Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode" _
            & " FROM PeoplePlus WHERE " & Me.Filter
        Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName" _
            & " FROM PeoplePlus WHERE " & Me.Filter
        Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod" _
            & " FROM PeoplePlus WHERE " & Me.Filter
    Else
        Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName" _
            & " FROM PeoplePlus;"
        Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode" _
            & " FROM PeoplePlus;"
        Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName" _
            & " FROM PeoplePlus;"
        Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod" _
            & " FROM PeoplePlus;"
    End If
    
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            MenuButton_Click
'   Parameters:      Void
'   Retval:          Void
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
'   Name:        EditButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:     Opens editor form based on current record data
'---------------------------------------------------------------------------------
Private Sub EditButton_Click()
    On Error GoTo ErrorHandler:
    Dim DataSet As DAO.Database
    Set DataSet = CurrentDb
    Dim Account As DAO.Recordset
    Dim ID As Integer
    ID = CInt(Me!CurrentIdTextBox.text)
    Me.RecordSource = "SELECT * FROM PeoplePlus WHERE PrcId = " & ID
    Dim frm As Form
    DoCmd.OpenForm ("AccountEditor")
    Form_AccountEditor.RecordSource = Me.RecordSource
    DoCmd.Close acForm, "AllocationForm", acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event / Sub-Procedure
'   Name:        CalculatorButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:     Launches the Windows 10 Calculator
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            EndDateTextBoxClick
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub EndDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.EndDateTextBox
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ClearFilterValues
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    LastNameFilter = vbNullString
    WorkCodeFilter = vbNullString
    ReportingCodeFilter = vbNullString
    PayPeriodFilter = vbNullString
    If Not Args.DivisionName & "" = "" Then
        Me.Filter = DivisionNameFilter
    Else
        Me.Filter = ""
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RefreshButton_Click
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearFilterValues
    ClearComboBoxValues
    Me.Filter = DivisionNameFilter
    Me.RecordSource = "SELECT * FROM PeoplePlus;"
    Me.LastNameComboBox.RowSource = "SELECT DISTINCT PeoplePlus.LastName" _
        & " FROM PeoplePlus WHERE " & Me.Filter
    Me.WorkCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.WorkCode" _
        & " FROM PeoplePlus WHERE " & Me.Filter
    Me.ReportingCodeComboBox.RowSource = "SELECT DISTINCT PeoplePlus.ReportingCodeName" _
        & " FROM PeoplePlus WHERE " & Me.Filter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PeoplePlus.PayPeriod" _
        & " FROM PeoplePlus WHERE " & Me.Filter
    Me.RecordSource = "SELECT *" _
        & " FROM PeoplePlus WHERE " & Me.Filter
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
    If Not Args.RcCode & "" = "" Then
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
            CurrentProject.Path & "\etc\png\AppIcons\interface\ui\employeetime.png"
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
'   Name:        GetReportData
'   Parameters:  Void
'   Retval:      DAO Recordset
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetReportData() As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim m_Data As DAO.Recordset
    Set m_Data = Me.Recordset
    m_Data.Filter = Me.Filter
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
    Dim p_Fields() As String
    p_Fields() = GetColumnNames
    Dim p_BudgetPath As BudgetPath
    Dim p_Excel As Excel.Application
    Dim p_Allocations As Excel.Workbook
    Dim p_Worksheet As Excel.Worksheet
    Dim p_List As Excel.ListObject
    Dim p_Range As Excel.Range
    Dim p_Cell As Object
    Dim p_Header As Excel.Range
    Dim p_Start As Excel.Range
    Dim p_End As Excel.Range
    Dim field As DAO.field
    Dim p_Data As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Set p_BudgetPath = New BudgetPath
    Set p_Excel = CreateObject("Excel.Application")
    Set p_Allocations = p_Excel.Workbooks.Open(p_BudgetPath.ReportTemplate)
    p_Allocations.Worksheets(2).Visible = False
    Set p_Worksheet = p_Allocations.Worksheets(1)
    p_Worksheet.Name = "People Plus"
    p_Worksheet.Cells.HorizontalAlignment = xlHAlignLeft
    p_Worksheet.Cells.Font.Name = "Source Code Pro"
    p_Worksheet.Cells.Font.Size = 8
    Set p_Start = p_Worksheet.Cells(1, 1)
    Set p_End = p_Worksheet.Cells(1, UBound(p_Fields) - 1)
    Set p_Header = p_Worksheet.Range(p_Start, p_End)
    p_Header.HorizontalAlignment = xlHAlignLeft
    p_Header.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    p_Header.Font.Name = "Source Code Pro"
    p_Header.Font.Color = vbBlack
    p_Header.Font.Bold = True
    p_Header.Font.Size = 8
    For i = LBound(p_Fields) To UBound(p_Fields)
        p_Header.Cells(i + 1).Value = p_Fields(i)
    Next i
    Set p_Data = Me.Recordset
    p_Data.Filter = Me.Filter
    p_Worksheet.Cells(2, 1).CopyFromRecordset p_Data
    p_Header.Font.Color = vbBlack
    p_Excel.WindowState = xlMaximized
    p_Excel.Visible = True
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





