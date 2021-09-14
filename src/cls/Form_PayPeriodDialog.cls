VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PayPeriodDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As PayrollArgs
Private FiscalYearFilter As String
Private FundNameFilter As String
Private DivisionNameFilter As String
Private PayPeriodFilter As String
Private pAnd As String
Private m_Error As String
Private m_Notification As String







'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    pAnd = " AND "
    Set Args = New PayrollArgs
    FiscalYearFilter = "[BFY] = '2021'"
    Me.Filter = FiscalYearFilter
    Args.BFY = "2021"
    Me.DivisionNameComboBox.RowSource = "SELECT DISTINCT UCase(PayrollObligations.DivisionName)" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT UCase(PayrollObligations.FundName)" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollObligations.PayPeriod" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetComboBoxColors
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.DivisionNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.DivisionNameComboBox.BackColor = RGB(33, 33, 33)
    Me.DivisionNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.PayPeriodComboBox.ForeColor = RGB(255, 255, 255)
    Me.PayPeriodComboBox.BackColor = RGB(33, 33, 33)
    Me.PayPeriodComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        AcceptButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AcceptButton_Click()
    On Error GoTo ErrorHandler:
    Me.Visible = Not Me.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExitButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    ClearFilterValues
    DoCmd.Close acForm, Me.Name
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DivisionNameComboBox_Change()
    On Error GoTo ErrorHandler:
    DivisionNameFilter = vbNullString
    Args.DivisionName = Me.DivisionNameComboBox.Value
    Args.RcCode = DLookup("RcCode", "PayrollObligations", "PayrollObligations.DivisionName = '" _
        & Me.DivisionNameComboBox.Value & "'")
    DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
    Me.Filter = GetDivisionNameFilter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT UCase(PayrollObligations.FundName)" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollObligations.PayPeriod" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetDivisionNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetDivisionNameFilter()
    On Error GoTo ErrorHandler:
    If Not DivisionNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString Then
            GetDivisionNameFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString Then
            GetDivisionNameFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString Then
            GetDivisionNameFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString Then
            GetDivisionNameFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter
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
    Args.FundName = Me.FundNameComboBox.Value
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollObligations.PayPeriod" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
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
        Not FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & DivisionNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & PayPeriodFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        DivisionNameFilter = vbNullString And _
        PayPeriodFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        DivisionNameFilter = vbNullString And _
        Not PayPeriodFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & PayPeriodFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayPeriodComboBox_Change()
    On Error GoTo ErrorHandler:
    PayPeriodFilter = vbNullString
    Args.PayPeriod = Me.PayPeriodComboBox.Value
    PayPeriodFilter = "[PayPeriod] = '" & Args.PayPeriod & "'"
    Me.Filter = GetPayPeriodFilter
    Me.DivisionNameComboBox.RowSource = "SELECT DISTINCT UCase(PayrollObligations.DivisionName)" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT UCase(PayrollObligations.FundName)" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
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
    If Not DivisionNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        Not DivisionNameFilter = "" And _
        Not FundNameFilter = "" Then
            GetPayPeriodFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not DivisionNameFilter = "" And _
        FiscalYearFilter = "" And _
        DivisionNameFilter = "" And _
        FundNameFilter = "" Then
            GetPayPeriodFilter = DivisionNameFilter
    End If
    If Not DivisionNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        DivisionNameFilter = "" And _
        FundNameFilter = "" Then
            GetPayPeriodFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not DivisionNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        Not DivisionNameFilter = "" And _
        FundNameFilter = "" Then
            GetPayPeriodFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & DivisionNameFilter
    End If
    If Not DivisionNameFilter = "" And _
        FiscalYearFilter = "" And _
        DivisionNameFilter = "" And _
        Not FundNameFilter = "" Then
            GetPayPeriodFilter = DivisionNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not DivisionNameFilter = "" And _
        FiscalYearFilter = "" And _
        Not DivisionNameFilter = "" And _
        Not FundNameFilter = "" Then
            GetPayPeriodFilter = DivisionNameFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not DivisionNameFilter = "" And _
        FiscalYearFilter = "" And _
        Not DivisionNameFilter = "" And _
        FundNameFilter = "" Then
            GetPayPeriodFilter = DivisionNameFilter _
                & pAnd & DivisionNameFilter
    End If
    If Not DivisionNameFilter = "" And _
        Not FiscalYearFilter = "" And _
        DivisionNameFilter = "" And _
        Not FundNameFilter = "" Then
            GetPayPeriodFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ResetOptionGroupValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
    On Error GoTo ErrorHandler:
    FiscalYearFilter = vbNullString
    Me.FiscalYearOptionGroup.Value = 2
    FiscalYearFilter = "[BFY] = '2021'"
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
    On Error GoTo ErrorHandler:
    FiscalYearFilter = vbNullString
    Args.BFY = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            Args.BFY = "2020"
            FiscalYearFilter = "[BFY] = '2020'"
        Case 2
            Args.BFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
    End Select
    Me.Filter = GetFiscalYearFilter
    Me.DivisionNameComboBox.RowSource = "SELECT DISTINCT UCase(PayrollObligations.DivisionName)" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT UCase(PayrollObligations.FundName)" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollObligations.PayPeriod" _
        & " FROM PayrollObligations" _
        & " WHERE " & Me.Filter
    Me.RecordSource = "SELECT * FROM PayrollObligations" _
        & " WHERE " & Me.Filter
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
        DivisionNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & DivisionNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & DivisionNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & DivisionNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & DivisionNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        DivisionNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & DivisionNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & DivisionNameFilter _
                & pAnd & DivisionNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        DivisionNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        DivisionNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & DivisionNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ResetOptionGroupValues
    Me.DivisionNameComboBox.RowSource = "SELECT DISTINCT UCase(PayrollObligations.DivisionName)" _
        & " FROM PayrollObligations;"
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT UCase(PayrollObligations.FundName)" _
        & " FROM PayrollObligations;"
    Me.PayPeriodComboBox.RowSource = "SELECT DISTINCT PayrollObligations.PayPeriod" _
        & " FROM PayrollObligations;"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ClearComboBoxValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.Value = vbNullString
    Me.DivisionNameComboBox.Value = vbNullString
    Me.PayPeriodComboBox.Value = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearFilterValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    DivisionNameFilter = vbNullString
    FundNameFilter = vbNullString
    PayPeriodFilter = vbNullString
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





