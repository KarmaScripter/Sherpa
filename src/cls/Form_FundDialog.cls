VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FundDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Args As AccountArgs
Private pAnd As String
Private m_FundCode As String
Private m_FundName As String
Private FundNameFilter As String
Private m_BFY As String
Private FiscalYearFilter As String
Private m_And As String
Private m_Filter As String
Private m_Error As String
Private m_Notification As String




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    SetComboBoxColors
    ClearComboBoxValues
    pAnd = " AND "
    Set Args = New AccountArgs
    Me.FiscalYearOptionGroup.DefaultValue = 2
    FiscalYearFilter = "[BFY] = '2021'"
    Args.BFY = "2021"
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT UCase(Allocations.FundName)" _
        & " FROM Allocations" _
        & " WHERE Allocations.BudgetLevel = '7'" _
        & " AND Allocations.BocCode = '10'" _
        & pAnd & FiscalYearFilter
    Me.FundNameComboBox.Value = vbNullString
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
Private Sub Form_Deactivate()
    On Error GoTo ErrorHandler:
    ClearFilterValues
    Me.FundNameComboBox.Value = vbNullString
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
    Me.FundNameComboBox.Value = vbNullString
    FundNameFilter = vbNullString
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
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    If Not Args Is Nothing Then
        m_FundName = vbNullString
        Me.Filter = vbNullString
        FundNameFilter = vbNullString
        FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
        Args.FundName = vbNullString
        Args.FundName = Me.FundNameComboBox.Value
        Args.FundCode = DLookup("FundCode", "Allocations", "Allocations.FundName = '" _
            & Me.FundNameComboBox.Value & "'")
    End If
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
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




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
    On Error GoTo ErrorHandler::
    If Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
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
    Me.FiscalYearOptionGroup.Value = 2
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ClearComboBoxValues
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    ClearFilterValues
    Me.FundNameComboBox.Value = vbNullString
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT UCase(Allocations.FundName)" _
        & " FROM Allocations;"
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
    FundNameFilter = vbNullString
    m_FundName = vbNullString
    Me.Filter = vbNullString
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







