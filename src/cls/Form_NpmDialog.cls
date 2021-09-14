VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_NpmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





Public Args As AccountArgs
Private m_BFY As String
Private m_RpioName As String
Private m_RpioCode As String
Private m_FundName As String
Private m_FundCode As String
Private pAnd As String
Private m_Error As String
Private m_Notification As String
Private m_RPIO As String




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    SetComboBoxColors
    pAnd = " AND "
    Set Args = New AccountArgs
    m_BFY = "[BFY] = '2021'"
    Me.RpioNameComboBox.RowSource = "SELECT DISTINCT ResourcePlanningOffices.Name" _
        & " FROM ResourcePlanningOffices " _
        & " WHERE ResourcePlanningOffices.Code NOT IN" _
        & " ('01', '02', '03', '04' ,'05', '06','07', '08', '09', '10')"
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Allocations.FundName" _
        & " FROM Allocations;"
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Allocations.BocName FROM Allocations;"
    Me.FiscalYearOptionGroup.DefaultValue = 2
    Me.FiscalYearOptionGroup.Value = 2
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    If Args Is Nothing Then
        Set Args = New AccountArgs
    End If
    Me.Filter = vbNullString
    m_BFY = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            m_BFY = "[BFY] = '2020'"
            Args.BFY = "2020"
            Me.Filter = GetFiscalYearFilter
        Case 2
            m_BFY = "[BFY] = '2021'"
            Args.BFY = "2021"
            Me.Filter = GetFiscalYearFilter
    End Select
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
Private Function GetFiscalYearFilter() As String
    On Error GoTo ErrorHandler:
    If Not m_BFY = "" And _
        Not m_RpioName = "" And _
        Not m_FundName = "" Then
            GetFiscalYearFilter = m_BFY _
                & pAnd & m_RpioName _
                & pAnd & m_FundName
    End If
    If Not m_BFY = "" And _
        m_RpioName = "" And _
        Not m_FundName = "" Then
            GetFiscalYearFilter = m_BFY _
                & pAnd & m_FundName
    End If
    If Not m_BFY = "" And _
        Not m_RpioName = "" And _
        m_FundName = "" Then
            GetFiscalYearFilter = m_BFY _
                & pAnd & m_RpioName
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RpioNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RpioNameComboBox_Change()
    On Error GoTo ErrorHandler:
    If Args Is Nothing Then
        Set Args = New AccountArgs
    End If
    Me.Filter = vbNullString
    m_RpioName = vbNullString
    Args.RpioName = Me.RpioNameComboBox.Value
    Args.RpioCode = DLookup("Code", "ResourcePlanningOffices", "ResourcePlanningOffices.Name = '" _
        & Me.RpioNameComboBox.Value & "'")
    m_RpioName = "[RpioName] = '" & Me.RpioNameComboBox.Value & "'"
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
Private Function GetDivisionNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not m_RpioName = "" And _
        Not m_FundName = "" And _
        Not m_BFY = "" Then
            GetDivisionNameFilter = m_RpioName _
                & pAnd & m_FundName _
                & pAnd & m_BFY
    End If
    If Not m_RpioName = "" And _
        m_FundName = "" And _
        Not m_BFY = "" Then
            GetDivisionNameFilter = m_RpioName _
                & pAnd & m_BFY
    End If
    If Not m_RpioName = "" And _
        Not m_FundName = "" And _
        m_BFY = "" Then
            GetDivisionNameFilter = m_RpioName _
                & pAnd & m_FundName
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
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    If Args Is Nothing Then
        Set Args = New AccountArgs
    End If
    Me.Filter = vbNullString
    m_FundName = vbNullString
    m_FundCode = DLookup("FundCode", "Allocations", _
        "Allocations.FundName = '" & Me.FundNameComboBox.Value & "'")
    m_FundName = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    Args.FundName = Me.FundNameComboBox.Value
    Args.FundCode = m_FundCode
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
    If Not m_FundName = "" And _
        Not m_RpioName = "" And _
        Not m_BFY = "" Then
            GetFundNameFilter = m_FundName _
                & pAnd & m_RpioName _
                & pAnd & m_BFY
    End If
    If Not m_FundName = "" And _
        m_RpioName = "" And _
        Not m_BFY = "" Then
            GetFundNameFilter = m_FundName _
                & pAnd & m_BFY
    End If
    If Not m_FundName = "" And _
        Not m_RpioName = "" And _
        m_BFY = "" Then
            GetFundNameFilter = m_FundName _
                & pAnd & m_RpioName
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExecuteButton_OnClick
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
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ClearOptionGroupValues
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetComboBoxColors
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me!RpioNameComboBox.ForeColor = RGB(255, 255, 255)
    Me!RpioNameComboBox.BackColor = RGB(33, 33, 33)
    Me!RpioNameComboBox.BorderColor = RGB(68, 114, 196)
    Me!FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me!FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me!FundNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearOptionGroupValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearOptionGroupValues()
    If Not IsNull(Me.FiscalYearOptionGroup) And _
        Not Me.FiscalYearOptionGroup.Value & "" = vbNullString Then
            Me.FiscalYearOptionGroup.Value = vbNullString
            m_BFY = vbNullString
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearComboBoxValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    Me.RpioNameComboBox.Value = vbNullString
    Me.FundNameComboBox.Value = vbNullString
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
    m_BFY = vbNullString
    m_RpioName = vbNullString
    m_FundName = vbNullString
    Me.Filter = vbNullString
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
    Set Args = Nothing
    DoCmd.Close acForm, Me.Name
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





