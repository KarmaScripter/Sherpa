VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DivisionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As AccountArgs
Public FiscalYearFilter As String
Public BocNameFilter As String
Public FundNameFilter As String
Private m_FundName As String
Private m_BocName As String
Private m_BFY As String
Private pAnd As String
Private m_Error As String






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
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Allocations.FundName FROM Allocations"
    Me.ObjectClassComboBox.RowSource = "SELECT DISTINCT Allocations.BocName FROM Allocations" _
        & " WHERE Allocations.BocCode IN ('21', '28', '36', '37', '38', '41')"
    Me.FiscalYearOptionGroup.DefaultValue = 2
    Args.BFY = "2021"
    FiscalYearFilter = "[BFY] = '2021'"
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
    m_BFY = vbNull
    Me.Filter = vbNullString
    FiscalYearFilter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            m_BFY = "2020"
            FiscalYearFilter = "[BFY] = '2020'"
            Args.BFY = m_BFY
            Me.Filter = GetFiscalYearFilter
        Case 2
            m_BFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
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
    If Not FiscalYearFilter = "" And _
        Not BocNameFilter = "" And _
        Not FundNameFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = "" And _
        BocNameFilter = "" And _
        Not FundNameFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = "" And _
        Not BocNameFilter = "" And _
        FundNameFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ObjectClassComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ObjectClassComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    BocNameFilter = vbNullString
    m_BocName = Me.ObjectClassComboBox.Value
    Args.BocName = m_BocName
    BocNameFilter = "[BocName] = '" & Me.ObjectClassComboBox.Value & "'"
    Me.Filter = GetBocNameFilter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetBocNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not BocNameFilter = "" And _
        Not FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If BocNameFilter = "" And _
        Not FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetBocNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If BocNameFilter = "" And _
        FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetBocNameFilter = FiscalYearFilter
    End If
    If Not BocNameFilter = "" And _
        FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter = "" And _
        Not FundNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
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
    m_FundName = DLookup("FundCode", "Allocations", "Allocations.FundName =" _
        & "'" & Me.FundNameComboBox.Value & "'")
    Me.Filter = vbNullString
    FundNameFilter = vbNullString
    FundNameFilter = "[FundName] = '" _
        & Me.FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    Args.FundName = m_FundName
    Me.ObjectClassComboBox.RowSource = "SELECT DISTINCT Allocations.BocName" _
        & " FROM Allocations WHERE " _
        & Me.Filter
    SetComboBoxVisibility
    SetButtonVisibility
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
    If FundNameFilter = "" And _
        Not BocNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If FundNameFilter = "" And _
        BocNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetFundNameFilter = FiscalYearFilter
    End If
    If FundNameFilter = "" And _
        Not BocNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetFundNameFilter = BocNameFilter
    End If
    If Not FundNameFilter = "" And _
        Not BocNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & BocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = "" And _
        BocNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = "" And _
        Not BocNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & BocNameFilter
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



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetComboBoxVisibility
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxVisibility()
    On Error GoTo ErrorHandler:
    If Me.ObjectClassComboBox.Visible = False And _
        Me.ObjectClassGroupBox.Visible = False And _
        Me.ObjectClassComboBoxLabel.Visible = False Then
            Me.ObjectClassComboBox.Visible = Not Me.ObjectClassComboBox.Visible
            Me.ObjectClassComboBoxLabel.Visible = Not Me.ObjectClassComboBoxLabel.Visible
            Me.ObjectClassGroupBox.Visible = Not Me.ObjectClassGroupBox.Visible
    Else
            Me.ObjectClassComboBox.Visible = Not Me.ObjectClassComboBox.Visible
            Me.ObjectClassComboBoxLabel.Visible = Not Me.ObjectClassComboBoxLabel.Visible
            Me.ObjectClassGroupBox.Visible = Not Me.ObjectClassGroupBox.Visible
    End If
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
    Me!ObjectClassComboBox.ForeColor = RGB(255, 255, 255)
    Me!ObjectClassComboBox.BackColor = RGB(33, 33, 33)
    Me!ObjectClassComboBox.BorderColor = RGB(68, 114, 196)
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
    On Error GoTo ErrorHandler:
    If Not IsNull(Me.FiscalYearOptionGroup) And _
        Not Me.FiscalYearOptionGroup.Value = 2 Then
            Me.FiscalYearOptionGroup.Value = 2
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetButtonVisibility
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.ExitButton.SetFocus
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.AcceptButton.Visible = Not Me.AcceptButton.Visible
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
    On Error GoTo ErrorHandler:
    Me.ObjectClassComboBox.Value = vbNullString
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
    On Error GoTo ErrorHandler:
    m_BFY = vbNullString
    Args.BFY = vbNullString
    FiscalYearFilter = vbNullString
    m_BocName = vbNullString
    Args.BocName = vbNullString
    BocNameFilter = vbNullString
    m_FundName = vbNullString
    Args.FundName = vbNullString
    FundNameFilter = vbNullString
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
    On Error GoTo ErrorHandler:
    ClearFilterValues
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





