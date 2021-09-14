VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SelectionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





Private m_BFY As String
Private m_BudgetLevel As String
Private m_FundCode As String
Private m_RcCode As String
Private m_DivisionName As String
Private m_AccountCode As String
Private m_BocCode As String
Public DialogArgs As BudgetArgs
Private m_OptionGroupFilter As String
Private m_FirstFilter As String
Private m_SecondFilter As String
Private m_DialogFilter As String
Private pAnd As String
Private m_Error As String






'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    SetComboBoxColors
    pAnd = " AND "
    m_BudgetLevel = "8"
    SetOptionGroupFilter
    Me.FirstComboBox.RowSource = "SELECT DISTINCT Divisions.Title FROM Divisions " _
        & "WHERE Divisions.Code NOT IN ('06', '06R')"
    Me.FirstComboBox.Value = ""
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
    Me!FirstComboBox.ForeColor = RGB(255, 255, 255)
    Me!FirstComboBox.BackColor = RGB(33, 33, 33)
    Me!FirstComboBox.BorderColor = RGB(68, 114, 196)
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
    If Not IsNull(Me.DialogOptionGroup) And _
        Not Me.DialogOptionGroup.Value & "" = "" Then
        Me.DialogOptionGroup.Value = ""
        m_OptionGroupFilter = ""
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetOptionGroupFilter
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetOptionGroupFilter()
    On Error GoTo ErrorHandler:
    m_OptionGroupFilter = ""
    m_BFY = ""
    Select Case Me.DialogOptionGroup.Value
        Case 1
            m_OptionGroupFilter = "[BFY] = '2020'"
            m_BFY = "2020"
        Case 2
            m_OptionGroupFilter = "[BFY] = '2021'"
            m_BFY = "2021"
    End Select
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DialogOptionGroup_AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DialogOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    m_OptionGroupFilter = ""
    m_BFY = ""
    DialogArgs.BFY = ""
    Select Case Me.DialogOptionGroup.Value
        Case 1
            m_OptionGroupFilter = "[BFY] = '2020'"
            m_BFY = "2020"
            DialogArgs.BFY = m_BFY
        Case 2
            m_OptionGroupFilter = "[BFY] = '2021'"
            m_BFY = "2021"
            DialogArgs.BFY = m_BFY
    End Select
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set DialogArgs = New BudgetArgs
    Me.FirstComboBox.Value = ""
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        RefreshButton_OnClick
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    m_FirstFilter = ""
    m_SecondFilter = ""
    Me.FirstComboBox.Value = ""
    Me.FirstComboBox.RowSource = "SELECT DISTINCT Divisions.Title FROM Divisions " _
        & "WHERE Divisions.Code NOT IN ('06', '06R')"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Property-Get
'   Name:        FiscalYear
'   Parameters:  Void
'   Purpose:
'----------------------------------------------------------------------------------
Public Property Get FiscalYear()
    On Error GoTo ErrorHandler:
    If Not m_BFY & "" = "" Then
        FiscalYear = m_BFY
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Property


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FirstFilterComboBox_Change
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FirstComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    m_FirstFilter = ""
    SetOptionGroupFilter
    m_FirstFilter = "[DivisionName] = '" & Me.FirstComboBox.SelText & "'"
    m_DivisionName = Me.FirstComboBox.SelText
    DialogArgs.DivisionName = m_DivisionName
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

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




