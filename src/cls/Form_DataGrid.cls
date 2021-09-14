VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Args As AccountArgs
Private m_BudgetLevel As String
Private m_BFY As String
Private m_FundName As String
Private m_ProgramProjectName As String
Private pAnd As String
Private m_DataSet As DAO.Recordset
Private m_Error As String






'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BudgetLevelFilter() As String
    On Error GoTo ErrorHandler:
    If Not m_BudgetLevel & "" = "" Then
        BudgetLevelFilter = m_BudgetLevel
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FiscalYearFilter() As String
    On Error GoTo ErrorHandler:
    If Not m_BFY & "" = "" Then
        FiscalYearFilter = m_BFY
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property






'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FundNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not m_FundName & "" = "" Then
        FundNameFilter = m_FundName
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ProgramProjectNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramProjectName & "" = "" Then
        ProgramProjectNameFilter = m_ProgramProjectName
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(100, 100, 100)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.ProgramProjectNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.ProgramProjectNameComboBox.BackColor = RGB(100, 100, 100)
    Me.ProgramProjectNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilters()
    On Error GoTo ErrorHandler:
    m_BudgetLevel = ""
    m_BFY = ""
    m_FundName = ""
    m_ProgramProjectName = ""
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxes()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.Value = ""
    Me.ProgramProjectNameComboBox = ""
    ClearFilters
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
    On Error GoTo ErrorHandler:
    If Not IsNull(Me.BudgetLevelOptionGroup) And _
        Not Me.BudgetLevelOptionGroup.Value & "" = "" Then
        Me.BudgetLevelOptionGroup.Value = ""
    End If
    If Not IsNull(Me.FiscalYearOptionGroup) And _
        Not Me.FiscalYearOptionGroup.Value & "" = "" Then
        Me.FiscalYearOptionGroup.Value = ""
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BudgetLevelOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    m_BudgetLevel = vbNullString
    Select Case Me.BudgetLevelOptionGroup.Value
        Case 1
            Args.BudgetLevel = "[BudgetLevel] = '7'"
        Case 2
            Args.BudgetLevel = "[BudgetLevel] = '8'"
    End Select
    m_BudgetLevel = Args.BudgetLevel
    If Not m_BFY & "" = "" Then
            Me.FilterOn = True
            Me.Filter = m_BudgetLevel & pAnd & m_BFY
            Me.Requery
    End If
    If m_BFY & "" = "" Then
            Me.FilterOn = True
        Me.Filter = m_BudgetLevel
        Me.Requery
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    If Not m_BFY = "" Then
        m_BFY = ""
    End If
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            Args.BFY = "[BFY] IN ('2019', '2020')"
        Case 2
            Args.BFY = "[BFY] IN ('2021', '2022')"
    End Select
    m_BFY = Args.BFY
    If Not m_BudgetLevel = "" Then
            Me.FilterOn = True
            Me.Filter = m_BFY & pAnd & m_BudgetLevel
            Me.Requery
    End If
    If m_BudgetLevel & "" = "" Then
        Me.FilterOn = True
        Me.Filter = m_BFY
        Me.Requery
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    pAnd = " AND "
    SetComboBoxColors
    ResetOptionGroupValues
    ClearComboBoxes
    Set Args = New AccountArgs
    Me.FilterOn = True
    Set m_DataSet = Me.Recordset
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    Me.FilterOnLoad = True
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub EditButton_Click()
    On Error GoTo ErrorHandler:
    Dim DataSet As DAO.Database
    Set DataSet = CurrentDb
    Dim Account As DAO.Recordset
    Dim ID As Integer
    Me.CurrentIdTextBox.SetFocus
    ID = CInt(Me!CurrentIdTextBox.text)
    Me.RecordSource = "SELECT * FROM Allocations WHERE PrcId = " & ID
    Dim frm As Form
    DoCmd.OpenForm ("AccountEditor")
    Form_AccountEditor.RecordSource = Me.RecordSource
    DoCmd.Close acForm, "AllocationForm", acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    If Not m_FundName = "" Then
        m_FundName = ""
        ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT ProgramProjectName" _
            & " FROM Allocations WHERE " & m_FundName
    End If
    m_FundName = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    If Not m_BudgetLevel = "" And _
        m_BFY = "" Then
            Me.Filter = m_FundName & pAnd & m_BudgetLevel
            Me.Requery
    End If
    If Not m_BFY = "" And _
        m_BudgetLevel = "" Then
            Me.Filter = m_FundName & pAnd & m_BFY
            Me.Requery
    End If
    If Not m_BFY = "" And _
        Not m_BudgetLevel = "" Then
            Me.Filter = m_BFY & pAnd & m_BudgetLevel & pAnd & m_FundName
    End If
    If m_BFY = "" And _
        m_BudgetLevel = "" Then
            Me.Filter = m_FundName
            Me.Requery
    End If
    ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT ProgramProjectName" _
        & " FROM Allocations WHERE " & m_FundName
    Me.RecordSource = "SELECT * FROM Allocations WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProgramProjectNameComboBox_Change()
    On Error GoTo ErrorHandler:
    If Not m_ProgramProjectName = "" Then
        m_ProgramProjectName = ""
    End If
    m_ProgramProjectName = "[ProgramProjectName] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    If Not m_BudgetLevel = "" And _
        m_BFY = "" Then
            Me.Filter = m_ProgramProjectName & pAnd & m_BudgetLevel
            Me.Requery
    End If
    If Not m_BFY = "" And _
        m_BudgetLevel = "" Then
            Me.Filter = m_ProgramProjectName & pAnd & m_BFY
            Me.Requery
    End If
    If Not m_BFY = "" And _
        Not m_BudgetLevel = "" Then
            Me.Filter = m_ProgramProjectName & pAnd & m_BudgetLevel & pAnd & m_FundName
    End If
    If m_BFY = "" And _
        m_BudgetLevel = "" Then
            Me.Filter = m_FundName
    End If
    Me.Filter = m_ProgramProjectName
    Me.RecordSource = "SELECT * FROM Allocations WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearFilters
    ClearComboBoxes
    ResetOptionGroupValues
    Me.Filter = ""
    Me.RecordSource = "SELECT * FROM Allocations WHERE '" & Me.Filter & "'"
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





