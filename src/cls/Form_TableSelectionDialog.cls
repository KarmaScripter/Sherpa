VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TableSelectionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Args As AccountArgs
Public DataModelFilter As String
Public TableNameFilter As String
Private m_TableName As String
Private m_DataModel As String
Private pAnd As String
Private m_Error As String





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set Args = New AccountArgs
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Caption = "Data Selector"
    ClearComboBoxValues
    ClearFilterValues
    SetComboBoxColors
    pAnd = " AND "
    TableNameFilter = vbNullString
    Me.DataModelOptionGroup.Value = 1
    m_DataModel = "EXECUTION"
    DataModelFilter = "[Model] = '" & m_DataModel & "'"
    Me.Filter = DataModelFilter
    Me.TableNameComboBox.RowSource = "SELECT DISTINCT ApplicationTables.TableName" _
        & " FROM ApplicationTables" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub






'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DataModelOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DataModelOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me.TableNameComboBox.Value = vbNullString
    m_DataModel = vbNullString
    Me.Filter = vbNullString
    DataModelFilter = vbNullString
    Select Case Me.DataModelOptionGroup.Value
        Case 1
            m_DataModel = "EXECUTION"
            DataModelFilter = "[Model] = 'EXECUTION'"
            Me.Filter = DataModelFilter
            Me.TableNameComboBox.RowSource = "SELECT DISTINCT ApplicationTables.TableName" _
                & " FROM ApplicationTables" _
                & " WHERE " & Me.Filter
        Case 2
            m_DataModel = "REFERENCE"
            DataModelFilter = "[Model] = 'REFERENCE'"
            Me.Filter = DataModelFilter
            Me.TableNameComboBox.RowSource = "SELECT DISTINCT ApplicationTables.TableName" _
                & " FROM ApplicationTables" _
                & " WHERE " & Me.Filter
    End Select
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        TableNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub TableNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    TableNameFilter = vbNullString
    TableNameFilter = "[Name] = '" & Me.TableNameComboBox.Value & "'"
    m_TableName = Me.TableNameComboBox.Value
    Me.Filter = TableNameFilter
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



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearOptionGroupValues
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
    Me!TableNameComboBox.ForeColor = RGB(255, 255, 255)
    Me!TableNameComboBox.BackColor = RGB(33, 33, 33)
    Me!TableNameComboBox.BorderColor = RGB(68, 114, 196)
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
    Me.DataModelOptionGroup.Value = 1
    m_DataModel = "EXECUTION"
    DataModelFilter = "[Model] = 'EXECUTION'"
    Me.Filter = DataModelFilter
    Me.TableNameComboBox.RowSource = "SELECT DISTINCT ApplicationTables.TableName" _
        & " FROM ApplicationTables" _
        & " WHERE " & Me.Filter
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
    Me.TableNameComboBox.Value = vbNullString
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
    m_DataModel = vbNullString
    DataModelFilter = vbNullString
    m_TableName = vbNullString
    TableNameFilter = vbNullString
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







