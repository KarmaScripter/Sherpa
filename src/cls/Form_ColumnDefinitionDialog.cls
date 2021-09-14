VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ColumnDefinitionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Args As AccountArgs
Private m_DataColumns As Scripting.Dictionary
Private m_Name As String
Private m_DataType As String
Private m_Database As String
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
    Me.DatabaseComboBox.RowSource = "SELECT DISTINCT SchemaTypes.Database FROM SchemaTypes;"
    Me.DataTypeComboBox.RowSource = vbNullString
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
    Me.Caption = "Budget Execution"
    SetComboBoxColors
    pAnd = " AND "
    ClearComboBoxValues
    Set m_DataColumns = CreateObject("Scripting.Dictionary")
ErrorHandler:
    ProcessError
    Exit Sub
End Sub






'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DataTypeComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DatabaseComboBox_Change()
    On Error GoTo ErrorHandler:
    m_Database = vbNullString
    m_Database = Me.DatabaseComboBox.Value
    Me.DataTypeComboBox.RowSource = "SELECT DISTINCT SchemaTypes.TypeName" _
        & " FROM SchemaTypes" _
        & " WHERE [Database] = '" & m_Database & "'"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DataTypeComboBox_Change()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DataTypeComboBox_Change()
    On Error GoTo ErrorHandler:
    m_DataType = vbNullString
    m_DataType = Me.DataTypeComboBox.Value
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        GetSchemaData()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetSchemaData() As Scripting.Dictionary
    On Error GoTo ErrorHandler:
    If Not Me.ColumnNameTextBox.Value & "" = "" And _
        Not m_DataType & "" = "" Then
            m_Name = vbNullString
            m_Name = Me.ColumnNameTextBox.text
            m_DataColumns.add Key:="Name", Item:=m_Name
            m_DataColumns.add Key:="Type", Item:=m_DataType
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
Private Sub AddButton_Click()
    On Error GoTo ErrorHandler:
    If Not Me.ColumnNameTextBox.Value & "" = "" And _
        Not m_DataType & "" = "" Then
            Dim p_DataColumn As Scripting.Dictionary
            Set p_DataColumn = CreateObject("Scripting.Dictionary")
            Set p_DataColumn = GetSchemaData
    End If
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
    Me.ColumnNameTextBox.Value = "Enter Column Name"
    Me.DataTypeComboBox.Value = vbNullString
    Me.DatabaseComboBox.Value = vbNullString
    Me.DataTypeComboBox.RowSource = vbNullString
    Me.DatabaseComboBox.RowSource = "SELECT DISTINCT SchemaTypes.Database FROM SchemaTypes;"
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
    Me!ColumnNameTextBox.ForeColor = RGB(255, 255, 255)
    Me!ColumnNameTextBox.BackColor = RGB(33, 33, 33)
    Me!ColumnNameTextBox.BorderColor = RGB(68, 114, 196)
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
    Me.ColumnNameTextBox.Value = "Enter Column Name"
    Me.DataTypeComboBox.Value = vbNullString
    Me.DatabaseComboBox.Value = vbNullString
    Me.DataTypeComboBox.RowSource = vbNullString
    Me.DatabaseComboBox.RowSource = vbNullString
    Me.DatabaseComboBox.RowSource = "SELECT DISTINCT SchemaTypes.Database FROM SchemaTypes;"
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
Private Sub ColumnNameTextBox_Enter()
    On Error GoTo ErrorHandler:
    Me.ColumnNameTextBox.Value = vbNullString
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







