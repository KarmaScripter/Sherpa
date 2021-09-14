VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_DataAccess As DataAccess
Private m_Base As DAO.Recordset
Private m_Name As String
Private m_DataSchema As DataSchema
Private m_DataColumns()  As DataColumn
Private m_DataRows() As DataRow
Private m_Error As String
Private m_Notification As String






'----------------------------------------------------------------------------------
'   Type:        Constructor/Finalizer
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_DataAccess = New DataAccess
End Sub

Private Sub Class_Terminate()
    KillData
End Sub




'----------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Base(pBase As DAO.Recordset)
    On Error GoTo ErrorHandler:
    If Not pBase Is Nothing Then
        Set m_Base = pBase
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        DAO.Field
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Base() As DAO.Recordset
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Set Base = m_Base
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:        Name
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Name(pName As String)
    On Error GoTo ErrorHandler:
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        Name
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Name() As String
    On Error GoTo ErrorHandler:
    If Not m_Name & "" = "" Then
        Name = m_Name
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        RowCount
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RowCount() As Integer
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        RowCount = m_Base.RecordCount
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        ColumnCount
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ColumnCount() As Integer
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        ColumnCount = m_Base.Fields.count
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property








'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        Columns
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Columns() As DataColumn()
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Columns = GetDataColumns(m_Base.Fields)
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        Columns
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Rows() As DataRow()
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Rows = GetDataRows(m_Base.Fields)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataTable" _
            & vbCrLf & "Member:     Rows" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property







'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        Columns
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Schema() As DataSchema
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Set m_DataSchema = New DataSchema
        m_DataSchema.Base = m_Base
        Schema = m_DataSchema
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:
'   Parameters:  DAO.Fields
'   RetVal:      Variant()
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetDataColumns(pData As DAO.Fields) As DataColumn()
    On Error GoTo ErrorHandler:
    If Not pData Is Nothing Then
        Dim p_DataColumns() As New DataColumn
        Dim i As Integer
        For i = 0 To i = pData.count
            Dim p_Column As DataColumn
            Set p_Column = New DataColumn
            p_Column.Base = pData(i)
            p_DataColumns(i) = p_Column
        Next i
        GetDataColumns = p_DataColumns
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function





'--------------------------------------------------------------------------------
'   Type:        Function
'   Name:
'   Parameters:  DAO.Fields
'   RetVal:      Variant()
'   Purpose:
'--------------------------------------------------------------------------------
Private Function GetDataRows(pData As DAO.Recordset) As DataRow()
    On Error GoTo ErrorHandler:
    If Not pData Is Nothing Then
        Dim p_DataRows() As New DataRow
        Dim i As Integer
        For i = 0 To i = pData.RecordCount
            Dim p_Row As DataRow
            Set p_Row = New DataRow
            p_Row.Base = pData(i)
            p_DataRows(i) = p_Row
        Next i
        GetDataRows = p_DataRows
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        IsConnected
'   Parameters:  Void
'   RetVal:      Boolean
'   Purpose:     Function returns boolean value indicating if Workspace is open
'---------------------------------------------------------------------------------
Public Property Get IsConnected() As Boolean
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        IsConnected = True
    End If
    If m_Base Is Nothing Then
        IsConnected = False
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'--------------------------------------------------------------------------------
'   Type:        Function
'   Name:        KillData()
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:     Releases resources used by the Database
'--------------------------------------------------------------------------------
Private Sub KillData()
    On Error GoTo ErrorHandler:
    If Not m_DataAccess Is Nothing Then
        Set m_DataAccess = Nothing
    End If
    If Not m_Base Is Nothing Then
        Set m_Base = Nothing
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



