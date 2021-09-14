VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_DataAccess As DataAccess
Private m_Base As Record
Private m_Fields As DAO.Fields
Private m_Row As Record
Private m_Data As Scripting.Dictionary
Private m_Values()  As Variant
Private m_Names() As String
Private m_Index As DataColumn
Private m_DataColumns() As DataColumn
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
    If Not m_DataAccess Is Nothing Then
        Set m_DataAccess = Nothing
    End If
    If Not m_Base Is Nothing Then
        Set m_Base = Nothing
    End If
    If Not m_Row Is Nothing Then
        Set m_Row = Nothing
    End If
End Sub




'----------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Base(pBase As Record)
    On Error GoTo ErrorHandler:
    If Not pBase Is Nothing Then
        Set m_Base = pBase(0)
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Base() As Record
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Set Base = m_Base
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Fields() As DAO.Fields
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Set Fields = m_Base.Fields
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Values() As Variant()
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Dim p_Values() As Variant
        Dim i As Integer
        For i = 0 To i = m_Base.Fields.count
            p_Values(i) = m_Base.Fields(i).Value
        Next i
        Values = p_Values
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Index() As DataColumn
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Set Index = GetIndex(m_Base.Fields)
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
Private Function GetValues(pData As DAO.Fields) As Variant()
    On Error GoTo ErrorHandler:
    If Not pData Is Nothing Then
        Dim p_Values() As Variant
        Dim i As Integer
        For i = 0 To i = m_Base.Fields.count
            p_Values(i) = m_Base.Fields(i).Value
        Next i
        GetValues = p_Values
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:
'   Parameters:  DAO.Fields
'   RetVal:      Variant()
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetNames(pData As DAO.Fields) As String()
    On Error GoTo ErrorHandler:
    If Not pData Is Nothing Then
        Dim p_Names() As String
        Dim i As Integer
        For i = 0 To i = m_Base.Fields.count
            p_Names(i) = m_Base.Fields(i).Name
        Next i
        GetNames = p_Names
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function





'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:
'   Parameters:  DAO.Fields
'   RetVal:      Variant()
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetIndex(pData As DAO.Fields) As DataColumn
    On Error GoTo ErrorHandler:
    If Not pData Is Nothing Then
        Dim p_Index As DataColumn
        Set p_Index = New DataColumn
        p_Index.Base = pData(0)
        Set GetIndex = p_Index
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




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



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        IsConnected
'   Parameters:  Void
'   RetVal:      Boolean
'   Purpose:     Function returns boolean value indicating if Recordset is valid
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






