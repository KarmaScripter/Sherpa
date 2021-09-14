VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_DataAccess As DataAccess
Private m_Base As DAO.Recordset
Private m_Values()  As Variant
Private m_Names() As String
Private m_Type() As String
Private m_Error As String
Private m_Notification As String





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
ErrorHandler:
    ProcessError
    Exit Property
End Property


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



