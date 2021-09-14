VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmployeeArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_RcCode As String
Private m_EmployeeNumber As String
Private m_LastName As String
Private m_FirstName As String
Private m_PayGrade As String
Private m_Step As String
Private m_JobTitle As String
Private m_Error As String

Public Sub Class_Initialize()
End Sub

Public Sub Class_Terminate()
End Sub





'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        EmployeeNumber
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let EmployeeNumber(eid As String)
    On Error GoTo ErrorHandler:
    If eid & "" = "" Then
        m_EmployeeNumber = eid
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        EmployeeNumber
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get EmployeeNumber() As String
    On Error GoTo ErrorHandler:
    If Not m_EmployeeNumber & "" = "" Then
        EmployeeNumber = m_EmployeeNumber
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        LastName
'   Parameters:  String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let LastName(ln As String)
    On Error GoTo ErrorHandler:
    If ln & "" = "" Then
        m_LastName = ln
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        LastName
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get LastName() As String
    On Error GoTo ErrorHandler:
    If Not m_LastName & "" = "" Then
        LastName = m_LastName
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        FirstName
'   Parameters:  String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FirstName(fn As String)
    On Error GoTo ErrorHandler:
    If fn & "" = "" Then
    m_FirstName = fn
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        FirstName
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FirstName() As String
    On Error GoTo ErrorHandler:
    If Not m_FirstName & "" = "" Then
        FirstName = m_FirstName
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        RcCode
'   Parameters:  String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let RcCode(rc As String)
    On Error GoTo ErrorHandler:
    If Not rc & "" = "" Then
        m_RcCode = rc
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        RcCode
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RcCode() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        RcCode = m_RcCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        PayGrade
'   Parameters:  String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let PayGrade(pg As String)
    On Error GoTo ErrorHandler:
    If Not pg & "" = "" Then
        m_PayGrade = pg
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        PayGrade
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get PayGrade() As String
    On Error GoTo ErrorHandler:
    If Not m_PayGrade & "" = "" Then
        PayGrade = m_PayGrade
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        JobTitle
'   Parameters:  String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let JobTitle(jt As String)
    On Error GoTo ErrorHandler:
    If Not jt & "" = "" Then
        m_JobTitle = jt
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        JobTitle
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get JobTitle() As String
    On Error GoTo ErrorHandler:
    If Not m_JobTitle & "" = "" Then
        JobTitle = m_JobTitle
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



