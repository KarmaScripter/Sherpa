VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgramArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_Code As String
Private m_Name As String
Private m_ProgramProjectCode As String
Private m_ProgramProjectName As String
Private m_ID As Integer
Private m_Error As String


'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ID(num As Integer)
    On Error GoTo ErrorHandler:
    If Not num < 0 Then
        m_ID = num
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ID() As Integer
    On Error GoTo ErrorHandler:
    If Not m_ID < 0 Then
        ID = m_ID
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Code(cd As String)
    On Error GoTo ErrorHandler:
    If Not cd & "" = "" Then
         m_Code = cd
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Code() As String
    On Error GoTo ErrorHandler:
    If Not m_Code & "" = "" Then
        Code = m_Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Name(nm As String)
    On Error GoTo ErrorHandler:
    If Not nm & "" = "" Then
         m_Name = nm
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
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
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ProgramProjectCode(pp As String)
    On Error GoTo ErrorHandler:
    If Not pp & "" = "" Then
         m_ProgramProjectCode = pp
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ProgramProjectCode() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramProjectCode & "" = "" Then
        ProgramProjectCode = m_ProgramProjectCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property






'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ProgramProjectName(ppn As String)
    On Error GoTo ErrorHandler:
    If Not ppn & "" = "" Then
         m_ProgramProjectName = ppn
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ProgramProjectName() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramProjectName & "" = "" Then
        ProgramProjectName = m_ProgramProjectName
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



