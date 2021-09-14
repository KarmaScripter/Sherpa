VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Division"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_RcName As String
Private m_RcCode As String
Private m_Caption As String
Private m_Title As String
Private m_ID As Integer
Private m_Error As String



'----------------------------------------------------------------------------------
'   Type        Constructor
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub


'----------------------------------------------------------------------------------
'   Type        Destructor
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Class_Terminate()
End Sub



'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ID(bocid As Integer)
    On Error GoTo ErrorHandler:
    If Not bocid > 0 Then
        m_ID = bocid
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ID() As Integer
    On Error GoTo ErrorHandler:
    If Not m_ID > 0 Then
        ID = m_ID
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Code(cd As String)
    On Error GoTo ErrorHandler:
    If Not cd & "" = "" Then
        m_RcCode = cd
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Code() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        Code = m_RcCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Name(nm As String)
    On Error GoTo ErrorHandler:
    If Not nm & "" = "" Then
        m_RcName = nm
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Name() As String
    On Error GoTo ErrorHandler:
    If Not m_RcName & "" = "" Then
        Name = m_RcName
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Title(tl As String)
    On Error GoTo ErrorHandler:
    If Not tl & "" = "" Then
        m_Title = tl
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Title() As String
    On Error GoTo ErrorHandler:
    If Not m_Title & "" = "" Then
        Title = m_Title
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Caption(cp As String)
    On Error GoTo ErrorHandler:
    If Not cp & "" = "" Then
        m_Caption = cp
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Caption() As String
    On Error GoTo ErrorHandler:
    If Not m_Caption & "" = "" Then
        Caption = m_Caption
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



