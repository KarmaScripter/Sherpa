VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Error As String
Private m_Sender As String
Private m_Reciever As String
Private m_Subject As String
Private m_Body As String






'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  String
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Sender(pAddress As String)
    On Error GoTo ErrorHandler:
    If Not pAddress & "" = "" Then
        m_Sender = pAddress
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Sender() As String
    On Error GoTo ErrorHandler:
    If Not m_Sender & "" = "" Then
        Sender = m_Sender
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  String
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Reciever(pAddress As String)
    On Error GoTo ErrorHandler:
    If Not pAddress & "" = "" Then
        m_Reciever = pAddress
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Reciever() As String
    On Error GoTo ErrorHandler:
    If Not m_Reciever & "" = "" Then
        Reciever = m_Reciever
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  String
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Subject(pSubject As String)
    On Error GoTo ErrorHandler:
    If Not pSubject & "" = "" Then
        m_Subject = pSubject
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Subject() As String
    On Error GoTo ErrorHandler:
    If Not m_Subject & "" = "" Then
        Subject = m_Subject
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  String
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Body(pMessage As String)
    On Error GoTo ErrorHandler:
    If Not pMessage & "" = "" Then
        m_Body = pMessage
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Body() As String
    On Error GoTo ErrorHandler:
    If Not m_Body & "" = "" Then
        Body = m_Body
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


