Attribute VB_Name = "MessageFactory"
Option Compare Database
Option Explicit


Private m_Message As String



'----------------------------------------------------------------------------------
'   Type        Property
'   Name        Message
'   Parameters  String
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Message(pText As String)
    If Not pText & "" = "" Then
        m_Message = pText
    End If
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name        Message
'   Parameters  Void
'   RetVal      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Message() As String
    If Not m_Message & "" = "" Then
        Message = m_Message
    End If
End Property




'----------------------------------------------------------------------------------
'   Type        Sub-Procedure
'   Name        ShowError
'   Parameters  String
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Sub ShowError(pMessage As String)
    If Not pMessage & "" = "" Then
        m_Message = pMessage
        DoCmd.OpenForm FormName:="ErrorDialog", _
            OpenArgs:=mMessage, WindowMode:=acDialog
    End If
End Sub




'----------------------------------------------------------------------------------
'   Type        Sub-Procedure
'   Name        ShowError
'   Parameters  String
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Sub ShowNotfication(pMessage As String)
    If Not pMessage & "" = "" Then
        m_Message = pMessage
        DoCmd.OpenForm FormName:="MessageDialog", _
            OpenArgs:=mMessage, WindowMode:=acDialog
    End If
End Sub


