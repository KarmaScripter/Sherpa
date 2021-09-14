VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ErrorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_Message As String
Private m_Title As String
Private m_Number As Integer
Private m_Description As String
Private m_ErrorArgs As Variant
Private m_MessageArgs As Variant


'----------------------------------------------------------------------------------
'   Type:        Constructor
'   Name:        Class_Initialize
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub










'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    If Not IsNull(Me.OpenArgs) Then
        Me.ErrorTextBox.Value = ""
        Me.ErrorTextBox.Value = CStr(Me.OpenArgs)
    End If
End Sub






'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
End Sub




'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:        Text
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let text(pMessage As String)
    If Not pMessage & "" = "" Then
        m_Message = pMessage
    End If
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:        Text
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get text() As String
    If Not m_Message & "" = "" Then
        text = m_Message
    End If
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FaultCode(pNumber As Integer)
    If pNumber > 0 Then
        m_Number = pNumber
    End If
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FaultCode() As Integer
    If Not m_Number > 0 Then
        FaultCode = m_Number
    End If
End Property





'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Description(pDescription As String)
    If Not pDescription & "" = "" Then
        m_Description = pDescription
    End If
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Description() As String
    If Not m_Description & "" = "" Then
        Description = m_Description
    End If
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ErrorArgs(pError As Variant)
    If Not IsNull(pError) Then
        m_ErrorArgs = pError
    End If
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ErrorArgs() As Variant
    If Not IsNull(m_ErrorArgs) Then
        ErrorArgs = m_ErrorArgs
        text = CStr(ErrorArgs)
    End If
End Property





'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    DoCmd.Close acForm, Me.Name
End Sub
