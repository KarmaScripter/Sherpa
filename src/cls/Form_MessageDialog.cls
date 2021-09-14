VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MessageDialog"
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
Private m_Error As String




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
    On Error GoTo ErrorHandler:
    If Not IsNull(Me.OpenArgs) Then
        Me.MessageBox.Value = ""
        Me.MessageBox.Value = CStr(Me.OpenArgs)
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let MessageArgs(pMessage As String)
    On Error GoTo ErrorHandler:
    If Not pMessage & "" = "" Then
        m_Message = pMessage
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get MessageArgs() As String
    On Error GoTo ErrorHandler:
    If Not m_Message & "" = "" Then
        MessageArgs = m_Message
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
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, Me.Name, acSaveNo
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





