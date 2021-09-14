Option Compare Database
Option Explicit



Private mMessage As String
Private mTitle As String
Private mNumber As Integer
Private mDescription As String
Private mErrorArgs As Variant



'----------------------------------------------------------------------------------
'   Type:        Constructor
'   Name:        Class_Initialize
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub


'----------------------------------------------------------------------------------
'   Type:        Constructor
'   Name:        Class_Initialize
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Terminate()
End Sub



'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let text(msg As String)
    If Not msg & "" = "" Then
        mMessage = msg
    End If
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get text() As String
    If Not mMessage & "" = "" Then
        text = mMessage
    End If
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Fault(en As Integer)
    If en > 0 Then
        mNumber = en
    End If
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Fault() As Integer
    If Not mNumber > 0 Then
        Fault = mNumber
    End If
End Property





'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Description(e As String)
    If Not e & "" = "" Then
        mDescription = e
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
    If Not mDescription & "" = "" Then
        Description = mDescription
    End If
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ErrorArgs(oa As Variant)
    If Not IsNull(oa) Then
        mErrorArgs = oa
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
    If Not IsNull(mErrorArgs) Then
        ErrorArgs = mErrorArgs
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
Private Sub CloseButton_Click()
    DoCmd.Close acForm, Me.Name
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
