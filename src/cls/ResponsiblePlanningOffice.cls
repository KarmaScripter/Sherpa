Option Compare Database
Option Explicit




Private mCode As String
Private mName As String
Private mID As Integer



Public Sub Class_Initialize()
End Sub

Public Sub Class_Terminate()
End Sub



'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ID(num As Integer)
    If Not num < 0 Then
        mID = num
    End If
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ID() As Integer
    If Not mID < 0 Then
        ID = mID
    End If
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Code(cd As String)
    If Not cd & "" = "" Then
        mCode = cd
    End If
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Code() As String
    If Not mCode & "" = "" Then
        Code = mCode
    End If
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Name(nm As String)
    If Not nm & "" = "" Then
        mName = nm
    End If
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Name() As String
    If Not mName & "" = "" Then
        Name = mName
    End If
End Property


