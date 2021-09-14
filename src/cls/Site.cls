VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Site"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private m_Code As String
Private m_Name As String
Private m_Identifier As String
Private m_Address As String
Private m_State As String
Private m_City As String
Private m_ID As Integer
Private m_Error As String




Public Sub Class_Initialize()
End Sub

Public Sub Class_Terminate()
End Sub





'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ID(intId As Integer)
    On Error GoTo ErrorHandler:
    If intId > -1 Then
        m_ID = intId
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     ID" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Not m_ID > -1 Then
        ID = m_ID
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     ID" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     Code" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     Code" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Name(OrgName As String)
    On Error GoTo ErrorHandler:
    If Not OrgName & "" = "" Then
        m_Name = OrgName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     Name" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     Name" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property






'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Identifier(pIdent As String)
    On Error GoTo ErrorHandler:
    If Not pIdent & "" = "" Then
        m_Identifier = pIdent
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     Identifier" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name        Identifier
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Identifier() As String
    On Error GoTo ErrorHandler:
    If Not m_Identifier & "" = "" Then
        Identifier = m_Identifier
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     Identifier" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property






'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let State(pState As String)
    On Error GoTo ErrorHandler:
    If Not pState & "" = "" Then
        m_State = pState
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     State" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name        State
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get State() As String
    On Error GoTo ErrorHandler:
    If Not m_State & "" = "" Then
        State = m_State
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     State" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property







'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let City(pCity As String)
    On Error GoTo ErrorHandler:
    If Not pCity & "" = "" Then
        m_City = pCity
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     City" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name        State
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get City() As String
    On Error GoTo ErrorHandler:
    If Not m_City & "" = "" Then
        City = m_City
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     City" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Address(pAddress As String)
    On Error GoTo ErrorHandler:
    If Not pAddress & "" = "" Then
        m_Address = pAddress
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     Address" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name        State
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Address() As String
    On Error GoTo ErrorHandler:
    If Not m_Address & "" = "" Then
        Address = m_Address
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   Site" _
            & vbCrLf & "Member:     Address" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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



