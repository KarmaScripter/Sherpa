VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NationalProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private m_Code As String
Private m_Name As String
Private m_Title As String
Private m_RPIO As String
Private m_NpmId As Integer
Private m_Error As String

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
    On Error GoTo ErrorHandler:
    If num > 0 Then
        m_NpmId = num
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      ID" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Not m_NpmId < 0 Then
        ID = m_NpmId
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      ID" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
        m_Code = cd
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      Code" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Not m_Code & "" = "" Then
        Code = m_Code
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      Code" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
        m_Name = nm
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      Name" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Not m_Name & "" = "" Then
        Name = m_Name
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      Name" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let RPIO(rp As String)
    On Error GoTo ErrorHandler:
    If Not rp & "" = "" Then
        m_RPIO = rp
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      RPIO" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get RPIO() As String
    On Error GoTo ErrorHandler:
    If Not m_RPIO & "" = "" Then
        RPIO = m_RPIO
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      RPIO" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      Title" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Err.Number > 0 Then
        m_Error = "Source:   NationalProgram" _
            & vbCrLf & "Member:      Title" _
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



