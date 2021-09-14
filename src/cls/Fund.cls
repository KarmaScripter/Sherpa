VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private m_Code As String
Private m_Name As String
Private m_ID As Integer
Private m_Title As String
Private m_Symbol As String
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
    If Not num < 0 Then
        m_ID = num
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Fund" _
            & vbCrLf & "Member:     ID" _
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
    If Not m_ID > 0 Then
        ID = m_ID
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Fund" _
            & vbCrLf & "Member:     ID" _
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
        m_Error = "Source:      Fund" _
            & vbCrLf & "Member:     Code" _
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
        m_Error = "Source:      Fund" _
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
Public Property Let Name(nm As String)
    On Error GoTo ErrorHandler:
    If Not nm & "" = "" Then
        m_Name = nm
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Fund" _
            & vbCrLf & "Member:     Name" _
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
        m_Error = "Source:      Fund" _
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
Public Property Let Title(tl As String)
    On Error GoTo ErrorHandler:
    If Not tl & "" = "" Then
        m_Title = tl
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Fund" _
            & vbCrLf & "Member:     Title" _
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
        m_Error = "Source:      Fund" _
            & vbCrLf & "Member:     Title" _
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
Public Property Let TreasurySymbol(ts As String)
    On Error GoTo ErrorHandler:
    If Not ts & "" = "" Then
        m_Symbol = ts
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Fund" _
            & vbCrLf & "Member:     TreasurySymbol" _
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
Public Property Get TreasurySymbol() As String
    On Error GoTo ErrorHandler:
    If Not m_Symbol & "" = "" Then
        TreasurySymbol = m_Symbol
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Fund" _
            & vbCrLf & "Member:     TreasurySymbol" _
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



