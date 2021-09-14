VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Appropriation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





Private m_FundCode As String
Private m_FundName As String
Private m_FundTitle As String
Private m_TreasurySymbol As String
Private m_ID As Integer
Private m_Error As String



'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ID(num As Integer)
    On Error GoTo ErrorHandler:
    If num > 0 Then
        m_ID = num
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       ID" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ID() As Integer
    On Error GoTo ErrorHandler:
    If m_ID > 0 Then
        ID = m_ID
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       ID" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Code(cd As String)
    On Error GoTo ErrorHandler:
    If Not cd & "" = "" Then
        m_FundCode = cd
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       Code" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Code() As String
    On Error GoTo ErrorHandler:
    If Not m_FundCode & "" = "" Then
        Code = m_FundCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       Code" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Name(nm As String)
    On Error GoTo ErrorHandler:
    If Not nm & "" = "" Then
        m_FundName = nm
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       Name" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Name() As String
    On Error GoTo ErrorHandler:
    If Not m_FundName & "" = "" Then
        Name = m_FundName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       Name" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Title(tit As String)
    On Error GoTo ErrorHandler:
    If Not tit & "" = "" Then
        m_FundTitle = tit
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       Title" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Title() As String
    On Error GoTo ErrorHandler:
    If Not m_FundTitle & "" = "" Then
        Title = m_FundTitle
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       Title" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let TreasurySymbol(ts As String)
    On Error GoTo ErrorHandler:
    If Not ts & "" = "" Then
        m_TreasurySymbol = ts
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       TreasurySymbol" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get TreasurySymbol() As String
    On Error GoTo ErrorHandler:
    If Not m_TreasurySymbol & "" = "" Then
        TreasurySymbol = m_TreasurySymbol
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      Appropriation" _
            & vbCrLf & "Member:       TreasurySymbol" _
            & vbCrLf & "Descript:   " & Err.Description
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



