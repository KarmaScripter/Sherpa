VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private m_ID As Integer
Private m_Code As String
Private m_GoalCode As String
Private m_ObjectiveCode As String
Private m_ProgramProjectCode As String
Private m_ActivityCode As String
Private m_Error As String



'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name            ID
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ID(num As Integer)
    On Error GoTo ErrorHandler:
    If Not num < 0 Then
        m_ID = num
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     ID" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name            ID
'   Parameters      Void
'   Return          Integer
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ID() As Integer
    On Error GoTo ErrorHandler:
    If Not m_ID < 0 Then
        ID = m_ID
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     ID" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name            Code
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Code(cd As String)
    On Error GoTo ErrorHandler:
    If Not cd & "" = "" Then
        m_Code = cd
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     Code" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name            Code
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Code() As String
    On Error GoTo ErrorHandler:
    If Not m_Code & "" = "" Then
        Code = m_Code
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     Code" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name            Goal
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Goal(gc As String)
    On Error GoTo ErrorHandler:
    If Not gc & "" = "" Then
        m_GoalCode = gc
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     Goal" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name            Goal
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Goal() As String
    On Error GoTo ErrorHandler:
    If Not m_GoalCode & "" = "" Then
        Goal = m_GoalCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     Goal" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name            Objective
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Objective(oc As String)
    On Error GoTo ErrorHandler:
    If Not oc & "" = "" Then
        m_ObjectiveCode = oc
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     Objective" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name            Objective
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Objective() As String
    On Error GoTo ErrorHandler:
    If Not m_ObjectiveCode & "" = "" Then
        Objective = m_ObjectiveCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     Objective" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name            ProgramProjectCode
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ProgramProjectCode(ppc As String)
    On Error GoTo ErrorHandler:
    If Not ppc & "" = "" Then
        m_ProgramProjectCode = ppc
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     ProgramProjectCode" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name            ProgramProjectCode
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ProgramProjectCode() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramProjectCode & "" = "" Then
        ProgramProjectCode = m_ProgramProjectCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:       Account" _
            & vbCrLf & "Member:     ProgramProjectCode" _
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



