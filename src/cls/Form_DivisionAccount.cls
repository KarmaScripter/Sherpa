VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DivisionAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_Level As String
Private m_FiscalYear As String
Private m_FundCode As String
Private m_AhCode As String
Private m_OrgCode As String
Private m_BocCode As String
Private m_AccountCode As String
Private m_RcCode As String
Private m_Error As String
Private m_Notification As String






'----------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Level(pLevel As String)
    On Error GoTo ErrorHandler:
    If Not pLevel & "" = "" Then
        m_Level = pLevel
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Level() As String
    On Error GoTo ErrorHandler:
    If Not m_Level & "" = "" Then
        Level = m_Level
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FiscalYear(pBFY As String)
    On Error GoTo ErrorHandler:
    If Not pBFY & "" = "" Then
        m_FiscalYear = pBFY
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FiscalYear() As String
    On Error GoTo ErrorHandler:
    If m_FiscalYear & "" = "" Then
        FiscalYear = m_FiscalYear
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    Me.Filter = ""
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




