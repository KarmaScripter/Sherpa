VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RegionalAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





Private p_Level As String
Private p_FiscalYear As String
Private p_FundCode As String
Private p_AhCode As String
Private p_OrgCode As String
Private p_BocCode As String
Private p_AccountCode As String
Private p_RcCode As String
Private m_Error As String








'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Level(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        p_Level = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property








'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Level() As String
    On Error GoTo ErrorHandler:
    If (p_Level <> "" And Not IsNull(p_Level)) Then
        Level = p_Level
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property








'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FiscalYear(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        p_FiscalYear = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FiscalYear() As String
    On Error GoTo ErrorHandler:
    If (p_FiscalYear <> "" And Not IsNull(p_FiscalYear)) Then
        FiscalYear = p_FiscalYear
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property









'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
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
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RegionalAmountTextBox_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
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





