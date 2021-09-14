VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FullTimeEquivalent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_RpioCode As String
Private m_BFY As String
Private m_BudgetLevel As String
Private m_FundCode As String
Private m_AhCode As String
Private m_OrgCode As String
Private m_AccountCode As String
Private m_ActivityCode As String
Private m_BocCode As String
Private m_RcCode As String
Private m_Amount As Double
Private m_Error As String



Public Sub Class_Initialize()
End Sub

Public Sub Class_Terminate()
End Sub






'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BFY() As String
    On Error GoTo ErrorHandler:
    If Not m_BFY & "" = "" Then
        BFY = m_BFY
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         BFY" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BFY(year As String)
    On Error GoTo ErrorHandler:
    If Not year & "" = "" Then
        m_BFY = year
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         BFY" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RPIO() As String
    On Error GoTo ErrorHandler:
    If Not m_RpioCode & "" = "" Then
        RPIO = m_RpioCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         RPIO" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let RPIO(rp As String)
    On Error GoTo ErrorHandler:
    If Not rp & "" = "" Then
        m_RpioCode = rp
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         RPIO" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BudgetLevel(bl As String)
    On Error GoTo ErrorHandler:
    If Not bl & "" = "" Then
        m_BudgetLevel = bl
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         BudgetLevel" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BudgetLevel() As String
    On Error GoTo ErrorHandler:
    If Not m_BudgetLevel & "" = "" Then
        BudgetLevel = m_BudgetLevel
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         BudgetLevel" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FundCode(fc As String)
    On Error GoTo ErrorHandler:
    If Not fc & "" = "" Then
        m_FundCode = fc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         FundCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FundCode() As String
    On Error GoTo ErrorHandler:
    If Not m_FundCode & "" = "" Then
        FundCode = m_FundCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         FundCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let AhCode(AH As String)
    On Error GoTo ErrorHandler:
    If Not AH & "" = "" Then
        m_AhCode = AH
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         AhCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get AhCode() As String
    On Error GoTo ErrorHandler:
    If Not m_AhCode & "" = "" Then
        AhCode = m_AhCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         AhCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let OrgCode(ORG As String)
    On Error GoTo ErrorHandler:
    If Not ORG & "" = "" Then
        m_OrgCode = ORG
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         OrgCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get OrgCode() As String
    On Error GoTo ErrorHandler:
    If Not m_OrgCode & "" = "" Then
        OrgCode = m_OrgCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         OrgCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let AccountCode(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_AccountCode = ac
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         AccountCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get AccountCode() As String
    On Error GoTo ErrorHandler:
    If Not m_AccountCode & "" = "" Then
        AccountCode = m_AccountCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         AccountCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let RcCode(rc As String)
    On Error GoTo ErrorHandler:
    If Not rc & "" = "" Then
        m_RcCode = rc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         RcCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RcCode() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        RcCode = m_RcCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         RcCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Amount(funds As Double)
    On Error GoTo ErrorHandler:
    If Not IsNull(funds) Or Not funds < 0 Then
        m_Amount = funds
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         Amount" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Amount() As Double
    On Error GoTo ErrorHandler:
    If Not IsNull(m_Amount) Or Not m_Amount < 0 Then
        Amount = m_Amount
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      FullTimeEquivalent" _
            & vbCrLf & "Member:         Amount" _
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



