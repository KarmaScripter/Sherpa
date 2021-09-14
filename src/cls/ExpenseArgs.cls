VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpenseArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private m_BFY As String
Private m_RpioCode As String
Private m_AhCode As String
Private m_FundCode As String
Private m_AccountCode As String
Private m_WorkCode As String
Private m_BocCode As String
Private m_RcCode As String
Private m_FocCode As String
Private m_FocName As String
Private m_Commitments As Double
Private m_OpenCommitments As Double
Private m_Obligations As Double
Private m_Deobligations As Double
Private m_ULO As Double
Private m_Expenditures As Double
Private m_Error As String



Public Sub Class_Initialize()
End Sub

Public Sub Class_Terminate()
End Sub




'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BFY(yr As String)
    On Error GoTo ErrorHandler:
    If Not yr & "" = "" Then
        m_BFY = yr
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


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
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let RpioCode(rp As String)
    On Error GoTo ErrorHandler:
    If Not rp & "" = "" Then
        m_RpioCode = rp
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RpioCode() As String
    On Error GoTo ErrorHandler:
    If Not m_RpioCode & "" = "" Then
        RpioCode = m_RpioCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let AhCode(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_AhCode = ac
    End If
ErrorHandler:
    ProcessError
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
    ProcessError
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
    ProcessError
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
    ProcessError
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
    ProcessError
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
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BocCode(bc As String)
    On Error GoTo ErrorHandler:
    If Not bc & "" = "" Then
        m_BocCode = bc
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BocCode() As String
    On Error GoTo ErrorHandler:
    If Not m_BocCode & "" = "" Then
        BocCode = m_BocCode
    End If
ErrorHandler:
    ProcessError
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
    ProcessError
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
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FocCode(fcc As String)
    On Error GoTo ErrorHandler:
    If Not fcc & "" = "" Then
        m_FocCode = fcc
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FocCode() As String
    On Error GoTo ErrorHandler:
    If Not m_FocCode & "" = "" Then
        FocCode = m_FocCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FocName(fcn As String)
    On Error GoTo ErrorHandler:
    If Not fcn & "" = "" Then
        m_FocName = fcn
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FocName() As String
    On Error GoTo ErrorHandler:
    If Not m_FocName & "" = "" Then
        FocName = m_FocName
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        Commitments
'   Parameters:  String
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Commitments(cm As Double)
    On Error GoTo ErrorHandler:
    If Not cm & "" = "" Then
        m_Commitments = cm
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        Commitments
'   Parameters:  Void
'   Retval:      Double
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Commitments() As Double
    On Error GoTo ErrorHandler:
    If Not m_Commitments & "" = "" Then
        Commitments = m_Commitments
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property

'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        OpenCommitments
'   Parameters:  Double
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let OpenCommitments(oc As Double)
    On Error GoTo ErrorHandler:
    If Not oc & "" = "" Then
        m_OpenCommitments = oc
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        OpenCommitments
'   Parameters:  Void
'   Retval:      Double
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get OpenCommitments() As Double
    On Error GoTo ErrorHandler:
    If Not m_Commitments & "" = "" Then
        Commitments = m_Commitments
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  Double
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Obligations(ob As Double)
    On Error GoTo ErrorHandler:
    If Not ob & "" = "" Then
        m_Obligations = ob
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        Obligations
'   Parameters:  Void
'   Retval:      Double
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Obligations() As Double
    On Error GoTo ErrorHandler:
    If Not m_Obligations & "" = "" Then
        Obligations = m_Obligations
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        Deobligations
'   Parameters:  Double
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Deobligations(od As Double)
    On Error GoTo ErrorHandler:
    If Not od & "" = "" Then
        m_Deobligations = od
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        Deobligations
'   Parameters:  Void
'   Retval:      Double
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Deobligations() As Double
    On Error GoTo ErrorHandler:
    If Not m_Deobligations & "" = "" Then
        Deobligations = m_Deobligations
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:
'   Parameters:  Double
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ULO(uo As Double)
    On Error GoTo ErrorHandler:
    If Not uo & "" = "" Then
        m_ULO = uo
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        ULO
'   Parameters:  Void
'   Retval:      Double
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ULO() As Double
    On Error GoTo ErrorHandler:
    If Not m_ULO & "" = "" Then
        ULO = m_ULO
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure Let
'   Name:        Expenditures
'   Parameters:  Double
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Expenditure(ex As Double)
    On Error GoTo ErrorHandler:
    If Not ex & "" = "" Then
        m_Expenditures = ex
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Procedure Get
'   Name:        Expenditures
'   Parameters:  Void
'   Retval:      Double
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Expenditures() As Double
    On Error GoTo ErrorHandler:
    If Not m_Expenditures & "" = "" Then
        Expenditures = m_Expenditures
    End If
ErrorHandler:
    ProcessError
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



