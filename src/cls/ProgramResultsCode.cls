VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgramResultsCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_BFY As BudgetFiscalYear
Private m_Account As Account
Private m_Fund As Fund
Private m_RPIO As ResourcePlanningImplementationOffice
Private m_RC As ResponsibilityCenter
Private m_AH As AllowanceHolder
Private m_Org As Organization
Private m_Activity As Activity
Private m_ProgramProject As ProgramProject
Private m_ProgramArea As ProgramArea
Private m_Amount As Double
Private m_Error As String






'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  BudgetFiscalYear
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BFY(pBFY As BudgetFiscalYear)
    On Error GoTo ErrorHandler:
    If Not pBFY Is Nothing Then
        Set m_BFY = New BudgetFiscalYear
        m_BFY.Base = pBFY.Base
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      BudgetFiscalYear
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BFY() As BudgetFiscalYear
    On Error GoTo ErrorHandler:
    If Not m_BFY Is Nothing Then
        BFY = m_BFY
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Fund
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Fund(pFund As Fund)
    On Error GoTo ErrorHandler:
    If Not pFund Is Nothing Then
        Set m_Fund = New Fund
        m_Fund = pFund
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Account
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Fund() As Fund
    On Error GoTo ErrorHandler:
    If Not m_Fund Is Nothing Then
        Fund = m_Fund
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        RPIO
'   Parameters:  ResourcePlanningImplementationOffice
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let RPIO(pRPIO As ResourcePlanningImplementationOffice)
    On Error GoTo ErrorHandler:
    If Not pRPIO Is Nothing Then
        Set m_RPIO = New ResourcePlanningImplementationOffice
        m_RPIO = pRPIO
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        RPIO
'   Parameters:  Void
'   RetVal:      ResourcePlanningImplementationOffice
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RPIO() As ResourcePlanningImplementationOffice
    On Error GoTo ErrorHandler:
    If Not m_RPIO Is Nothing Then
        RPIO = m_RPIO
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        AH
'   Parameters:  AllowanceHolder
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let AH(pAH As AllowanceHolder)
    On Error GoTo ErrorHandler:
    If Not pAH Is Nothing Then
        Set m_AH = New AllowanceHolder
        m_AH = pAH
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        AH
'   Parameters:  Void
'   RetVal:      AllowancHolder
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get AH() As AllowanceHolder
    On Error GoTo ErrorHandler:
    If Not m_AH Is Nothing Then
        AH = m_AH
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        Account
'   Parameters:  Account
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Account(pAccount As Account)
    On Error GoTo ErrorHandler:
    If Not pAccount Is Nothing Then
        Set m_Account = New Account
        m_Account = pAccount
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        Acciybt
'   Parameters:  Void
'   RetVal:      Account
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Account() As Account
    On Error GoTo ErrorHandler:
    If Not m_Account Is Nothing Then
        Account = m_Account
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        RC
'   Parameters:  ResponsibilityCenter
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let RC(pRC As ResponsibilityCenter)
    On Error GoTo ErrorHandler:
    If Not pRC Is Nothing Then
        Set m_RC = New ResponsibilityCenter
        m_RC = pRC
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        RC
'   Parameters:  Void
'   RetVal:      ResponsibilityCenter
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RC() As ResponsibilityCenter
    On Error GoTo ErrorHandler:
    If Not m_RC Is Nothing Then
        RC = m_RC
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        ORG
'   Parameters:  Organization
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ORG(pORG As Organization)
    On Error GoTo ErrorHandler:
    If Not pORG Is Nothing Then
        Set m_Org = New Organization
        m_Org = pORG
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        ORG
'   Parameters:  Void
'   RetVal:      Organization
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ORG() As Organization
    On Error GoTo ErrorHandler:
    If Not m_Org Is Nothing Then
        ORG = m_Org
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        Activity
'   Parameters:  Activity
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Activity(pActivity As Activity)
    On Error GoTo ErrorHandler:
    If Not pActivity Is Nothing Then
        Set m_Activity = New Activity
        m_Activity = pActivity
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        Activity
'   Parameters:  Void
'   RetVal:      Activity
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Activity() As Activity
    On Error GoTo ErrorHandler:
    If Not m_Activity Is Nothing Then
        Activity = m_Activity
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        ProgramProject
'   Parameters:  ProgramProject
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ProgramProject(pProgramProject As ProgramProject)
    On Error GoTo ErrorHandler:
    If Not pProgramProject Is Nothing Then
        Set m_ProgramProject = New ProgramProject
        m_ProgramProject = pProgramProject
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        ProgramProject
'   Parameters:  Void
'   RetVal:      ProgramProject
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ProgramProject() As ProgramProject
    On Error GoTo ErrorHandler:
    If Not m_ProgramProject Is Nothing Then
        ProgramProject = m_ProgramProject
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        ProgramArea
'   Parameters:  ProgramArea
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ProgramArea(pProgramArea As ProgramArea)
    On Error GoTo ErrorHandler:
    If Not pProgramArea Is Nothing Then
        Set m_ProgramArea = New ProgramArea
        m_ProgramArea = pProgramArea
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        ProgramArea
'   Parameters:  Void
'   RetVal:      ProgramArea
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ProgramArea() As ProgramArea
    On Error GoTo ErrorHandler:
    If Not m_ProgramArea Is Nothing Then
        ProgramArea = m_ProgramArea
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        Amount
'   Parameters:  Double
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Amount(pAmount As Double)
    On Error GoTo ErrorHandler:
    If pAmount > -1 Then
        m_Amount = pAmount
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:        Amount
'   Parameters:  Void
'   RetVal:      Double
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Amount() As Double
    On Error GoTo ErrorHandler:
    If m_Amount > -1 Then
        Amount = m_Amount
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




