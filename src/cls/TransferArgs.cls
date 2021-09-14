VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private m_TransferId As Integer
Private m_BFY As String
Private m_RpioCode As String
Private m_RpioName As String
Private m_AhCode As String
Private m_AhName As String
Private m_OrgCode As String
Private m_OrgName As String
Private m_FundCode As String
Private m_FundName As String
Private m_AccountCode As String
Private m_ProgramProjectCode As String
Private m_ProgramProjectName As String
Private m_ProgramAreaCode As String
Private m_ProgramAreaName As String
Private m_NpmCode As String
Private m_NpmName As String
Private m_FromTo As String
Private m_BocCode As String
Private m_RcCode As String
Private m_DivisionName As String
Private m_StartDate As String
Private m_EndDate As String
Private m_ProcessedDate As Date
Private m_ReprogrammingNumber As String
Private m_ControlNumber As String
Private m_DocPrefix As String
Private m_DocType As String
Private m_Line As String
Private m_Subline As String
Private m_Purpose As String
Private m_ExtendedPurpose As String
Private m_ResourceType As String
Private m_Amount As Double
Private m_Error As String
Private m_Notification As String





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
Public Property Let TransferId(tid As Integer)
    On Error GoTo ErrorHandler:
    If tid > 0 Then
        m_TransferId = tid
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      TransferId" _
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
Public Property Get TransferId() As Integer
    On Error GoTo ErrorHandler:
    If m_TransferId > 0 Then
        TransferId = m_TransferId
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      TransferId" _
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
Public Property Let BFY(yr As String)
    On Error GoTo ErrorHandler:
    If Not yr & "" = "" Then
        m_BFY = yr
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      BFY" _
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
Public Property Get BFY() As String
    On Error GoTo ErrorHandler:
    If Not m_BFY & "" = "" Then
        BFY = m_BFY
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      BFY" _
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
Public Property Let StartDate(sd As String)
    On Error GoTo ErrorHandler:
    If Not sd & "" = "" Then
        m_StartDate = sd
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      StartDate" _
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
Public Property Get StartDate() As String
    On Error GoTo ErrorHandler:
    If Not m_StartDate & "" = "" Then
        StartDate = m_StartDate
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      StartDate" _
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
Public Property Let EndDate(ed As String)
    On Error GoTo ErrorHandler:
    If Not ed & "" = "" Then
        m_EndDate = ed
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferEndDateArgs" _
            & vbCrLf & "Member:      " _
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
Public Property Get EndDate() As String
    On Error GoTo ErrorHandler:
    If IsDate(m_EndDate) Then
        EndDate = m_EndDate
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      EndDate" _
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
Public Property Let RpioCode(rp As String)
    On Error GoTo ErrorHandler:
    If Not rp & "" = "" Then
        m_RpioCode = rp
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      RpioCode" _
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
Public Property Get RpioCode() As String
    On Error GoTo ErrorHandler:
    If Not m_RpioCode & "" = "" Then
        RpioCode = m_RpioCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      RpioCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let RpioName(rn As String)
    On Error GoTo ErrorHandler:
    If Not rn & "" = "" Then
        m_RpioName = rn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      RpioName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get RpioName() As String
    On Error GoTo ErrorHandler:
    If Not m_RpioName & "" = "" Then
        RpioName = m_RpioName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      AhCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property

'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let AhCode(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_AhCode = ac
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      AhCode" _
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
Public Property Get AhCode() As String
    On Error GoTo ErrorHandler:
    If Not m_AhCode & "" = "" Then
        AhCode = m_AhCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      AhName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let AhName(an As String)
    On Error GoTo ErrorHandler:
    If Not an & "" = "" Then
        m_AhName = an
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      AhName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get AhName() As String
    On Error GoTo ErrorHandler:
    If Not m_AhName & "" = "" Then
        AhName = m_AhName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      OrgCode" _
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
Public Property Let OrgCode(oc As String)
    On Error GoTo ErrorHandler:
    If Not oc & "" = "" Then
       m_OrgCode = oc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      OrgCode" _
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
Public Property Get OrgCode() As String
    On Error GoTo ErrorHandler:
    If Not m_OrgCode & "" = "" Then
        OrgCode = m_OrgCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      OrgName" _
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
Public Property Let OrgName(wc As String)
    On Error GoTo ErrorHandler:
    If Not wc & "" = "" Then
        m_OrgName = wc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      OrgName" _
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
Public Property Get OrgName() As String
    On Error GoTo ErrorHandler:
    If Not m_OrgName & "" = "" Then
        OrgName = m_OrgName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      OrgName" _
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
Public Property Let FundCode(fc As String)
    On Error GoTo ErrorHandler:
    If Not fc & "" = "" Then
        m_FundCode = fc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      FundCode" _
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
Public Property Get FundCode() As String
    On Error GoTo ErrorHandler:
    If Not m_FundCode & "" = "" Then
        FundCode = m_FundCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      FundCode" _
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
Public Property Get FundName() As String
    On Error GoTo ErrorHandler:
    If Not m_FundName & "" = "" Then
        FundName = m_FundName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      FundName" _
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
Public Property Let FundName(fn As String)
    On Error GoTo ErrorHandler:
    If Not fn & "" = "" Then
        m_FundName = fn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      FundName" _
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
Public Property Let AccountCode(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_AccountCode = ac
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      AccountCode" _
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
Public Property Get AccountCode() As String
    On Error GoTo ErrorHandler:
    If Not m_AccountCode & "" = "" Then
        AccountCode = m_AccountCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      AccountCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property






'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ProgramProjectCode(ppc As String)
    On Error GoTo ErrorHandler:
    If Not ppc & "" = "" Then
        m_ProgramProjectCode = ppc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProgramProjectCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ProgramProjectCode() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramProjectCode & "" = "" Then
        ProgramProjectCode = m_ProgramProjectCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProgramProjectCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ProgramProjectName(ppn As String)
    On Error GoTo ErrorHandler:
    If Not ppn & "" = "" Then
        m_ProgramProjectName = ppn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProgramProjectName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ProgramProjectName() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramProjectName & "" = "" Then
        ProgramProjectName = m_ProgramProjectName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProgramProjectName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ProgramAreaCode(pac As String)
    On Error GoTo ErrorHandler:
    If Not pac & "" = "" Then
        m_ProgramAreaCode = pac
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProgramAreaCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ProgramAreaCode() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramAreaCode & "" = "" Then
        ProgramAreaCode = m_ProgramAreaCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProgramAreaCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ProgramAreaName(pan As String)
    On Error GoTo ErrorHandler:
    If Not pan & "" = "" Then
        m_ProgramAreaName = pan
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProgramAreaName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ProgramAreaName() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramAreaName & "" = "" Then
        ProgramAreaName = m_ProgramAreaName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProgramAreaName" _
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
Public Property Let BocCode(bc As String)
    On Error GoTo ErrorHandler:
    If Not bc & "" = "" Then
        m_BocCode = bc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      BocCode" _
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
Public Property Get BocCode() As String
    On Error GoTo ErrorHandler:
    If Not m_BocCode & "" = "" Then
        BocCode = m_BocCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      BocCode" _
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
Public Property Let RcCode(rc As String)
    On Error GoTo ErrorHandler:
    If Not rc & "" = "" Then
        m_RcCode = rc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      RcCode" _
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
Public Property Get RcCode() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        RcCode = m_RcCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      RcCode" _
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
Public Property Let DivisionName(dn As String)
    On Error GoTo ErrorHandler:
    If Not dn & "" = "" Then
        m_DivisionName = dn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      DivisionName" _
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
Public Property Get DivisionName() As String
    On Error GoTo ErrorHandler:
    If Not m_DivisionName & "" = "" Then
        DivisionName = m_DivisionName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      DivisionName" _
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
Public Property Let ProcessedDate(pd As String)
    On Error GoTo ErrorHandler:
    If Not pd & "" = "" Then
        m_ProcessedDate = pd
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProcessedDate" _
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
Public Property Get ProcessedDate() As String
    On Error GoTo ErrorHandler:
    If Not m_ProcessedDate & "" = "" Then
        ProcessedDate = m_ProcessedDate
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ProcessedDate" _
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
Public Property Let DocType(dt As String)
    On Error GoTo ErrorHandler:
    If Not dt & "" = "" Then
        m_DocType = dt
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      DocType" _
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
Public Property Get DocType() As String
    On Error GoTo ErrorHandler:
    If Not m_DocType & "" = "" Then
        DocType = m_DocType
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      DocType" _
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
Public Property Let ReprogrammingNumber(rpn As String)
    On Error GoTo ErrorHandler:
    If Not rpn & "" = "" Then
        m_ReprogrammingNumber = rpn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ReprogrammingNumber" _
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
Public Property Get ReprogrammingNumber() As String
    On Error GoTo ErrorHandler:
    If Not m_ReprogrammingNumber & "" = "" Then
        ReprogrammingNumber = m_ReprogrammingNumber
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ReprogrammingNumber" _
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
Public Property Let ControlNumber(cn As String)
    On Error GoTo ErrorHandler:
    If Not cn & "" = "" Then
        m_ControlNumber = cn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ControlNumber" _
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
Public Property Get ControlNumber() As String
    On Error GoTo ErrorHandler:
    If Not m_ControlNumber & "" = "" Then
        ControlNumber = m_ControlNumber
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ControlNumber" _
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
Public Property Let Line(eid As String)
    On Error GoTo ErrorHandler:
    If eid & "" = "" Then
        m_Line = eid
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      Line" _
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
Public Property Get Line() As String
    On Error GoTo ErrorHandler:
    If Not m_Line & "" = "" Then
        Line = m_Line
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      Line" _
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
Public Property Let Subline(sl As String)
    On Error GoTo ErrorHandler:
    If Not sl & "" = "" Then
        m_Subline = sl
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      Subline" _
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
Public Property Get Subline() As String
    On Error GoTo ErrorHandler:
    If Not m_Subline & "" = "" Then
        Subline = m_Subline
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      Subline" _
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
Public Property Let FromTo(fcc As String)
    On Error GoTo ErrorHandler:
    If Not fcc & "" = "" Then
        m_FromTo = fcc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      FromTo" _
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
Public Property Get FromTo() As String
    On Error GoTo ErrorHandler:
    If Not m_FromTo & "" = "" Then
        FromTo = m_FromTo
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      FromTo" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Purpose(pr As String)
    On Error GoTo ErrorHandler:
    If Not pr & "" = "" Then
        m_Purpose = pr
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      Purpose" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Purpose() As String
    On Error GoTo ErrorHandler:
    If Not m_Purpose & "" = "" Then
        Purpose = m_Purpose
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      Purpose" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ExtendedPurpose(ep As String)
    On Error GoTo ErrorHandler:
    If Not ep & "" = "" Then
        m_ExtendedPurpose = ep
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ExtendedPurpose" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ExtendedPurpose() As String
    On Error GoTo ErrorHandler:
    If Not m_ExtendedPurpose & "" = "" Then
        ExtendedPurpose = m_ExtendedPurpose
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ExtendedPurpose" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ResourceType(rt As String)
    On Error GoTo ErrorHandler:
    If Not rt & "" = "" Then
        m_ResourceType = rt
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ResourceType" _
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
Public Property Get ResourceType() As String
    On Error GoTo ErrorHandler:
    If Not m_ResourceType & "" = "" Then
        ResourceType = m_ResourceType
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      ResourceType" _
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
Public Property Let Amount(am As Double)
    On Error GoTo ErrorHandler:
    If am > 0 Then
        m_Amount = am
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      Amount" _
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
Public Property Get Amount() As Double
    On Error GoTo ErrorHandler:
    If m_Amount > 0 Then
        Amount = m_Amount
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   TransferArgs" _
            & vbCrLf & "Member:      Amount" _
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



