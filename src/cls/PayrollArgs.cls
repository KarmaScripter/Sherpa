VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PayrollArgs"
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
Private m_FundName As String
Private m_HROrgCode As String
Private m_AccountCode As String
Private m_WorkCode As String
Private m_BocCode As String
Private m_BocName As String
Private m_RcCode As String
Private m_DivisionName As String
Private m_EmployeeNumber As String
Private m_LastName As String
Private m_FirstName As String
Private m_PayPeriod As String
Private m_PayGrade As String
Private m_Step As String
Private m_StartDate As String
Private m_EndDate As String
Private m_JobTitle As String
Private m_ReportingCode As String
Private m_ReportingCodeName As String
Private m_FocCode As String
Private m_FocName As String
Private m_Error As String


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
Public Property Let BFY(yr As String)
    On Error GoTo ErrorHandler:
    If Not yr & "" = "" Then
        m_BFY = yr
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     BFY" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     BFY" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     StartDate" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     StartDate" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     EndDate" _
            & vbCrLf & "Descript:   " & Err.Description
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
    If Not m_EndDate & "" = "" Then
        EndDate = m_EndDate
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     EndDate" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     RpioCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     RpioCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let AhCode(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_AhCode = ac
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     AhCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     AhCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let HrOrgCode(oc As String)
    On Error GoTo ErrorHandler:
    If Not oc & "" = "" Then
       m_HROrgCode = oc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     HrOrgCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get HrOrgCode() As String
    On Error GoTo ErrorHandler:
    If Not m_HROrgCode & "" = "" Then
        HrOrgCode = m_HROrgCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     HrOrgCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let WorkCode(wc As String)
    On Error GoTo ErrorHandler:
    If Not wc & "" = "" Then
        m_WorkCode = wc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     WorkCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get WorkCode() As String
    On Error GoTo ErrorHandler:
    If Not m_WorkCode & "" = "" Then
        WorkCode = m_WorkCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     WorkCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     FundCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     FundName" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     FundName" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     FundCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     AccountCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     AccountCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let EmployeeNumber(eid As String)
    On Error GoTo ErrorHandler:
    If eid & "" = "" Then
        m_EmployeeNumber = eid
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     EmployeeNumber" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get EmployeeNumber() As String
    On Error GoTo ErrorHandler:
    If Not m_EmployeeNumber & "" = "" Then
        EmployeeNumber = m_EmployeeNumber
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     EmployeeNumber" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     BocCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     BocCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     RcCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     RcCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     DivisionName" _
            & vbCrLf & "Descript:   " & Err.Description
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
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     DivisionName" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let PayPeriod(pyp As String)
    On Error GoTo ErrorHandler:
    If Not pyp & "" = "" Then
        m_PayPeriod = pyp
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     PayPeriod" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get PayPeriod() As String
    On Error GoTo ErrorHandler:
    If Not m_PayPeriod & "" = "" Then
        PayPeriod = m_PayPeriod
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     PayPeriod" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let PayGrade(pg As String)
    On Error GoTo ErrorHandler:
    If Not pg & "" = "" Then
        m_PayGrade = pg
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     PayGrade" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get PayGrade() As String
    On Error GoTo ErrorHandler:
    If Not m_PayGrade & "" = "" Then
        PayGrade = m_PayGrade
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     PayGrade" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let JobTitle(jt As String)
    On Error GoTo ErrorHandler:
    If Not jt & "" = "" Then
        m_JobTitle = jt
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     JobTitle" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get JobTitle() As String
    On Error GoTo ErrorHandler:
    If Not m_JobTitle & "" = "" Then
        JobTitle = m_JobTitle
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     JobTitle" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let ReportingCode(ppl As String)
    On Error GoTo ErrorHandler:
    If Not ppl & "" = "" Then
        m_ReportingCode = ppl
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     ReportingCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get ReportingCode() As String
    On Error GoTo ErrorHandler:
    If Not m_ReportingCode & "" = "" Then
        ReportingCode = m_ReportingCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     ReportingCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let ReportingCodeName(ppn As String)
    On Error GoTo ErrorHandler:
    If Not ppn & "" = "" Then
        m_ReportingCodeName = ppn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     ReportingCodeName" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get ReportingCodeName() As String
    On Error GoTo ErrorHandler:
    If Not m_ReportingCodeName & "" = "" Then
        ReportingCodeName = m_ReportingCodeName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     ReportingCodeName" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let FocCode(fcc As String)
    On Error GoTo ErrorHandler:
    If Not fcc & "" = "" Then
        m_FocCode = fcc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     FocCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get FocCode() As String
    On Error GoTo ErrorHandler:
    If Not m_FocCode & "" = "" Then
        FocCode = m_FocCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     FocCode" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Let FocName(fcn As String)
    On Error GoTo ErrorHandler:
    If Not fcn & "" = "" Then
        m_FocName = fcn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     FocName" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Property Get FocName() As String
    On Error GoTo ErrorHandler:
    If Not m_FocName & "" = "" Then
        FocName = m_FocName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     FocName" _
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
'   Retval          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let BocName(bn As String)
    On Error GoTo ErrorHandler:
    If Not bn & "" = "" Then
        m_BocName = bn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     BocName" _
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
'   Retval          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get BocName() As String
    On Error GoTo ErrorHandler:
    If Not m_BocName & "" = "" Then
        BocName = m_BocName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   PayrollArgs" _
            & vbCrLf & "Member:     BocName" _
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



