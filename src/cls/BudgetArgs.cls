VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BudgetArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_BFY As String
Private m_RcCode As String
Private m_FundCode As String
Private m_FundName As String
Private m_AhCode As String
Private m_RPIO As String
Private m_Division As String
Private m_DivisionName As String
Private m_BocCode As String
Private m_BocName As String
Private m_AccountCode As String
Private m_ProgramProjectCode As String
Private m_ProgramProjectName As String
Private m_StartDate As String
Private m_EndDate As String
Private m_Error As String


Public Sub Class_Initialize()
End Sub

Public Sub Class_Terminate()
End Sub







'---------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BocCode(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_BocCode = ac
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      BocCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BocCode() As String
    On Error GoTo ErrorHandler:
    If Not m_BocCode & "" = "" Then
        BocCode = m_BocCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      BocCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BocName(bn As String)
    On Error GoTo ErrorHandler:
    If Not bn & "" = "" Then
        m_BocName = bn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      BocName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BocName() As String
    On Error GoTo ErrorHandler:
    If Not m_BocName & "" = "" Then
        BocName = m_BocName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      BocName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let AhCode(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_AhCode = Left(ac, 2)
        m_RPIO = m_AhCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      AhCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get AhCode() As String
    On Error GoTo ErrorHandler:
    If Not m_AhCode & "" = "" Then
        AhCode = m_AhCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      RPIO" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RPIO() As String
    On Error GoTo ErrorHandler:
    If Not m_RPIO & "" = "" Then
        RPIO = m_RPIO
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      RPIO" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BFY(yr As String)
    On Error GoTo ErrorHandler:
    If Not yr & "" = "" Then
        m_BFY = yr
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      BFY" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BFY() As String
    On Error GoTo ErrorHandler:
    If Not m_BFY & "" = "" Then
        BFY = m_BFY
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      BFY" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ProgramProjectName(pn As String)
    On Error GoTo ErrorHandler:
    If Not pn & "" = "" Then
        m_ProgramProjectName = pn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      ProgramProjectName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ProgramProjectName() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramProjectName & "" = "" Then
        ProgramProjectName = m_ProgramProjectName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      ProgramProjectName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ProgramProjectCode(pc As String)
    On Error GoTo ErrorHandler:
    If Not pc & "" = "" Then
        m_ProgramProjectCode = pc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      ProgramProjectCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ProgramProjectCode() As String
    On Error GoTo ErrorHandler:
    If Not m_ProgramProjectCode & "" = "" Then
        ProgramProjectCode = m_ProgramProjectCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      ProgramProjectCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
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
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      RcCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RcCode() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        RcCode = m_RcCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      RcCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let DivisionName(dn As String)
    On Error GoTo ErrorHandler:
    If Not dn & "" = "" Then
        m_DivisionName = dn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      DivisionName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Division() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        RcCode = DLookup("Caption", "Divisions", _
            "Divisions.Code = '" & m_RcCode & "'")
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      DivisionName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DivisionName() As String
    On Error GoTo ErrorHandler:
    If Not m_DivisionName & "" = "" Then
        DivisionName = m_DivisionName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      DivisionName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
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
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      FundCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FundCode() As String
    On Error GoTo ErrorHandler:
    If Not m_FundCode & "" = "" Then
        FundCode = m_FundCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      FundCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FundName() As String
    On Error GoTo ErrorHandler:
    If Not m_FundName & "" = "" Then
        FundName = m_FundName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
            & vbCrLf & "Member:      FundName" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FundName(fn As String)
    On Error GoTo ErrorHandler:
    If Not fn & "" = "" Then
        m_FundName = fn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
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
Public Property Let StartDate(sd As String)
    On Error GoTo ErrorHandler:
    If Not sd & "" = "" Then
        m_StartDate = sd
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      BudgetArgs" _
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
        m_Error = "Source:      BudgetArgs" _
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
        m_Error = "Source:      BudgetArgs" _
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
Public Property Get EndDate() As String
    On Error GoTo ErrorHandler:
    If IsDate(m_EndDate) Then
        EndDate = m_EndDate
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetArgs" _
            & vbCrLf & "Member:      EndDate" _
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



