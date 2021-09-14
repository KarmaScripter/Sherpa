VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Budget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_AhCode As String
Private m_BFY As String
Private m_RcCode As String
Private m_FundCode As String
Private Args As BudgetArgs
Private m_Division As String
Private m_ControlNumber As String
Private m_Error As String

Private Sub Class_Initialize()
    Set Args = New BudgetArgs
End Sub


'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let AhCode(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_AhCode = ac
        Args.AhCode = m_AhCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
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


'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let BFY(fy As String)
    On Error GoTo ErrorHandler:
    If Not fy & "" = "" Then
        m_BFY = fy
        Args.BFY = m_BFY
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
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


'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let RcCode(rc As String)
    On Error GoTo ErrorHandler:
    If Not rc & "" = "" Then
        m_RcCode = rc
        Args.RcCode = m_RcCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
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



'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let FundCode(fc As String)
    On Error GoTo ErrorHandler:
    If Not fc & "" = "" Then
        m_FundCode = fc
        Args.FundCode = m_FundCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
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


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get PublicLaw() As String
    On Error GoTo ErrorHandler:
    If Not m_BFY & "" = "" Then
        PublicLaw = DLookup("PublicLaw", "Appropriations", _
            "Appropriations.BFY = '" & m_BFY & "'")
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get AppropriationBill() As String
    On Error GoTo ErrorHandler:
    If Not m_BFY & "" = "" Then
        AppropriationBill = DLookup("Title", "Appropriations", _
            "Appropriations.BFY = '" & m_BFY & "'")
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get DivisionName() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        DivisionName = DLookup("Title", "Divisions", _
            "Divisions.Code = '" & m_RcCode & "'")
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get DivisionCaption() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        DivisionCaption = DLookup("Caption", "Divisions", _
            "Divisions.Code = '" & m_RcCode & "'")
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name            Generate
'   Parameters      Void
'   Purpose         Initializes BudgetArg m_embers
'---------------------------------------------------------------------------------
Public Sub Generate(bargs As BudgetArgs)
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Dim Data As DAO.Recordset
    Dim query As DAO.QueryDef
    Dim Budget As Report
    If Not Args.AhCode & "" = "" Then
        AhCode = Args.AhCode
    Else
        AhCode = "06"
    End If
    If Not bargs.BFY & "" = "" Then
        BFY = Args.BFY
    End If
    If Not bargs.RcCode & "" = "" Then
        RcCode = Args.RcCode
    End If
    If Not bargs.FundCode & "" = "" Then
        FundCode = Args.FundCode
    End If
    Set DB = CurrentDb
    Set Budget = Forms("DivisionBudgetAuthority")
    Set query = DB.QueryDefs("SelectParamBudgetAuthorityData")
    query.Parameters("AhCodeArgs") = AhCode
    query.Parameters("FiscalYearArgs") = BFY
    query.Parameters("RcCodeArgs") = RcCode
    query.Parameters("FundCodeArgs") = FundCode
    Set Data = query.OpenRecordset(dbOpenDynaset)
    Budget.Recordset = Data
    DoCmd.OpenReport (Budget.Name)
    query.Close
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



