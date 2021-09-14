VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private pAnd As String
Private m_Insert As String
Private m_IssuedDate As Date
Private m_ControlNumberId As Integer
Private m_FundCode As String
Private m_FundControl As String
Private m_FundNumber As Integer
Private m_DivisionControl As String
Private m_DivisionNumber As Integer
Private m_RpioCode As String
Private m_RegionControl As String
Private m_RegionNumber As Integer
Private m_FCN As Integer
Private m_DCN As String
Private m_CalendarYear As String
Private m_BFY As String
Private m_OrgCode As String
Private m_AhCode As String
Private m_RcCode As String
Private m_DivisionName As String
Private m_Purpose As String
Private m_Error As String



Private Sub Class_Initialize()
    pAnd = "-"
End Sub




'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        ControlNumberId
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ControlNumberId(cid As Integer)
    On Error GoTo ErrorHandler:
    If cid > 0 Then
        m_ControlNumberId = cid
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      ControlNumberId" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        ControlNumberId
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ControlNumberId() As Integer
    On Error GoTo ErrorHandler:
    If m_ControlNumberId > 0 Then
        ControlNumberId = m_ControlNumberId
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      ControlNumberId" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        RpioCode
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get RpioCode() As String
    On Error GoTo ErrorHandler:
    If Not m_AhCode & "" = "" And _
        m_AhCode = "06" Then
            RpioCode = "06"
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      RpioCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        AhCode
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let AhCode(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_AhCode = ac
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      AhCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        AhCode
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get AhCode() As String
    On Error GoTo ErrorHandler:
    If Not m_AhCode & "" = "" Then
        AhCode = m_AhCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      AhCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        BFY
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let BFY(fy As String)
    On Error GoTo ErrorHandler:
    If Not fy & "" = "" Then
        m_BFY = fy
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      BFY" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        BFY
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get BFY() As String
    On Error GoTo ErrorHandler:
    If Not m_BFY & "" = "" Then
        BFY = m_BFY
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      BFY" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        RcCode
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let RcCode(rc As String)
    On Error GoTo ErrorHandler:
    If Not rc & "" = "" Then
        m_RcCode = rc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      RcCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        RcCode
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get RcCode() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        RcCode = m_RcCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      RcCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        FundCode
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let FundCode(fcn As String)
    On Error GoTo ErrorHandler:
    If Not fcn & "" = "" Then
        m_FundCode = fcn
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      FundCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        FundCode
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get FundCode() As String
    On Error GoTo ErrorHandler:
    If Not m_FundCode & "" = "" Then
        FundCode = m_FundCode
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      FundCode" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get FundNumber() As Integer
    On Error GoTo ErrorHandler:
    If Not m_FundCode & "" = "" Then
        m_FundNumber = DCount("FundCode", "ControlNumbers", _
            "ControlNumbers.FundCode = '" & m_FundCode & "'") + 1
        FundNumber = m_FundNumber
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      FundNumber" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        Division
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Division() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        Dim Name As String
        Name = DLookup("Name", "Divisions", _
            "Divisions.Code = '" & m_RcCode & "'")
        Division = Name
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      Division" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        DivisionNumber
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get DivisionNumber() As Integer
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        m_DivisionNumber = DCount("RcCode", "ControlNumbers", _
            "ControlNumbers.RcCode = '" & m_RcCode & "'") + 1
        DivisionNumber = m_DivisionNumber
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      DivisionNumber" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        RegionNumber
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get RegionNumber() As Integer
    On Error GoTo ErrorHandler:
    If Not m_AhCode & "" = "" Then
        m_RegionNumber = DCount("ControlNumberId", "ControlNumbers", _
            "ControlNumbers.AhCode = '" & m_AhCode & "'") + 1
        RegionNumber = m_RegionNumber
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      RegionNumber" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        CalendarYear
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get CalendarYear() As String
    On Error GoTo ErrorHandler:
    If Not m_BFY & "" = "" Then
        CalendarYear = Right(m_BFY, 2)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      CalendarYear" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        IssuedDate
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let IssuedDate(dt As Date)
    On Error GoTo ErrorHandler:
    If IsDate(dt) Then
        m_IssuedDate = dt
    Else
        m_IssuedDate = Now()
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      IssuedDate" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        IssuedDate
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get IssuedDate() As Date
    On Error GoTo ErrorHandler:
        m_IssuedDate = Now()
        IssuedDate = m_IssuedDate
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      IssuedDate" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        Purpose
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Purpose(ps As String)
    On Error GoTo ErrorHandler:
    If Not ps & "" = "" Then
        m_Purpose = ps
    End If
    If ps & "" = "" Then
        m_Purpose = "NOT SPECIFIED"
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      Purpose" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        Purpose
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Purpose() As String
    On Error GoTo ErrorHandler:
    If Not m_Purpose & "" = "" Then
        Purpose = m_Purpose
    End If
    If m_Purpose & "" = "" Then
        m_Purpose = "NOT SPECIFIED"
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      Purpose" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        Create
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Sub Create(ac As String, fy As String, rc As String, fc As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
       AhCode = ac
    End If
    If Not fy & "" = "" Then
        BFY = fy
    End If
    If Not rc & "" = "" Then
        RcCode = rc
    End If
    If Not fc & "" = "" Then
        FundCode = fc
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      Create" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name        Generate
'   Parameters  String
'   RetVal      String
'   Purpose
'---------------------------------------------------------------------------------
Public Function Generate(ac As String, fy As String, rc As String, fc As String) As String
    On Error GoTo ErrorHandler:
    Dim control As String
    If Not ac & "" = "" Then
        AhCode = ac
    End If
    If Not fy & "" = "" Then
        BFY = fy
    End If
    If Not rc & "" = "" Then
        RcCode = rc
    End If
    If Not fc & "" = "" Then
        FundCode = fc
    End If
    If Not AhCode & "" = "" And _
        Not BFY & "" = "" And _
        Not RcCode & "" = "" And _
        Not FundCode & "" = "" And _
        Not Division & "" = "" Then
            control = AhCode & pAnd & BFY & pAnd _
                & CalendarYear & FundCode & pAnd _
                & FundNumber & pAnd _
                & Division & pAnd _
                & DivisionNumber & pAnd _
                & RegionNumber
    End If
    Generate = control
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      Generate" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type        Property Procedure
'   Name
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub add()
    On Error GoTo ErrorHandler:
    Dim SQL As String
    SQL = "INSERT INTO ControlNumbers (RpioCode, RegionNumber, AhCode, BFY, CalendarYear, " _
        & "RcCode, DivisionNumber, FundCode, FundNumber, IssuedDate) " _
        & "VALUES " & Me.RpioCode & ", " & Me.RegionNumber _
        & ", " & Me.AhCode & ", " & Me.BFY & ", " & Me.CalendarYear _
        & ", " & Me.RcCode & ", " & Me.DivisionNumber _
        & ", " & Me.FundCode & ", " & Me.FundNumber _
        & ", " & Me.IssuedDate & ";"
    DoCmd.RunSQL SQL
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      ControlNumber" _
            & vbCrLf & "Member:      Add" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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



