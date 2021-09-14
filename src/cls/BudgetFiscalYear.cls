VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BudgetFiscalYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_BFY As String
Private m_EFY As String
Private m_ID As Integer
Private m_StartDate As Date
Private m_EndDate As Date
Private m_CalendarDate As Date
Private m_CalendarYear As String
Private m_Day As String
Private m_Month As String
Private m_ExpirationDate As Date
Private m_ExpiringYear As String
Private m_SweepDate As Date
Private m_IsWorkDay As Boolean
Private m_IsExpiring As Boolean
Private m_Data As AccessData
Private m_Connection As ADODB.Connection
Private m_Records As ADODB.Recordset
Private m_Error As String
Private m_Notification As String
Private m_Base As String


'---------------------------------------------------------------------------------
'   Type            Constructor
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
        m_CalendarYear = CStr(year(Now()))
        m_CalendarDate = Now()
        m_Month = Month(Now())
        m_Day = Day(Now())
End Sub



'---------------------------------------------------------------------------------
'   Type            Property-Let
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Base(yr As String)
    On Error GoTo ErrorHandler:
    If Not yr & "" = "" Then
        m_Base = yr
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Base() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Base = m_Base
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
Public Property Let ID(num As Integer)
    On Error GoTo ErrorHandler:
    If num > 0 Then
        m_ID = num
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
Public Property Get ID() As Integer
    On Error GoTo ErrorHandler:
    If m_ID > 0 Then
        ID = m_ID
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name            SetSweepDate
'   Parameters      Date
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let SweepDate(sd As Date)
    On Error GoTo ErrorHandler:
    If IsDate(sd) And _
        IsExpiring() And _
        Not IsDate(m_SweepDate) Then
            m_SweepDate = sd
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name            SetSweepDate
'   Parameters      Date
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get SweepDate() As Date
    On Error GoTo ErrorHandler:
    If IsDate(m_SweepDate) And _
        IsExpiring() And _
        Not IsDate(m_SweepDate) Then
            SweepDate = m_SweepDate
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let BFY(by As String)
    On Error GoTo ErrorHandler:
    If Not by & "" = "" Then
        m_BFY = by
    Else
        m_BFY = str(year(Now()))
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type            Function
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
'   Type            Sub-Procedure
'   Name            SetLastYear
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let EFY(by As String)
    On Error GoTo ErrorHandler:
    If Not by & "" = "" Then
        m_EFY = by
    Else
        m_EFY = m_BFY
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Function
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get EFY() As String
    On Error GoTo ErrorHandler:
    If Not m_EFY & "" = "" Then
        EFY = m_EFY
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ExpirationYear(ex As String)
    On Error GoTo ErrorHandler:
    If Not ex & "" = "" And _
        Not m_BFY & "" = "" And _
        Not m_EFY & "" = "" And _
        m_BFY <> m_EFY Then
            m_ExpiringYear = m_EFY
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type            Function
'   Name
'   Parameters      Void
'   Return          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ExpirationYear() As String
    On Error GoTo ErrorHandler:
    If Not m_ExpiringYear & "" = "" Then
        ExpirationYear = m_ExpiringYear
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type            Function
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get CalendarDate() As Date
    On Error GoTo ErrorHandler:
    If IsDate(m_CalendarDate) Then
        CalendarDate = m_CalendarDate
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Function
'   Name
'   Parameters      Void
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get CalendarMonth() As Date
    On Error GoTo ErrorHandler:
    If IsDate(m_Month) Then
        CalendarMonth = m_Month
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Function
'   Name
'   Parameters      Void
'   Return          Date
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get CalendarDay() As Date
    On Error GoTo ErrorHandler:
    If IsDate(m_Day) Then
        CalendarDay = m_Day
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Function
'   Name            GetCalendarYear
'   Parameters      Void
'   Return          Date
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get CalendarYear() As String
    On Error GoTo ErrorHandler:
    If Not m_CalendarYear & "" = "" Then
        CalendarYear = m_CalendarYear
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type            Function
'   Name            IsExpiring
'   Parameters      Void
'   Return          Boolean
'   Purpose
'---------------------------------------------------------------------------------
Public Function IsExpiring() As Boolean
    On Error GoTo ErrorHandler:
    If Not m_ExpiringYear & "" = "" Then
        IsExpiring = True
    End If
    If m_ExpiringYear & "" = "" Then
        IsExpiring = False
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function






'---------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name            SetSweepDate
'   Parameters      Date
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Function CalculateCalendarDaysUntil(sd As Date) As Integer
    On Error GoTo ErrorHandler:
    If IsDate(sd) Then
        CalculateCalendarDaysUntil = DateDiff("d", Now, sd)
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function





'---------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name            SetSweepDate
'   Parameters      Date
'   Return          Void
'   Purpose
'---------------------------------------------------------------------------------
Public Function CalculateRemainingFullTimeDays() As Integer
    On Error GoTo ErrorHandler:
    If Not m_CalendarYear & "" = "" Then
        Dim yr As String
        yr = "10/01/" & m_CalendarYear
        Dim dt As Date
        dt = CDate(yr)
        CalculateRemainingFullTimeDays = DateDiff("d", Now, dt)
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function








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



