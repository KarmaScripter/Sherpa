VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionBudgeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public m_BudgetArgs As BudgetArgs
Public m_Dialog As Form_BudgetDialog
Public m_ControlNumber As ControlNumber
Public ControlStringArgs As String
Public PublicLawArgs As String
Public FundTitleArgs As String
Public AppropriationArgs As String
Public FundCodeArgs As String
Public RcCodeArgs As String
Public FiscalYearArgs As String
Public AhCodeArgs As String
Public TreasurySymbolArgs As String
Public DivisionTitleArgs As String






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Report_Load
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Load()
    On Error GoTo ErrorHandler:
    Me.ControlNumberTextBox.SetFocus
    Me.ControlNumberTextBox.text = ControlStringArgs
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Integer
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErroHandler:
    Set m_BudgetArgs = New BudgetArgs
    Set m_ControlNumber = New ControlNumber
    DoCmd.OpenForm FormName:="BudgetDialog", WindowMode:=acDialog
    Set m_BudgetArgs = Forms("BudgetDialog").Args
    FundCodeArgs = m_BudgetArgs.FundCode
    RcCodeArgs = m_BudgetArgs.RcCode
    FiscalYearArgs = m_BudgetArgs.BFY
    AhCodeArgs = m_BudgetArgs.AhCode
    AppropriationArgs = DLookup("Title", "Appropriations", "BFY = '" & FiscalYearArgs & "'")
    PublicLawArgs = DLookup("PublicLaw", "Appropriations", "BFY = '" & FiscalYearArgs & "'")
    FundTitleArgs = DLookup("Title", "Funds", "Funds.Code = '" & FundCodeArgs & "'")
    TreasurySymbolArgs = DLookup("TreasurySymbol", "Funds", "Funds.Code = '" & FundCodeArgs & "'")
    DivisionTitleArgs = DLookup("Title", "Divisions", "Divisions.Code = '" & RcCodeArgs & "'")
    Dim SQL As String
    SQL = "SELECT *" _
    & " FROM Allocations" _
    & " WHERE Allocations.BocCode <> '10' AND Allocations.BocCode <> '17'" _
    & " AND Allocations.AhCode = '" & m_BudgetArgs.AhCode & "'" _
    & " AND Allocations.Amount <> 0" _
    & " AND Allocations.BFY = '" & m_BudgetArgs.BFY & "'" _
    & " AND Allocations.FundCode = '" & m_BudgetArgs.FundCode & "'" _
    & " AND ALlocations.RcCode = '" & m_BudgetArgs.RcCode & "'" _
    & " AND Allocations.BudgetLevel = '8';"
    Me.RecordSource = SQL
     m_ControlNumber.AhCode = m_BudgetArgs.AhCode
     m_ControlNumber.BFY = m_BudgetArgs.BFY
     m_ControlNumber.RcCode = m_BudgetArgs.RcCode
     m_ControlNumber.FundCode = m_BudgetArgs.FundCode
    ControlStringArgs = m_ControlNumber.AhCode & "-" _
        & m_ControlNumber.CalendarYear & m_ControlNumber.FundCode & "-" _
        & Format(m_ControlNumber.FundNumber, "00") & "-" _
        & m_ControlNumber.Division & "-" & Format(m_ControlNumber.DivisionNumber, "00") & "-" _
        & year(m_ControlNumber.IssuedDate) _
        & Month(m_ControlNumber.IssuedDate) _
        & Day(m_ControlNumber.IssuedDate)
    DoCmd.Close ObjectType:=acForm, ObjectName:="BudgetDialog"
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




