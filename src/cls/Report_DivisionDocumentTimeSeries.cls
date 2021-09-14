VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionDocumentTimeSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Public Args As AccountArgs
Public m_Dialog As Form_ReportDialog
Public FundCodeArgs As String
Public RcCodeArgs As String
Public FiscalYearArgs As String
Public DivisionTitleArgs As String
Public AhCodeArgs As String
Public RpioCodeArgs As String
Public PublicLawArgs As String
Public FundTitleArgs As String
Public TreasurySymbolArgs As String
Private m_Error As String







'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Load()
    On Error GoTo ErrorHandler:
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Integer
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set Args = New AccountArgs
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = m_Dialog.Args
    Dim SQL As String
    If Not Args.FundName & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
            FiscalYearArgs = Args.BFY
            DivisionTitleArgs = DLookup("Title", "Divisions", "Divisions.Code = '" & Args.RcCode & "'")
            PublicLawArgs = DLookup("PublicLaw", "Appropriations", "BFY = '" & FiscalYearArgs & "'")
            SQL = "SELECT Purchases.*, Allocations.Amount" _
            & " FROM Allocations" _
            & " LEFT JOIN Purchases" _
            & " ON (Allocations.RcCode = Purchases.RcCode)" _
            & " AND (Allocations.BocCode = Purchases.BocCode)" _
            & " AND (Allocations.AccountCode = Purchases.AccountCode)" _
            & " AND (Allocations.FundCode = Purchases.FundCode)" _
            & " AND (Allocations.BFY = Purchases.BFY)" _
            & " WHERE NOT IsNull(Purchases.DocumentControlNumbers)" _
            & " AND Purchases.Used > 0" _
            & " AND Purchases.BFY = '" & Args.BFY & "'" _
            & " AND Purchases.DivisionName = '" & Args.DivisionName & "'" _
            & " AND Purchases.FundName = '" & Args.FundName & "';"
    End If
    If Args.FundCode = "" Then
        SQL = "SELECT Purchases.*, Allocations.Amount" _
        & " FROM Allocations" _
        & " LEFT JOIN Purchases" _
        & " ON (Allocations.RcCode = Purchases.RcCode)" _
        & " AND (Allocations.BocCode = Purchases.BocCode)" _
        & " AND (Allocations.AccountCode = Purchases.AccountCode)" _
        & " AND (Allocations.BFY = Purchases.BFY)" _
        & " WHERE NOT IsNull(Purchases.DocumentControlNumbers)" _
        & " AND Purchases.Used > 0" _
        & " AND Purchases.BFY = '" & Args.BFY & "'" _
        & " AND Purchases.RcCode = '" & Args.RcCode & "';"
    End If
    Me.RecordSource = SQL
    DoCmd.Close ObjectType:=acForm, ObjectName:="ReportDialog"
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




