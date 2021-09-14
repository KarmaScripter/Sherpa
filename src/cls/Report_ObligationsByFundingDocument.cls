VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_ObligationsByFundingDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public RcCodeArgs As String
Public FundCodeArgs As String
Public FiscalYearArgs As String
Private SQL As String
Public m_Dialog As Form_ReportDialog
Public pAnd As String
Public Args As AccountArgs
Private m_Error As String



'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set Args = New AccountArgs
    pAnd = " AND "
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = m_Dialog.Args
    If Not Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT Purchases.RpioCode AS RPIO, Purchases.AhCode, Purchases.BFY, Purchases.FundCode, Purchases.FundName," _
            & " Purchases.RcCode, Purchases.DivisionName," _
            & " Purchases.DocumentControlNumbers AS DocumentControlNumbers, Purchases.OrgCode, Purchases.AccountCode, Purchases.ActivityCode," _
            & " Purchases.ProgramProjectName, Purchases.BocCode, Purchases.BocName, Purchases.FocCode, Purchases.FocName," _
            & " Purchases.OriginalActionDate, Purchases.LastActionDate, CCur(SUM(Purchases.OpenCommitments)) AS OpenCommitments," _
            & " CCur(SUM(Purchases.Obligations)) AS Obligations, CCur(ABS(SUM(Purchases.Deobligations))) AS Deobligations," _
            & " CCur(ABS(SUM(Purchases.ULO))) AS ULO, CCur(ABS(SUM(Purchases.Expenditures))) AS Expenditures" _
            & " FROM Allocations LEFT JOIN Purchases ON (Allocations.BFY = Purchases.BFY)" _
            & " AND (Allocations.AhCode = LEFT(Purchases.AhCode, 2)) AND (Allocations.FundCode = Purchases.FundCode)" _
            & " AND (Allocations.AccountCode = Purchases.AccountCode) AND (Allocations.BocCode = Purchases.BocCode)" _
            & " AND (Allocations.RcCode = Purchases.RcCode)" _
            & " GROUP BY Purchases.RpioCode, Purchases.AhCode, Purchases.BFY," _
            & " Purchases.RcCode, Purchases.FundCode, Purchases.FundName, Purchases.RcCode, Purchases.DivisionName, Purchases.DocumentControlNumbers," _
            & " Purchases.OriginalActionDate, Purchases.LastActionDate," _
            & " Purchases.OrgCode, Purchases.AccountCode, Purchases.ActivityCode," _
            & " Purchases.ProgramProjectName, Purchases.BocCode, Purchases.BocName," _
            & " Purchases.FocCode, Purchases.FocName" _
            & " HAVING Purchases.BFY = '" & Args.BFY & "'" _
            & " AND Purchases.DivisionName = '" & Args.DivisionName & "'" _
            & " AND Purchases.FundName = '" & Args.FundName & "'" _
            & " AND NOT IsNull(Purchases.DocumentControlNumbers);"
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT Purchases.RpioCode AS RPIO, Purchases.AhCode, Purchases.BFY, Purchases.FundCode, Purchases.FundName," _
            & " Purchases.RcCode, Purchases.DivisionName," _
            & " Purchases.DocumentControlNumbers AS DocumentControlNumbers, Purchases.OrgCode, Purchases.AccountCode, Purchases.ActivityCode," _
            & " Purchases.ProgramProjectName, Purchases.BocCode, Purchases.BocName, Purchases.FocCode, Purchases.FocName," _
            & " Purchases.OriginalActionDate, Purchases.LastActionDate, CCur(SUM(Purchases.OpenCommitments)) AS OpenCommitments," _
            & " CCur(SUM(Purchases.Obligations)) AS Obligations, CCur(ABS(SUM(Purchases.Deobligations))) AS Deobligations," _
            & " CCur(ABS(SUM(Purchases.ULO))) AS ULO, CCur(ABS(SUM(Purchases.Expenditures))) AS Expenditures" _
            & " FROM Allocations LEFT JOIN Purchases ON (Allocations.BFY = Purchases.BFY)" _
            & " AND (Allocations.AhCode = LEFT(Purchases.AhCode, 2)) AND (Allocations.FundCode = Purchases.FundCode)" _
            & " AND (Allocations.AccountCode = Purchases.AccountCode) AND (Allocations.BocCode = Purchases.BocCode)" _
            & " AND (Allocations.RcCode = Purchases.RcCode)" _
            & " GROUP BY Purchases.RpioCode, Purchases.AhCode, Purchases.BFY," _
            & " Purchases.RcCode, Purchases.FundCode, Purchases.FundName, Purchases.RcCode, Purchases.DivisionName, Purchases.DocumentControlNumbers," _
            & " Purchases.OriginalActionDate, Purchases.LastActionDate," _
            & " Purchases.OrgCode, Purchases.AccountCode, Purchases.ActivityCode," _
            & " Purchases.ProgramProjectName, Purchases.BocCode, Purchases.BocName," _
            & " Purchases.FocCode, Purchases.FocName" _
            & " HAVING Purchases.BFY = '" & Args.BFY & "'" _
            & " AND Purchases.DivisionName = '" & Args.DivisionName & "'" _
            & " AND NOT IsNull(Purchases.DocumentControlNumbers);"
    End If
    If Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT Purchases.RpioCode AS RPIO, Purchases.AhCode, Purchases.BFY, Purchases.FundCode, Purchases.FundName," _
            & " Purchases.RcCode, Purchases.DivisionName," _
            & " Purchases.DocumentControlNumbers AS DocumentControlNumbers, Purchases.OrgCode, Purchases.AccountCode, Purchases.ActivityCode," _
            & " Purchases.ProgramProjectName, Purchases.BocCode, Purchases.BocName, Purchases.FocCode, Purchases.FocName," _
            & " Purchases.OriginalActionDate, Purchases.LastActionDate, CCur(SUM(Purchases.OpenCommitments)) AS OpenCommitments," _
            & " CCur(SUM(Purchases.Obligations)) AS Obligations, CCur(ABS(SUM(Purchases.Deobligations))) AS Deobligations," _
            & " CCur(ABS(SUM(Purchases.ULO))) AS ULO, CCur(ABS(SUM(Purchases.Expenditures))) AS Expenditures" _
            & " FROM Allocations LEFT JOIN Purchases ON (Allocations.BFY = Purchases.BFY)" _
            & " AND (Allocations.AhCode = LEFT(Purchases.AhCode, 2)) AND (Allocations.FundCode = Purchases.FundCode)" _
            & " AND (Allocations.AccountCode = Purchases.AccountCode) AND (Allocations.BocCode = Purchases.BocCode)" _
            & " AND (Allocations.RcCode = Purchases.RcCode)" _
            & " GROUP BY Purchases.RpioCode, Purchases.AhCode, Purchases.BFY," _
            & " Purchases.RcCode, Purchases.FundCode, Purchases.FundName, Purchases.RcCode, Purchases.DivisionName, Purchases.DocumentControlNumbers," _
            & " Purchases.OriginalActionDate, Purchases.LastActionDate," _
            & " Purchases.OrgCode, Purchases.AccountCode, Purchases.ActivityCode," _
            & " Purchases.ProgramProjectName, Purchases.BocCode, Purchases.BocName," _
            & " Purchases.FocCode, Purchases.FocName" _
            & " AND Purchases.DivisionName = '" & Args.DivisionName & "'" _
            & " AND NOT IsNull(Purchases.DocumentControlNumbers);"
    End If
        Me.RecordSource = SQL
        Me.Requery
        DoCmd.Close ObjectType:=acForm, _
            ObjectName:="ReportDialog", Save:=acSaveNo
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



