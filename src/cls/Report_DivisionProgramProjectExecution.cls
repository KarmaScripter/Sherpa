VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionProgramProjectExecution"
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
Private pAnd As String
Private m_Error As String



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set Args = New AccountArgs
    Dim SQL As String
    pAnd = " AND "
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = m_Dialog.Args
    FiscalYearArgs = Args.BFY
    PublicLawArgs = DLookup("PublicLaw", "Appropriations", "BFY = '" & FiscalYearArgs & "'")
    If Not Args.DivisionName & "" = "" And _
        Not Args.FundCode & "" = "" And _
        Not Args.BFY & "" = "" Then
            RcCodeArgs = Args.RcCode
            DivisionTitleArgs = DLookup("Title", "Divisions", "Divisions.Code = '" & RcCodeArgs & "'")
            FundCodeArgs = Args.FundCode
            FundTitleArgs = DLookup("Title", "Funds", "Funds.Code = '" & Args.FundCode & "'")
            TreasurySymbolArgs = DLookup("TreasurySymbol", "Funds", "Funds.Code = '" & Args.FundCode & "'")
            SQL = "SELECT DivisionExecution.BFY, DivisionExecution.RcCode, DivisionExecution.DivisionName," _
                & " DivisionExecution.FundCode, DivisionExecution.FundName," _
                & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName, DivisionExecution.AccountCode," _
                & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName," _
                & " CCur(Sum(DivisionExecution.Amount)) AS Amount, CCur(Sum(DivisionExecution.OpenCommitments)) AS OpenCommitments," _
                & " CCur(Sum(DivisionExecution.Obligations)) AS Obligations, CCur(Sum(DivisionExecution.Used)) AS Used," _
                & " CCur(Sum(DivisionExecution.Available)) AS Available" _
                & " FROM DivisionExecution" _
                & " WHERE DivisionExecution.Amount > 0" _
                & " AND DivisionExecution.RcCode NOT IN ('06', '06R')" _
                & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'" _
                & " AND DivisionExecution.FundName = '" & Args.FundName & "'" _
                & " AND DivisionExecution.BFY = '" & Args.BFY & "'" _
                & " AND DivisionExecution.Available > 0" _
                & " GROUP BY DivisionExecution.BFY, DivisionExecution.RcCode, DivisionExecution.DivisionName," _
                & " DivisionExecution.FundCode, DivisionExecution.FundName," _
                & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName, DivisionExecution.AccountCode," _
                & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName" _
                & " ORDER BY DivisionExecution.BFY DESC , DivisionExecution.ProgramProjectCode;"
    End If
    If Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Not Args.FundCode & "" = "" Then
            SQL = "SELECT DivisionExecution.BFY, DivisionExecution.RcCode, DivisionExecution.DivisionName," _
                & " DivisionExecution.FundCode, DivisionExecution.FundName," _
                & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName, DivisionExecution.AccountCode," _
                & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName," _
                & " CCur(Sum(DivisionExecution.Amount)) AS Amount, CCur(Sum(DivisionExecution.OpenCommitments)) AS OpenCommitments," _
                & " CCur(Sum(DivisionExecution.Obligations)) AS Obligations, CCur(Sum(DivisionExecution.Used)) AS Used," _
                & " CCur(Sum(DivisionExecution.Available)) AS Available" _
                & " FROM DivisionExecution" _
                & " WHERE DivisionExecution.Amount > 0" _
                & " AND DivisionExecution.RcCode NOT IN ('06', '06R')" _
                & " AND DivisionExecution.FundName = '" & Args.FundName & "'" _
                & " AND DivisionExecution.BFY = '" & Args.BFY & "'" _
                & " AND DivisionExecution.Available > 0" _
                & " GROUP BY DivisionExecution.BFY, DivisionExecution.RcCode, DivisionExecution.DivisionName," _
                & " DivisionExecution.FundCode, DivisionExecution.FundName," _
                & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName, DivisionExecution.AccountCode," _
                & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName" _
                & " ORDER BY DivisionExecution.BFY DESC , DivisionExecution.ProgramProjectCode;"
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundCode & "" = "" And _
        Not Args.BFY & "" = "" Then
            RcCodeArgs = Args.RcCode
            DivisionTitleArgs = DLookup("Title", "Divisions", "Divisions.Code = '" & RcCodeArgs & "'")
            SQL = "SELECT DivisionExecution.BFY, DivisionExecution.RcCode, DivisionExecution.DivisionName," _
                & " DivisionExecution.FundCode, DivisionExecution.FundName," _
                & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName, DivisionExecution.AccountCode," _
                & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName," _
                & " CCur(Sum(DivisionExecution.Amount)) AS Amount, CCur(Sum(DivisionExecution.OpenCommitments)) AS OpenCommitments," _
                & " CCur(Sum(DivisionExecution.Obligations)) AS Obligations, CCur(Sum(DivisionExecution.Used)) AS Used," _
                & " CCur(Sum(DivisionExecution.Available)) AS Available" _
                & " FROM DivisionExecution" _
                & " WHERE DivisionExecution.Amount > 0" _
                & " AND DivisionExecution.RcCode NOT IN ('06', '06R')" _
                & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'" _
                & " AND DivisionExecution.BFY = '" & Args.BFY & "'" _
                & " AND DivisionExecution.Available > 0" _
                & " GROUP BY DivisionExecution.BFY, DivisionExecution.RcCode, DivisionExecution.DivisionName," _
                & " DivisionExecution.FundCode, DivisionExecution.FundName," _
                & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName, DivisionExecution.AccountCode," _
                & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName" _
                & " ORDER BY DivisionExecution.BFY DESC , DivisionExecution.ProgramProjectCode;"
    End If
    If Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Args.FundCode & "" = "" Then
            SQL = "SELECT DivisionExecution.BFY, DivisionExecution.RcCode, DivisionExecution.DivisionName," _
                & " DivisionExecution.FundCode, DivisionExecution.FundName," _
                & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName, DivisionExecution.AccountCode," _
                & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName," _
                & " CCur(Sum(DivisionExecution.Amount)) AS Amount, CCur(Sum(DivisionExecution.OpenCommitments)) AS OpenCommitments," _
                & " CCur(Sum(DivisionExecution.Obligations)) AS Obligations, CCur(Sum(DivisionExecution.Used)) AS Used," _
                & " CCur(Sum(DivisionExecution.Available)) AS Available" _
                & " FROM DivisionExecution" _
                & " WHERE DivisionExecution.Amount > 0" _
                & " AND DivisionExecution.RcCode NOT IN ('06', '06R')" _
                & " AND DivisionExecution.BFY = '" & Args.BFY & "'" _
                & " AND DivisionExecution.Available > 0" _
                & " GROUP BY DivisionExecution.BFY, DivisionExecution.RcCode, DivisionExecution.DivisionName," _
                & " DivisionExecution.FundCode, DivisionExecution.FundName," _
                & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName, DivisionExecution.AccountCode," _
                & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName" _
                & " ORDER BY DivisionExecution.BFY DESC , DivisionExecution.ProgramProjectCode;"
    End If
    Me.RecordSource = SQL
    Me.Requery
    DoCmd.Close ObjectType:=acForm, ObjectName:="ReportDialog"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



