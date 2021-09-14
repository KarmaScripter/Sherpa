VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionTravelAllocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Args As AccountArgs
Public m_Dialog As Form_ReportDialog
Public FundCodeArgs As String
Public FundNameArgs As String
Public RcCodeArgs As String
Public FiscalYearArgs As String
Public PublicLawArgs As String
Public FundTitleArgs As String
Public TreasurySymbolArgs As String
Public DivisionNameArgs As String
Private m_Error As String



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = New Form_ReportDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = New AccountArgs
    Set Args = m_Dialog.Args
    Dim SQL As String
    If Not Args.FundName & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT  DivisionExecution.BFY, DivisionExecution.AhCode, DivisionExecution.RcCode," _
            & " DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.FundName," _
            & " DivisionExecution.OrgCode, DivisionExecution.AccountCode,DivisionExecution.ProgramAreaName, DivisionExecution.ProgramProjectName," _
            & " Sum(DivisionExecution.Amount) AS Amount, Sum(DivisionExecution.OpenCommitments) AS OpenCommitments," _
            & " Sum(DivisionExecution.Obligations) AS Obligations, Sum(DivisionExecution.Used) AS Used, Sum(DivisionExecution.Available) AS Available" _
            & " FROM DivisionExecution" _
            & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
            & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'" _
            & " AND DivisionExecution.FundName = '" & Args.FundName & "'" _
            & " AND DivisionExecution.Boccode IN ('21', '28')" _
            & " GROUP BY DivisionExecution.BFY, DivisionExecution.AhCode, DivisionExecution.RcCode," _
            & " DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.FundName," _
            & " DivisionExecution.OrgCode, DivisionExecution.AccountCode, DivisionExecution.ProgramProjectName;"
    End If
    If Not Args.FundName & "" = "" And _
        Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT  DivisionExecution.BFY, DivisionExecution.AhCode, DivisionExecution.RcCode," _
            & " DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.FundName," _
            & " DivisionExecution.OrgCode, DivisionExecution.AccountCode,DivisionExecution.ProgramAreaName, DivisionExecution.ProgramProjectName," _
            & " Sum(DivisionExecution.Amount) AS Amount, Sum(DivisionExecution.OpenCommitments) AS OpenCommitments," _
            & " Sum(DivisionExecution.Obligations) AS Obligations, Sum(DivisionExecution.Used) AS Used, Sum(DivisionExecution.Available) AS Available" _
            & " FROM DivisionExecution" _
            & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
            & " AND DivisionExecution.FundName = '" & Args.FundName & "'" _
            & " AND DivisionExecution.Boccode IN ('21', '28')" _
            & " GROUP BY DivisionExecution.BFY, DivisionExecution.AhCode, DivisionExecution.RcCode," _
            & " DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.FundName," _
            & " DivisionExecution.OrgCode, DivisionExecution.ProgramAreaName, DivisionExecution.ProgramProjectName, DivisionExecution.AccountCode;"
    End If
    If Args.FundName & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
        SQL = "SELECT  DivisionExecution.BFY, DivisionExecution.AhCode, DivisionExecution.RcCode," _
        & " DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.FundName," _
        & " DivisionExecution.OrgCode, DivisionExecution.AccountCode,DivisionExecution.ProgramAreaName, DivisionExecution.ProgramProjectName," _
        & " Sum(DivisionExecution.Amount) AS Amount, Sum(DivisionExecution.OpenCommitments) AS OpenCommitments," _
        & " Sum(DivisionExecution.Obligations) AS Obligations, Sum(DivisionExecution.Used) AS Used, Sum(DivisionExecution.Available) AS Available" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
        & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'" _
        & " AND DivisionExecution.Boccode IN ('21', '28')" _
        & " GROUP BY DivisionExecution.BFY, DivisionExecution.AhCode, DivisionExecution.RcCode," _
        & " DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.FundName," _
        & " DivisionExecution.OrgCode, DivisionExecution.ProgramAreaName, DivisionExecution.ProgramProjectName, DivisionExecution.AccountCode;"
    End If
    If Args.FundName & "" = "" And _
        Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
        SQL = "SELECT  DivisionExecution.BFY, DivisionExecution.AhCode, DivisionExecution.RcCode," _
        & " DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.FundName," _
        & " DivisionExecution.OrgCode, DivisionExecution.AccountCode, DivisionExecution.ProgramAreaName, DivisionExecution.ProgramProjectName," _
        & " Sum(DivisionExecution.Amount) AS Amount, Sum(DivisionExecution.OpenCommitments) AS OpenCommitments," _
        & " Sum(DivisionExecution.Obligations) AS Obligations, Sum(DivisionExecution.Used) AS Used, Sum(DivisionExecution.Available) AS Available" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
        & " AND DivisionExecution.Boccode IN ('21', '28')" _
        & " GROUP BY DivisionExecution.BFY, DivisionExecution.AhCode, DivisionExecution.RcCode," _
        & " DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.FundName," _
        & " DivisionExecution.OrgCode, DivisionExecution.ProgramAreaName, DivisionExecution.ProgramProjectName, DivisionExecution.AccountCode;"
    End If
    Me.RecordSource = SQL
    DoCmd.Close ObjectType:=acForm, ObjectName:="ReportDialog"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Load()
    On Error GoTo ErrorHandler:
    If Args.DivisionName & "" = "" Then
        Me.RcCodeTextBox.SetFocus
        Me.RcCodeTextBox.Value = "R6 ALL"
    Else
        Me.RcCodeTextBox.SetFocus
        Me.RcCodeTextBox.Value = Args.RcCode
    End If
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




