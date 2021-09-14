VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionAppropriations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As AccountArgs
Public m_Dialog As Form_ExpiringDialog
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
Private m_Notification As String



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Load()
    On Error GoTo ErrorHandler:
    Set Args = New AccountArgs
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
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm FormName:="ExpiringDialog", WindowMode:=acDialog
    Set m_Dialog = New Form_ExpiringDialog
    Set m_Dialog = Forms("ExpiringDialog")
    Set Args = m_Dialog.Args
    Dim SQL As String
    If Not Args.FundCode & "" = "" Then
        SQL = "SELECT DivisionExecution.BFY, DivisionExecution.AhCode," _
        & " DivisionExecution.FundCode, DivisionExecution.FundName, DivisionExecution.OrgCode, DivisionExecution.AccountCode," _
        & " DivisionExecution.BocCode, DivisionExecution.BocName," _
        & " DivisionExecution.RcCode, DivisionExecution.DivisionName," _
        & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName," _
        & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName," _
        & " SUM(DivisionExecution.Amount) AS Amount, SUM(DivisionExecution.OpenCommitments) As OpenCommitments," _
        & " SUM(DivisionExecution.Obligations) As Obligations, SUM(DivisionExecution.Used) AS Used, Sum(DivisionExecution.Available) AS Available" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.FundCode = '" & Args.FundCode & "'" _
        & " AND DivisionExecution.Available > 0" _
        & " AND DivisionExecution.BFY = '" & Args.BFY & "'" _
        & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'" _
        & " GROUP BY DivisionExecution.BFY, DivisionExecution.AhCode," _
        & " DivisionExecution.FundCode, DivisionExecution.FundName, DivisionExecution.OrgCode, DivisionExecution.AccountCode," _
        & " DivisionExecution.BocCode, DivisionExecution.BocName," _
        & " DivisionExecution.RcCode, DivisionExecution.DivisionName," _
        & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName," _
        & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName;"
    End If
    If Args.FundCode & "" = "" Then
        SQL = "SELECT DivisionExecution.BFY, DivisionExecution.AhCode," _
        & " DivisionExecution.FundCode, DivisionExecution.FundName, DivisionExecution.OrgCode, DivisionExecution.AccountCode," _
        & " DivisionExecution.BocCode, DivisionExecution.BocName," _
        & " DivisionExecution.RcCode, DivisionExecution.DivisionName," _
        & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName," _
        & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName," _
        & " SUM(DivisionExecution.Amount) AS Amount, SUM(DivisionExecution.OpenCommitments) As OpenCommitments," _
        & " SUM(DivisionExecution.Obligations) As Obligations, SUM(DivisionExecution.Used) AS Used, Sum(DivisionExecution.Available) AS Available" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.Available > 0" _
        & " AND DivisionExecution.BFY = '" & Args.BFY & "'" _
        & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'" _
        & " GROUP BY DivisionExecution.BFY, DivisionExecution.AhCode," _
        & " DivisionExecution.FundCode, DivisionExecution.FundName, DivisionExecution.OrgCode, DivisionExecution.AccountCode," _
        & " DivisionExecution.BocCode, DivisionExecution.BocName," _
        & " DivisionExecution.RcCode, DivisionExecution.DivisionName," _
        & " DivisionExecution.ProgramProjectCode, DivisionExecution.ProgramProjectName," _
        & " DivisionExecution.ProgramAreaCode, DivisionExecution.ProgramAreaName;"
    End If
    Me.RecordSource = SQL
    Me.Requery
    DoCmd.Close ObjectType:=acForm, ObjectName:=mDialog.Name, Save:=acSaveNo
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




