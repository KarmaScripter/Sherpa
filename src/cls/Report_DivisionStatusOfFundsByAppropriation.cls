VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionStatusOfFundsByAppropriation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As AccountArgs
Public m_Dialog As Form_ReportDialog
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
    Dim SQL As String
    pAnd = " AND "
    Set Args = New AccountArgs
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = m_Dialog.Args
    If Not Args.FundName & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT DISTINCT DivisionExecution.BFY AS BFY, DivisionExecution.DivisionName AS DivisionName," _
                & " DivisionExecution.FundName AS FundName, Sum(CCur(DivisionExecution.Amount)) AS Authority, SUM(CCur(DivisionExecution.Used)) As Used," _
                & " SUM(CCur(DivisionExecution.Available)) AS Available" _
                & " FROM DivisionExecution" _
                & " WHERE DivisionExecution.DivisionName = '" & Args.DivisionName & "'" _
                & " AND DivisionExecution.BFY = '" & Args.BFY & "'" _
                & " AND DivisionExecution.FundName = '" & Args.FundName & "'" _
                & " GROUP BY DivisionExecution.BFY, DivisionExecution.DivisionName, DivisionExecution.FundName;"
    End If
    If Args.FundName & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT DISTINCT DivisionExecution.BFY AS BFY, DivisionExecution.DivisionName AS DivisionName," _
                & " DivisionExecution.FundName AS FundName, Sum(CCur(DivisionExecution.Amount)) AS Authority, SUM(CCur(DivisionExecution.Used)) As Used," _
                & " SUM(CCur(DivisionExecution.Available)) AS Available" _
                & " FROM DivisionExecution" _
                & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
                & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'" _
                & " GROUP BY DivisionExecution.BFY, DivisionExecution.DivisionName, DivisionExecution.FundName;"
    End If
    If Not Args.FundName & "" = "" And _
        Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT DISTINCT DivisionExecution.BFY AS BFY, DivisionExecution.DivisionName AS DivisionName," _
                & " DivisionExecution.FundName AS FundName, Sum(CCur(DivisionExecution.Amount)) AS Authority, SUM(CCur(DivisionExecution.Used)) As Used," _
                & " SUM(CCur(DivisionExecution.Available)) AS Available" _
                & " FROM DivisionExecution" _
                & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
                & " AND DivisionExecution.FundName = '" & Args.FundName & "'" _
                & " GROUP BY DivisionExecution.BFY, DivisionExecution.DivisionName, DivisionExecution.FundName;"
    End If
    If Args.FundName & "" = "" And _
        Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT DISTINCT DivisionExecution.BFY AS BFY, DivisionExecution.DivisionName AS DivisionName," _
                & " DivisionExecution.FundName AS FundName, Sum(CCur(DivisionExecution.Amount)) AS Authority, SUM(CCur(DivisionExecution.Used)) As Used," _
                & " SUM(CCur(DivisionExecution.Available)) AS Available" _
                & " From DivisionExecution" _
                & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
                & " GROUP BY DivisionExecution.BFY, DivisionExecution.DivisionName, DivisionExecution.FundName;"
    End If
    Me.RecordSource = SQL
    Me.Requery
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

