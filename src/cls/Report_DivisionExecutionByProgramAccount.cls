VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionExecutionByProgramAccount"
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
    Set m_Dialog = New Form_ReportDialog
    Set Args = New AccountArgs
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = m_Dialog.Args
    RcCodeArgs = Args.RcCode
    FiscalYearArgs = Args.BFY
    PublicLawArgs = DLookup("PublicLaw", "Appropriations", "BFY = '" & FiscalYearArgs & "'")
    Dim SQL As String
    If Not Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            FundCodeArgs = Args.FundCode
            FundTitleArgs = DLookup("Title", "Funds", "Funds.Code = '" & FundCodeArgs & "'")
            TreasurySymbolArgs = DLookup("TreasurySymbol", "Funds", "Funds.Code = '" & FundCodeArgs & "'")
            SQL = "SELECT *" _
            & " FROM DivisionExecution" _
            & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
            & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'" _
            & " AND DivisionExecution.FundName = '" & Args.FundName & "';"
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundCode & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT *" _
            & " FROM DivisionExecution" _
            & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
            & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "';"
    End If
    If Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            FundCodeArgs = Args.FundCode
            FundTitleArgs = DLookup("Title", "Funds", "Funds.Code = '" & FundCodeArgs & "'")
            TreasurySymbolArgs = DLookup("TreasurySymbol", "Funds", "Funds.Code = '" & FundCodeArgs & "'")
            SQL = "SELECT *" _
            & " FROM DivisionExecution" _
            & " WHERE DivisionExecution.FundName = '" & Args.FundName & "'" _
            & " AND DivisionExecution.BFY = '" & Args.BFY & "';"
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




