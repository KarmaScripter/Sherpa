VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_PayrollStatusOfFundsByProgramArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------   FIELDS   ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------

Public Args As PayrollArgs
Public m_Dialog As Form_PayrollStatusDialog
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
Private mNotification As String


'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------   METHODS  ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Retval:      Void
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
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Dim SQL As String
    DoCmd.OpenForm FormName:="PayrollStatusDialog", WindowMode:=acDialog
    Set m_Dialog = New Form_PayrollStatusDialog
    Set m_Dialog = Forms("PayrollStatusDialog")
     m_Dialog.FundNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.FundName" _
        & " WHERE StatusOfFunds.BocCode = '10'" _
        & " AND StatusOfFunds.BudgetLevel = '7'"
    Set Args = New PayrollArgs
    Set Args = m_Dialog.Args
    If Not Args.FundName = "" Then
            SQL = "SELECT *" _
            & " FROM StatusOfFunds" _
            & " WHERE StatusOfFunds.BudgetLevel = '7'" _
            & " AND StatusOfFunds.FundName = '" & Args.FundName & "'" _
            & " AND StatusOfFunds.BFY = '" & Args.BFY & "'" _
            & " AND StatusOfFunds.BocCode = '10'"
    End If
    If Args.FundName & "" = "" Then
            SQL = "SELECT *" _
            & " FROM StatusOfFunds" _
            & " WHERE StatusOfFunds.BudgetLevel = '7'" _
            & " AND StatusOfFunds.BFY = '" & Args.BFY & "'" _
            & " AND StatusOfFunds.BocCode = '10'"
    End If
    Me.RecordSource = SQL
    Me.PayUtilizationChart.RowSource = Me.RecordSource
    Me.Requery
    DoCmd.Close ObjectType:=acForm, ObjectName:="PayrollStatusDialog"
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

