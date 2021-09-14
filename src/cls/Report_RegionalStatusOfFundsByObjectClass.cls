VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RegionalStatusOfFundsByObjectClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As AccountArgs
Public m_Dialog As Form_FundDialog
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
    pAnd = " AND "
    DoCmd.OpenForm FormName:="FundDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("FundDialog")
    Set Args = m_Dialog.Args
    Dim SQL As String
    If Not Args.FundName & "" = "" Then
        SQL = "SELECT *" _
            & " FROM SelectRegionalObjectClassExecution" _
            & " WHERE SelectRegionalObjectClassExecution.FundName = '" & Args.FundName & "'" _
            & " AND SelectRegionalObjectClassExecution.BFY = '" & Args.BFY & "'" _
            & " ORDER BY SelectRegionalObjectClassExecution.BFY DESC;"
    End If
    If Args.FundName & "" = "" Then
        SQL = "SELECT *" _
            & " FROM SelectRegionalObjectClassExecution" _
            & " WHERE SelectRegionalObjectClassExecution.BFY = '" & Args.BFY & "'" _
            & " ORDER BY SelectRegionalObjectClassExecution.BFY DESC;"
    End If
    Me.RecordSource = SQL
    Me.Requery
    DoCmd.Close ObjectType:=acForm, ObjectName:="FundDialog"
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

