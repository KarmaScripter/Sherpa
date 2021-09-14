VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionExecutionTimeSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As AccountArgs
Public m_Dialog As Form_ReportDialog
Private FundCodeArgs As String
Private FundNameArgs As String
Private RcCodeArgs As String
Private FiscalYearArgs As String
Private DivisionNameArgs As String
Private AhCodeArgs As String
Private RpioCodeArgs As String
Private PublicLawArgs As String
Private FundTitleArgs As String
Public TreasurySymbolArgs As String
Private m_Database As DAO.Database
Private m_DataSet As DAO.Recordset
Private m_Query As DAO.QueryDef
Private m_Error As String




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Load()
    On Error GoTo ErrorHandler:
    Me.DivisionTextBox.SetFocus
    Me.DivisionTextBox.text = Args.DivisionName
    Me.FundTitleTextBox.SetFocus
    Me.FundTitleTextBox.text = FundTitleArgs
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
    Set Args = New AccountArgs
    Dim SQL As String
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = m_Dialog.Args
    DivisionNameArgs = Args.DivisionName
    FiscalYearArgs = Args.BFY
    RcCodeArgs = DLookup("Code", "Divisions", "Divisions.Title = '" & Args.DivisionName & "'")
    PublicLawArgs = DLookup("PublicLaw", "Appropriations", "BFY = '" & FiscalYearArgs & "'")
    If Not Args.FundCode & "" = "" Then
        FundCodeArgs = Args.FundCode
        FundNameArgs = Args.FundName
        FundTitleArgs = DLookup("Title", "Funds", "Funds.Code = '" & Args.FundCode & "'")
        TreasurySymbolArgs = DLookup("TreasurySymbol", "Funds", "Funds.Code = '" & Args.FundCode & "'")
        SQL = "SELECT *" _
            & " FROM Purchases" _
            & " WHERE Purchases.BFY = '" & FiscalYearArgs & "' AND" _
            & " Purchases.RcCode = '" & RcCodeArgs & "' AND" _
            & " Purchases.FundCode = '" & Args.FundCode & "' AND" _
            & " Purchases.BocCode NOT IN ('10', '17', '21', '38')" _
            & " ORDER BY Purchases.OriginalActionDate;"
    End If
    If Args.FundCode & "" = "" Then
        SQL = "SELECT *" _
            & " FROM Purchases" _
            & " WHERE Purchases.BFY = '" & FiscalYearArgs & "' AND" _
            & " Purchases.RcCode = '" & RcCodeArgs & "' AND" _
            & " Purchases.BocCode NOT IN ('10', '17', '21', '38')" _
            & " ORDER BY Purchases.OriginalActionDate;"
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




