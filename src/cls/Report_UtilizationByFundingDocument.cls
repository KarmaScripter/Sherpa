VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_UtilizationByFundingDocument"
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
Private m_Notification As String





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Load()
    On Error GoTo ErrorHandler:
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:      ExpiringFundsUtilization" _
            & vbCrLf & "Member:     Report_Load()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (m_Error)
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
    Dim SQL As String
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = New AccountArgs
    Set Args = m_Dialog.Args
    RcCodeArgs = Args.RcCode
    FiscalYearArgs = Args.BFY
    DivisionTitleArgs = DLookup("Title", _
        "Divisions", "Divisions.Code = '" & RcCodeArgs & "'")
    PublicLawArgs = DLookup("PublicLaw", _
        "Appropriations", "BFY = '" & FiscalYearArgs & "'")
    If Not Args.FundName & "" = "" Then
        FundCodeArgs = Args.FundName
        FundTitleArgs = DLookup("Title", _
            "Funds", "Funds.Name = '" & FundCodeArgs & "'")
        TreasurySymbolArgs = DLookup("TreasurySymbol", _
            "Funds", "Funds.Name = '" & FundCodeArgs & "'")
        SQL = "SELECT *" _
        & " FROM Purchases" _
        & " WHERE Purchases.BFY = '" & Args.BFY & "'" _
        & " AND Purchases.DivisionName= '" & Args.DivisionName & "'" _
        & " AND Purchases.FundName = '" & Args.FundName & "';"
    End If
    If Args.FundName & "" = "" Then
        SQL = "SELECT *" _
        & " FROM Purchases" _
        & " WHERE Purchases.BFY = '" & Args.BFY & "'" _
        & " AND Purchases.DivisionName= '" & Args.DivisionName & "';"
    End If
    Me.RecordSource = SQL
    Me.Requery
    DoCmd.Close ObjectType:=acForm, ObjectName:="ReportDialog"
ErrorHandler:
    If Err.Number > 0 Then
         m_Error = "Source:      ExpiringFundsUtilization" _
            & vbCrLf & "Member:     Report_Open()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (m_Error)
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

