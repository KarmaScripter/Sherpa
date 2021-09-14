VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionProgrammaticAuthority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Args As AccountArgs
Public m_Dialog As Form_ReportDialog
Public AhCodeArgs As String
Public RpioCodeArgs As String
Public PublicLawArgs As String
Public FundTitleArgs As String
Public TreasurySymbolArgs As String
Private m_Error As String





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Integer
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Dim SQL As String
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = New Form_ReportDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = New AccountArgs
    Set Args = m_Dialog.Args
    If Not Args.FundName & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
            Args.DivisionName = DLookup("Title", "Divisions", _
                "Divisions.Code = '" & Args.RcCode & "'")
            FundTitleArgs = DLookup("Title", "Funds", _
                "Funds.Code = '" & Args.FundCode & "'")
            TreasurySymbolArgs = DLookup("TreasurySymbol", "Funds", _
                "Funds.Code = '" & Args.FundCode & "'")
            PublicLawArgs = DLookup("PublicLaw", "Appropriations", _
                "BFY = '" & Args.BFY & "'")
            SQL = "SELECT *" _
            & " FROM DivisionAuthority" _
            & " WHERE DivisionAuthority.BFY = '" & Args.BFY & "'" _
            & " AND DivisionAuthority.DivisionName = '" & Args.DivisionName & "'" _
            & " AND DivisionAuthority.FundName = '" & Args.FundName & "'" _
            & " AND DivisionAuthority.Amount > 0;"
    End If
    If Not Args.FundName & "" = "" And _
        Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
        FundTitleArgs = DLookup("Title", "Funds", _
            "Funds.Code = '" & Args.FundCode & "'")
        TreasurySymbolArgs = DLookup("TreasurySymbol", "Funds", _
            "Funds.Code = '" & Args.FundCode & "'")
        PublicLawArgs = DLookup("PublicLaw", "Appropriations", _
            "BFY = '" & Args.BFY & "'")
        SQL = "SELECT *" _
        & " FROM DivisionAuthority" _
        & " WHERE DivisionAuthority.BFY = '" & Args.BFY & "'" _
        & " AND DivisionAuthority.FundName = '" & Args.FundName & "'" _
        & " AND DivisionAuthority.Amount > 0;"
    End If
    If Args.FundName & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
            PublicLawArgs = DLookup("PublicLaw", "Appropriations", _
                "BFY = '" & Args.BFY & "'")
            SQL = "SELECT *" _
            & " FROM DivisionAuthority" _
            & " WHERE DivisionAuthority.BFY = '" & Args.BFY & "'" _
            & " AND DivisionAuthority.DivisionName = '" & Args.DivisionName & "'" _
            & " AND DivisionAuthority.Amount > 0;"
    End If
    If Args.FundName & "" = "" And _
        Args.DivisionName & "" = "" And _
        Not Args.BFY & "" = "" Then
        PublicLawArgs = DLookup("PublicLaw", "Appropriations", _
            "BFY = '" & Args.BFY & "'")
        SQL = "SELECT *" _
        & " FROM DivisionAuthority" _
        & " WHERE DivisionAuthority.BFY = '" & Args.BFY & "'" _
        & " AND DivisionAuthority.Amount > 0;"
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




