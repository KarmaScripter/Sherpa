VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_DivisionStateTribalGrantExecution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Args As AccountArgs
Private m_Dialog As Form_ReportDialog
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
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      DivisionStateTribalGrantExecution" _
            & vbCrLf & "Member:     Report_Open()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    DoCmd.OpenForm FormName:="ReportDialog", WindowMode:=acDialog
    Set m_Dialog = New Form_ReportDialog
    Set m_Dialog = Forms("ReportDialog")
    Set Args = New AccountArgs
    Set Args = m_Dialog.Args
    If Not Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            SQL = "SELECT *" _
            & " FROM DivisionExecution" _
            & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
            & " AND LEFT(DivisionExecution.FundCode, 1) = 'E'" _
            & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'"
    End If
    If Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
        SQL = "SELECT *" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
        & " AND LEFT(DivisionExecution.FundCode, 1) = 'E'"
    End If
    If Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
        SQL = "SELECT *" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
        & " AND LEFT(DivisionExecution.FundCode, 1) = 'E'"
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
        SQL = "SELECT *" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BFY = '" & Args.BFY & "'" _
        & " AND LEFT(DivisionExecution.FundCode, 1) = 'E'" _
        & " AND DivisionExecution.DivisionName = '" & Args.DivisionName & "'"
    End If
    Me.RecordSource = SQL
    DoCmd.Close ObjectType:=acForm, ObjectName:="ReportDialog"
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      DivisionStateTribalGrantExecution" _
            & vbCrLf & "Member: Report_Open()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



