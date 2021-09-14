VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_PayrollAcrrualsByPayPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As PayrollArgs
Public m_Dialog As Form_PayPeriodDialog
Private m_StartDate As Date
Private m_EndDate As Date
Private m_Error As String




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Load()
    On Error GoTo ErrorHandler:
    Set m_Dialog = New Form_PayPeriodDialog
    Set Args = New PayrollArgs
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
    DoCmd.OpenForm FormName:="PayPeriodDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("PayPeriodDialog")
    Set Args = m_Dialog.Args
    Dim SQL As String
    If Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Not Args.PayPeriod & "" = "" Then
            SQL = "SELECT *" _
            & " FROM PayrollAccruals" _
            & " WHERE PayrollAccruals.PayPeriod = '" & Args.PayPeriod & "'" _
            & " AND PayrollAccruals.BFY = '" & Args.BFY & "'" _
            & " AND PayrollAccruals.FundName = '" & Args.FundName & "'" _
            & " AND PayrollAccruals.DivisionName = '" & Args.DivisionName & "'" _
            & " ORDER BY PayrollAccruals.BFY DESC, PayrollAccruals.FundName;"
    End If
    If Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Not Args.PayPeriod & "" = "" Then
            SQL = "SELECT *" _
            & " FROM PayrollAccruals" _
            & " WHERE PayrollAccruals.PayPeriod = '" & Args.PayPeriod & "'" _
            & " AND PayrollAccruals.BFY = '" & Args.BFY & "'" _
            & " AND PayrollAccruals.DivisionName = '" & Args.DivisionName & "'" _
            & " ORDER BY PayrollAccruals.BFY DESC, PayrollAccruals.FundName;"
    End If
    If Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Not Args.DivisionName & "" = "" And _
        Args.PayPeriod & "" = "" Then
            SQL = "SELECT *" _
            & " FROM PayrollAccruals" _
            & " WHERE PayrollAccruals.FundName = '" & Args.FundName & "'" _
            & " AND PayrollAccruals.BFY = '" & Args.BFY & "'" _
            & " AND PayrollAccruals.DivisionName = '" & Args.DivisionName & "'" _
            & " ORDER BY PayrollAccruals.BFY DESC, PayrollAccruals.FundName;"
    End If
    If Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" And _
        Args.DivisionName & "" = "" And _
        Not Args.PayPeriod & "" = "" Then
            SQL = "SELECT *" _
            & " FROM PayrollAccruals" _
            & " WHERE PayrollAccruals.PayPeriod = '" & Args.PayPeriod & "'" _
            & " AND PayrollAccruals.BFY = '" & Args.BFY & "'" _
            & " ORDER BY PayrollAccruals.BFY DESC, PayrollAccruals.FundName;"
    End If
    If Args.FundName & "" = "" And _
        Args.BFY & "" = "" And _
        Args.DivisionName & "" = "" And _
        Args.PayPeriod & "" = "" Then
            SQL = "SELECT *" _
            & " FROM PayrollAccruals" _
            & " WHERE PayrollAccruals.FundName = '" & Args.FundName & "'" _
            & " AND PayrollAccruals.BFY = '2021'" _
            & " WHERE PayrollAccruals.PayPeriod = '202101'" _
            & " AND PayrollAccruals.DivisionName = '" & Args.DivisionName & "'" _
            & " ORDER BY PayrollAccruals.BFY DESC, PayrollAccruals.FundName;"
    End If
    Me.RecordSource = SQL
    Me.Requery
    DoCmd.Close ObjectType:=acForm, ObjectName:=m_Dialog.Name
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

