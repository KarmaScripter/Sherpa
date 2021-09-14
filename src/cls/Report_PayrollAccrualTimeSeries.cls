VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_PayrollAccrualTimeSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As PayrollArgs
Public m_Dialog As Form_PayrollDialog
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
    Set m_Dialog = New Form_PayrollDialog
    Set Args = New PayrollArgs
    DoCmd.OpenForm FormName:="PayrollDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("PayrollDialog")
    Set Args = m_Dialog.Args
    Dim SQL As String
    If Not Args.FundName & "" = "" Then
        SQL = "SELECT *" _
        & " FROM PayrollAccruals" _
        & " WHERE PayrollAccruals.DivisionName = '" & Args.DivisionName & "'" _
        & " AND PayrollAccruals.FundName = '" & Args.FundName & "'" _
        & " AND PayrollAccruals.StartDate >= #" & Args.StartDate & "#" _
        & " AND PayrollAccruals.EndDate <= #" & Args.EndDate & "#;"
    End If
    If Args.FundName & "" = "" Then
        SQL = "SELECT *" _
        & " FROM PayrollAccruals" _
        & " WHERE PayrollAccruals.DivisionName = '" & Args.DivisionName & "'" _
        & " AND PayrollAccruals.StartDate >= #" & Args.StartDate & "#" _
        & " AND PayrollAccruals.EndDate <= #" & Args.EndDate & "#;"
    End If
    Me.RecordSource = SQL
    Me.Requery
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name
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



