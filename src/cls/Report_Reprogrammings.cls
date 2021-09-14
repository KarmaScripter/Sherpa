VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Reprogrammings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As BudgetArgs
Public m_Dialog As Form_TimeDialog
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
    Set m_Dialog = New Form_TimeDialog
    Set Args = New BudgetArgs
    DoCmd.OpenForm FormName:="TimeDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("TimeDialog")
    Set Args = m_Dialog.Args
    Dim SQL As String
    If Not Args.FundName & "" = "" And _
        Not Args.BocName & "" = "" And _
        Not Args.StartDate & "" = "" And _
        Not Args.EndDate & "" = "" Then
            SQL = "SELECT *" _
            & " FROM Reprogrammings" _
            & " WHERE Reprogrammings.FundName = '" & Args.FundName & "'" _
            & " AND Reprogrammings.BocName = '" & Args.BocName & "'" _
            & " AND Reprogrammings.ProcessedDate >= #" & Args.StartDate & "#" _
            & " AND Reprogrammings.ProcessedDate <= #" & Args.EndDate & "#;"
    End If
    If Not Args.FundName & "" = "" And _
        Args.BocName & "" = "" And _
        Not Args.StartDate & "" = "" And _
        Not Args.EndDate & "" = "" Then
            SQL = "SELECT *" _
            & " FROM Reprogrammings" _
            & " WHERE Reprogrammings.FundName = '" & Args.FundName & "'" _
            & " AND Reprogrammings.ProcessedDate >= #" & Args.StartDate & "#" _
            & " AND Reprogrammings.ProcessedDate <= #" & Args.EndDate & "#;"
    End If
    If Args.FundName & "" = "" And _
        Args.BocName & "" = "" And _
        Not Args.StartDate & "" = "" And _
        Not Args.EndDate & "" = "" Then
            SQL = "SELECT *" _
            & " FROM Reprogrammings" _
            & " WHERE Reprogrammings.ProcessedDate >= #" & Args.StartDate & "#" _
            & " AND Reprogrammings.ProcessedDate <= #" & Args.EndDate & "#;"
    End If
    If Not Args.FundName & "" = "" And _
        Not Args.BocName & "" = "" And _
        Args.StartDate & "" = "" And _
        Not Args.EndDate & "" = "" Then
            SQL = "SELECT *" _
            & " FROM Reprogrammings" _
            & " WHERE Reprogrammings.FundName = '" & Args.FundName & "'" _
            & " AND Reprogrammings.BocName = '" & Args.BocName & "'" _
            & " AND Reprogrammings.ProcessedDate <= #" & Args.EndDate & "#;"
    End If
    If Not Args.FundName & "" = "" And _
        Not Args.BocName & "" = "" And _
        Not Args.StartDate & "" = "" And _
        Args.EndDate & "" = "" Then
            SQL = "SELECT *" _
            & " FROM Reprogrammings" _
            & " WHERE Reprogrammings.FundName = '" & Args.FundName & "'" _
            & " AND Reprogrammings.BocName = '" & Args.BocName & "'" _
            & " AND Reprogrammings.ProcessedDate >= #" & Args.StartDate & "#;"
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

