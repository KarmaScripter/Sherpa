VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DataWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private m_Error As String
Private m_Notification As String




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BudgetDetailedSpendingButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close ObjectType:=acForm, ObjectName:="DataWarehouse", Save:=acSaveNo
    DoCmd.OpenForm FormName:="BudgetDetailedSpendingQuery", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            BudgetStatusQueryButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BudgetStatusQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close ObjectType:=acForm, ObjectName:="DataWarehouse", Save:=acSaveNo
    DoCmd.OpenForm FormName:="CompassBudgetStatusQuery", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            CompassAdvancedButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CompassAdvancedButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close ObjectType:=acForm, ObjectName:="DataWarehouse", Save:=acSaveNo
    DoCmd.OpenForm FormName:="CompassAdvancedQuery", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            PayrollAcrrualButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayrollAcrrualButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close ObjectType:=acForm, ObjectName:="DataWarehouse", Save:=acSaveNo
    DoCmd.OpenForm FormName:="PayrollAccrualQuery", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ReprogrammingButton_Click()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ReprogrammingButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close ObjectType:=acForm, ObjectName:="DataWarehouse", Save:=acSaveNo
    DoCmd.OpenForm FormName:="CompassReprogrammingQuery", WindowMode:=acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub TravelButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close ObjectType:=acForm, ObjectName:="DataWarehouse", Save:=acSaveNo
    DoCmd.OpenForm FormName:="CompassTravelQuery", WindowMode:=acDialog
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





