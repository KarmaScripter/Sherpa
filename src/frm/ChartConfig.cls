Option Compare Database
Option Explicit





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        cmdUpdateChart_Click()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdUpdateChart_Click()
    If IsFormOpen("BudgetChart") = False Then
        DoCmd.OpenForm "BudgetChart", acNormal
    End If
    
    Forms!BudgetChart.SetFocus
    DoEvents
    Forms!BudgetChart.cmdCreateChart_Click
    Me.Visible = False
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        fraChartType_AfterUpdate()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub fraChartType_AfterUpdate()
    Select Case Me.fraChartType
        Case chChartType.Line, chChartType.Bar, chChartType.HorizontalBar, chChartType.Bubble, chChartType.Scatter
            Me.chkShowYAxis = True
            Me.chkShowXAxis = True
        Case chChartType.Radar, chChartType.Doughnut, chChartType.Pie, chChartType.PolarArea
            Me.chkShowYAxis = False
            Me.chkShowXAxis = False
        Case Else
    End Select
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        cmdSaveChartJs_Click()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdSaveChartJs_Click()
    If SaveChartjsToDisk = True Then
        MsgBox "Chart.js successfully saved to hdd.", vbInformation, "Budget Execution"
    Else
        MsgBox "Something went wrong!", vbExclamation, "Budget Execution"
    End If
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        cmdSaveColorSchema_Click()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdSaveColorSchema_Click()
    If SaveChartjsPluginColorSchemesToDisk = True Then
        MsgBox "Files related to color schemes plug-in " & vbCrLf & _
               "successfully saved to hdd.", vbInformation, "Budget Execution"
    Else
        MsgBox "Something went wrong!", vbExclamation, "Budget Execution"
    End If
End Sub




