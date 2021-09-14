VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ChartConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_Error As String


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



