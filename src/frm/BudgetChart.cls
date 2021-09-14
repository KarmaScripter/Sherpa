Option Compare Database
Option Explicit





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    'Me.lblVersion.Caption = "Version: " & DMax("V_Number", "tbl_VersionHistory")
        If IsFormOpen("ChartConfig") = True Then
            Forms!ChartConfig.Visible = False
        End If
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub cmdCreateChart_Click()



    'Declarations
    Dim mChart As Chart
    Dim rst As DAO.Recordset
    
    'Make sure ChartConfig is open
    If IsFormOpen("ChartConfig") = False Then
        DoCmd.OpenForm "ChartConfig", acNormal
        If Forms!ChartConfig.Visible = False Then
            Forms!ChartConfig.Visible = True
        End If
        Forms!ChartConfig.SetFocus
    End If
    If IsFormOpen("ChartConfig") = True And Forms!ChartConfig.Visible = False Then
        Forms!ChartConfig.Visible = True
    End If
    
    'Create a new instance
    Set mChart = New Chart
    
    'Connect the web browser control to the class module
    Set mChart.control = Me.ctlWebbrowser
    
    'Define type of chart
    mChart.ChartType = Forms!ChartConfig.fraChartType
    mChart.StackedChart = Forms!ChartConfig.chkStackedChart
    mChart.AreaFilled = Forms!ChartConfig.chkAreaFilled
    
    'Define datasource
    Select Case Forms!ChartConfig.fraDataSource
        Case 1  'Table
            mChart.DataSourceObjectName = "tbl_DemoData"
        Case 2  'Query
            mChart.DataSourceObjectName = "qry_DemoData"
        Case 3  'SQL-Statement
            mChart.DataSourceSQLStatement = "Select * from tbl_DemoData order by pkid desc"
        Case 4  'Recordset
            Set rst = CurrentDb.OpenRecordset("Select * from tbl_DemoData order by pkid desc")
            Set mChart.DataSourceRecordset = rst
        Case 5  'None
            'Do Nothing
    End Select
    mChart.DataSourceLabelFieldName = "DataLabel"
    mChart.DataSourceDataFieldName = Array("Dataset1", "Dataset2", "Dataset3")
    
    'Define title
    mChart.ShowTitle = Forms!ChartConfig.chkShowTitle
    mChart.TitlePostition = Forms!ChartConfig.fraTitlePosition
    mChart.TitleText = Forms!ChartConfig.txtTitleText
    
    'Define legend
    mChart.ShowLegend = Forms!ChartConfig.chkShowLegend
    mChart.LegendPostition = Forms!ChartConfig.fraLegendPosition
    
    'Define axis
    mChart.ShowYAxis = Forms!ChartConfig.chkShowYAxis
    mChart.ShowXAxis = Forms!ChartConfig.chkShowXAxis
    mChart.YAxisLabelText = Nz(Forms!ChartConfig.txtLabelYAxis, vbNullString)
    mChart.XAxisLabelText = Nz(Forms!ChartConfig.txtLabelXAxis, vbNullString)
    mChart.YAxisBeginAtZero = Nz(Forms!ChartConfig.chkYAxisBeginAtZero, False)
    mChart.XAxisBeginAtZero = Nz(Forms!ChartConfig.chkXAxisBeginAtZero, False)
    
    'Define chart.js
    mChart.ChartJsSource = Forms!ChartConfig.cboChartJsSource
    mChart.ChartJsCDN = Forms!ChartConfig.txtChartJsCDN
    mChart.ChartJsPathFilename = Forms!ChartConfig.txtChartJsPathFilename
    
    'define color scheme plug-in
    mChart.ColorSchemaSource = Forms!ChartConfig.cboColorSchemaSource
    mChart.ColorSchemaCDN = Forms!ChartConfig.txtColorSchemaCDN
    mChart.ColorSchemaPathFilename = Forms!ChartConfig.txtColorSchemaPathFilename
    
    'Color schema
    mChart.ColorSchema = Forms!ChartConfig.cboChartColorSchema
    mChart.ChartBackgroundColor = Nz(Forms!ChartConfig.txtChartBackgroundColor, "#EAEAEA")
    
    'Default font
    mChart.DefaultFontColor = Nz(Forms!ChartConfig.txtDefaultFontColor, "Black")
    mChart.DefaultFontFamily = Nz(Forms!ChartConfig.txtDefaultFontFamily, vbNullString)
    mChart.DefaultFontSize = Nz(Forms!ChartConfig.txtDefaultFontSize, 0)
    
    'Draw chart
    mChart.DrawChart

End Sub


