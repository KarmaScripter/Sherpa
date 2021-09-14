VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private m_Error As String

'### Enums



'----------------------------------------------------------------------------------
'   Type        Enum
'   Name        chChartType
'   Parameters
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Enum chChartType
    Line = 1
    Bar = 2
    HorizontalBar = 3
    Radar = 4
    Doughnut = 5
    Pie = 6
    PolarArea = 7
    Bubble = 8
    Scatter = 9
End Enum


'----------------------------------------------------------------------------------
'   Type        Enum
'   Name        chChartType
'   Parameters
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Enum chPosition
    posTop = 1
    posLeft = 2
    posBottom = 3
    posRight = 4
End Enum


'----------------------------------------------------------------------------------
'   Type        Enum
'   Name        chChartType
'   Parameters
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Enum chScriptSource
    CDN = 1
    LocalFile = 2
End Enum

'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------   FIELDS  -------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------

Private Const HtmlFileName As String = "BudgetExecutionChart.html"
Private WithEvents mControl As WebbrowserControl
Attribute mControl.VB_VarHelpID = -1
Private mChartType As chChartType
Private mDataSourceObjectName As String
Private mDataSourceSQLStatement As String
Private mDataSourceRecordset As DAO.Recordset
Private mDataSourceLabelFieldName As String
Private mDataSourceDataFieldName As Variant
Private mShowTitle As Boolean
Private mTitlePostition As chPosition
Private mTitleText As String
Private mShowLegend As Boolean
Private mLegendPostition As chPosition
Private mShowYAxis As Boolean
Private mShowXAxis As Boolean
Private mYAxisLabelText As String
Private mXAxisLabelText As String
Private mChartJsSource As chScriptSource
Private mChartJsCDN As String
Private mChartJsPathFilename As String
Private mStackedChart As Boolean
Private mColorSchemaSource As chScriptSource
Private mColorSchemaCDN As String
Private mColorSchemaPathFilename As String
Private mColorSchema As String
Private mAreaFilled As Boolean
Private mChartBackgroundColor As String
Private mDefaultFontColor As String
Private mDefaultFontFamily As String
Private mDefaultFontSize As Integer
Private mYAxisBeginAtZero As Boolean
Private mXAxisBeginAtZero As Boolean
Private mDataLabels As Variant
Private mDataSetLabel As Variant
Private mDataSetData As Variant
Private mIsSingleDataset As Boolean
Private mIsSingleBarChart As Boolean


'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------       METHODS    ------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------





'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Set control(ByVal pControl As WebbrowserControl)
    Set mControl = pControl
End Property






'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get control() As WebbrowserControl
    Set control = mControl
End Property






'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ChartType(ByVal pChartType As chChartType)
    mChartType = pChartType
    Select Case pChartType
        Case chChartType.Bar, chChartType.HorizontalBar, chChartType.Line, chChartType.Bubble, chChartType.Scatter
            ShowYAxis = True
            ShowXAxis = True
        Case chChartType.Doughnut, chChartType.Pie, chChartType.PolarArea, chChartType.Radar
            ShowYAxis = False
            ShowXAxis = False
        Case Else
            'Do Nothing
    End Select
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ChartType() As chChartType
    ChartType = mChartType
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let DataSourceObjectName(ByVal pDataSourceObjectName As String)
    mDataSourceObjectName = pDataSourceObjectName
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DataSourceObjectName() As String
    DataSourceObjectName = mDataSourceObjectName
End Property






'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let DataSourceSQLStatement(ByVal pDataSourceSQLStatement As String)
    mDataSourceSQLStatement = pDataSourceSQLStatement
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DataSourceSQLStatement() As String
    DataSourceSQLStatement = mDataSourceSQLStatement
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Set DataSourceRecordset(ByVal pDataSourceRecordset As DAO.Recordset)
    Set mDataSourceRecordset = pDataSourceRecordset
End Property






'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DataSourceRecordset() As DAO.Recordset
    Set DataSourceRecordset = mDataSourceRecordset
End Property






'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let DataSourceLabelFieldName(ByVal pDataSourceLabelFieldName As String)
    mDataSourceLabelFieldName = pDataSourceLabelFieldName
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DataSourceLabelFieldName() As String
    DataSourceLabelFieldName = mDataSourceLabelFieldName
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let DataSourceDataFieldName(ByRef pDataSourceDataFieldName As Variant)
    If IsArray(pDataSourceDataFieldName) = True Then
        mDataSourceDataFieldName = pDataSourceDataFieldName
    Else
        mDataSourceDataFieldName = Array(pDataSourceDataFieldName)
    End If
    If LBound(mDataSourceDataFieldName) = UBound(mDataSourceDataFieldName) Then
        mIsSingleDataset = True
    Else
        mIsSingleDataset = False
    End If
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Variant
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DataSourceDataFieldName() As Variant
    DataSourceDataFieldName = mDataSourceDataFieldName
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ShowTitle(ByVal pShowTitle As Boolean)
    mShowTitle = pShowTitle
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ShowTitle() As Boolean
    ShowTitle = mShowTitle
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let TitlePostition(ByVal pTitlePostition As chPosition)
    mTitlePostition = pTitlePostition
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get TitlePostition() As chPosition
    TitlePostition = mTitlePostition
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let TitleText(ByVal pTitleText As String)
    mTitleText = pTitleText
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get TitleText() As String
    TitleText = mTitleText
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ShowLegend(ByVal pShowLegend As Boolean)
    mShowLegend = pShowLegend
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ShowLegend() As Boolean
    ShowLegend = mShowLegend
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let LegendPostition(ByVal pLegendPostition As chPosition)
    mLegendPostition = pLegendPostition
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get LegendPostition() As chPosition
    LegendPostition = mLegendPostition
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ShowYAxis(ByVal pShowYAxis As Boolean)
    mShowYAxis = pShowYAxis
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ShowYAxis() As Boolean
    ShowYAxis = mShowYAxis
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ShowXAxis(ByVal pShowXAxis As Boolean)
    mShowXAxis = pShowXAxis
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ShowXAxis() As Boolean
    ShowXAxis = mShowXAxis
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let YAxisLabelText(ByVal pYAxisLabelText As String)
    mYAxisLabelText = pYAxisLabelText
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get YAxisLabelText() As String
    YAxisLabelText = mYAxisLabelText
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let XAxisLabelText(ByVal pXAxisLabelText As String)
    mXAxisLabelText = pXAxisLabelText
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get XAxisLabelText() As String
    XAxisLabelText = mXAxisLabelText
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ChartJsSource(ByVal pChartJsSource As chScriptSource)
    mChartJsSource = pChartJsSource
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ChartJsSource() As chScriptSource
    ChartJsSource = mChartJsSource
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ChartJsCDN(ByVal pChartJsCDN As String)
    mChartJsCDN = pChartJsCDN
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ChartJsCDN() As String
    ChartJsCDN = mChartJsCDN
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ChartJsPathFilename(ByVal pChartJsPathFilename As String)
    mChartJsPathFilename = pChartJsPathFilename
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ChartJsPathFilename() As String
    ChartJsPathFilename = mChartJsPathFilename
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let StackedChart(ByVal pStackedChart As Boolean)
    mStackedChart = pStackedChart
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get StackedChart() As Boolean
    StackedChart = mStackedChart
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ColorSchemaSource(ByVal pColorSchemaSource As chScriptSource)
    mColorSchemaSource = pColorSchemaSource
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ColorSchemaSource() As chScriptSource
    ColorSchemaSource = mColorSchemaSource
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ColorSchemaCDN(ByVal pColorSchemaCDN As String)
    mColorSchemaCDN = pColorSchemaCDN
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ColorSchemaCDN() As String
    ColorSchemaCDN = mColorSchemaCDN
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ColorSchemaPathFilename(ByVal pColorSchemaPathFilename As String)
    mColorSchemaPathFilename = pColorSchemaPathFilename
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ColorSchemaPathFilename() As String
    ColorSchemaPathFilename = mColorSchemaPathFilename
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ColorSchema(ByVal pColorSchema As String)
    mColorSchema = pColorSchema
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ColorSchema() As String
    ColorSchema = mColorSchema
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let AreaFilled(ByVal pAreaFilled As Boolean)
    mAreaFilled = pAreaFilled
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get AreaFilled() As Boolean
    AreaFilled = mAreaFilled
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ChartBackgroundColor(ByVal pChartBackgroundColor As String)
    mChartBackgroundColor = pChartBackgroundColor
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ChartBackgroundColor() As String
    ChartBackgroundColor = mChartBackgroundColor
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let DefaultFontColor(ByVal pDefaultFontColor As String)
    mDefaultFontColor = pDefaultFontColor
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DefaultFontColor() As String
    DefaultFontColor = mDefaultFontColor
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let DefaultFontFamily(ByVal pDefaultFontFamily As String)
    mDefaultFontFamily = pDefaultFontFamily
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DefaultFontFamily() As String
    DefaultFontFamily = mDefaultFontFamily
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let DefaultFontSize(ByVal pDefaultFontSize As Integer)
    mDefaultFontSize = pDefaultFontSize
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DefaultFontSize() As Integer
    DefaultFontSize = mDefaultFontSize
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let YAxisBeginAtZero(ByVal pYAxisBeginAtZero As Boolean)
    mYAxisBeginAtZero = pYAxisBeginAtZero
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get YAxisBeginAtZero() As Boolean
    YAxisBeginAtZero = mYAxisBeginAtZero
End Property




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let XAxisBeginAtZero(ByVal pXAxisBeginAtZero As Boolean)
    mXAxisBeginAtZero = pXAxisBeginAtZero
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get XAxisBeginAtZero() As Boolean
    XAxisBeginAtZero = mXAxisBeginAtZero
End Property





'---------------------------------------------------------------------------------
'   Type:            Constructor
'   Name:            Class_Initialize()
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
    ChartType = chChartType.Bar
    StackedChart = False
    DefaultFontColor = "Black"
    ChartBackgroundColor = "#EAEAEA"
    ShowTitle = False
    TitlePostition = posTop
    TitleText = "Budget Execution"
    ShowLegend = False
    LegendPostition = posRight
    ShowYAxis = True
    ShowXAxis = True
    ChartJsSource = chScriptSource.CDN
    ChartJsCDN = "https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.9.4/Chart.min.js"
    ChartJsPathFilename = CurrentProject.path & "\Chart.min.js"
    ColorSchemaSource = chScriptSource.CDN
    ColorSchemaCDN = "https://cdn.jsdelivr.net/npm/chartjs-plugin-colorschemes"
    ColorSchemaPathFilename = CurrentProject.path & "chartjs-plugin-colorschemes.min.js"
    ColorSchema = "tableau.Traffic9"
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Terminate()
    ' Do Nothing
End Sub





'---------------------------------------------------------------------------------
'   Type:            Sub Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub DrawChart()
    Dim filename As String
    CreateChartFile
    filename = CurrentProject.path & "\" & HtmlFileName
    control.Object.Navigate filename
End Sub






'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CreateChartFile()
    Dim f As Integer
    Dim strPath As String
    Dim strFilename As String
    Dim strContent As String
    strPath = CurrentProject.path
    strFilename = "\" & HtmlFileName
    strContent = strContent & getHeader
    strContent = strContent & getChartType
    strContent = strContent & getChartData
    strContent = strContent & getChartOptions
    strContent = strContent & getFooter
    f = FreeFile()
    Open strPath & strFilename For Output As f
    Print #f, strContent;
    Close f
End Sub




'---------------------------------------------------------------------------------
'   Type:            Property Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function getHeader() As String
    Dim strText As String
    strText = strText & "<!DOCTYPE html>" & vbCrLf
    strText = strText & "<!-- saved from url=(0014)about:internet -->" & vbCrLf
    strText = strText & "<!-- saved from url=(0016)http://localhost -->" & vbCrLf
    strText = strText & "<html>" & vbCrLf
    strText = strText & "<head>" & vbCrLf
    strText = strText & "    <meta charset=""utf-8"">" & vbCrLf
    strText = strText & "    <meta http-equiv=""X-UA-Compatible"" content=""IE=Edge""/>" & vbCrLf
    strText = strText & "    <title>Chart</title>" & vbCrLf
    strText = strText & "    <style>body{margin:0;padding:0;}</style>" & vbCrLf
    strText = strText & "    <script src=" & getChartJsLocation & "></script>" & vbCrLf
    strText = strText & "    <script src=" & GetColorSchemaLocation & "></script>" & vbCrLf
    strText = strText & "</head>" & vbCrLf
    strText = strText & "<body>" & vbCrLf
    strText = strText & "<div id=""container"" style=""width: 100%;"">" & vbCrLf
    If Len(ChartBackgroundColor) > 0 Then
        strText = strText & "    <canvas id=""myChart"" style=""background-color: " & ChartBackgroundColor & """ ></canvas>" & vbCrLf
    Else
        strText = strText & "    <canvas id=""myChart""></canvas>" & vbCrLf
    End If
    strText = strText & "</div>" & vbCrLf
    strText = strText & "<script>" & vbCrLf
    If Len(DefaultFontColor) > 0 Then
        strText = strText & "    Chart.defaults.global.defaultFontColor = '" & DefaultFontColor & "';" & vbCrLf
    End If
    If Len(DefaultFontFamily) > 0 Then
        strText = strText & "    Chart.defaults.global.defaultFontFamily = '" & DefaultFontFamily & "';" & vbCrLf
    End If
    If DefaultFontSize > 0 Then
        strText = strText & "    Chart.defaults.global.defaultFontSize = " & DefaultFontSize & ";" & vbCrLf
    End If
    strText = strText & "    var ctx = document.getElementById('myChart').getContext('2d');" & vbCrLf
    strText = strText & "    var myChart = new Chart(ctx, {" & vbCrLf
    getHeader = strText
End Function




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function getChartType() As String
    Dim strText As String
    Dim strType As String
    Select Case ChartType
        Case chChartType.Line
            strType = "line"
        Case chChartType.Bar
            strType = "bar"
        Case chChartType.HorizontalBar
            strType = "horizontalBar"
        Case chChartType.Radar
            strType = "radar"
        Case chChartType.Doughnut
            strType = "doughnut"
        Case chChartType.Pie
            strType = "pie"
        Case chChartType.PolarArea
            strType = "polarArea"
        Case chChartType.Bubble
            strType = "bubble"
        Case chChartType.Scatter
            strType = "scatter"
        Case Else
    End Select
    strText = strText & "        type: '" & strType & "'," & vbCrLf
    getChartType = strText
End Function


'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function getChartData() As String
    Dim strText As String
    Dim i As Integer
    Call GetData
    strText = strText & "        data: {" & vbCrLf
    If ChartType <> chChartType.Bubble Then
        If ChartType <> chChartType.Scatter Then
            strText = strText & "            labels: [" & mDataLabels & "]," & vbCrLf
        End If
    End If
    strText = strText & "            datasets: [" & vbCrLf
    For i = LBound(mDataSetData) To UBound(mDataSetData)
        Select Case ChartType
            Case chChartType.Bubble, chChartType.Scatter
                strText = strText & "               {label: '" & mDataLabels(i) & "' ," & vbCrLf
            Case Else
                strText = strText & "               {label: '" & mDataSetLabel(i) & "' ," & vbCrLf
        End Select
        Select Case ChartType
            Case chChartType.Line, chChartType.Radar
                If AreaFilled = False Then
                    strText = strText & "                fill: false ," & vbCrLf
                End If
            Case Else
        End Select
        strText = strText & "                data: [" & mDataSetData(i) & "]," & vbCrLf
        If mIsSingleBarChart = True Then
            strText = strText & "                backgroundColor: Chart['colorschemes']." & ColorSchema & "," & vbCrLf
        End If
        strText = strText & "                borderWidth: 1}," & vbCrLf
    Next i
    strText = CutTrailingComma(strText)
    strText = strText & "             ]" & vbCrLf
    strText = strText & "        }," & vbCrLf
    getChartData = strText

End Function


'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function getChartOptions() As String
    Dim strText As String
    strText = strText & "        options: {" & vbCrLf
    strText = strText & GetTitle
    strText = strText & GetLegend
    strText = strText & GetScales
    strText = strText & GetColorSchema
    strText = CutTrailingComma(strText)
    strText = strText & "        }" & vbCrLf
    
    getChartOptions = strText

End Function


'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function getFooter() As String
    Dim strText As String
    strText = strText & "    });" & vbCrLf
    strText = strText & "</script>" & vbCrLf
    strText = strText & "</body>" & vbCrLf
    strText = strText & "</html>"
    getFooter = strText
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function getChartJsLocation() As String
    Dim strText As String
    Select Case ChartJsSource
        Case chScriptSource.CDN
            strText = """" & ChartJsCDN & """"
        Case chScriptSource.LocalFile
            strText = """" & ChartJsPathFilename & """"
        Case Else
    End Select
    getChartJsLocation = strText
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetColorSchemaLocation() As String
    Dim strText As String
    Select Case ColorSchemaSource
        Case chScriptSource.CDN
            strText = """" & ColorSchemaCDN & """"
        Case chScriptSource.LocalFile
            strText = """" & ColorSchemaPathFilename & """"
        Case Else
    End Select
    GetColorSchemaLocation = strText
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub GetData()
On Error GoTo Handle_Error
    Dim DB As DAO.Database
    Dim rst As DAO.Recordset
    Dim strLabels() As String
    Dim strData() As String
    Dim field As Variant
    Dim i As Integer
    Set DB = CurrentDb()
    If Len(DataSourceObjectName) > 0 Then
        Set rst = DB.OpenRecordset(DataSourceObjectName)
    ElseIf Len(DataSourceSQLStatement) > 0 Then
        Set rst = DB.OpenRecordset(DataSourceSQLStatement)
    ElseIf Not DataSourceRecordset Is Nothing Then
        Set rst = DataSourceRecordset
    Else
        DataSourceDataFieldName = "Demo"
    End If
    Select Case ChartType
        Case chChartType.Bubble
            If Not rst Is Nothing Then
                ReDim strLabels(rst.RecordCount - 1)
                ReDim strData(rst.RecordCount - 1)
                Do While Not rst.EOF
                    strLabels(i) = rst(DataSourceLabelFieldName)
                    strData(i) = "{x:" & str$(rst(DataSourceDataFieldName(0))) & _
                                 ", y:" & str$(rst(DataSourceDataFieldName(1))) & _
                                 ", r:" & str$(rst(DataSourceDataFieldName(2))) & "}"
                    i = i + 1
                    rst.MoveNext
                Loop
            Else
                ReDim strLabels(0)
                ReDim strData(0)
                strLabels(0) = "Demo"
                strData(0) = "{x:2,y:3,r:17}"
            End If
                mDataLabels = strLabels()
                mDataSetData = strData()
        Case chChartType.Scatter
            If Not rst Is Nothing Then
                ReDim strLabels(rst.RecordCount - 1)
                ReDim strData(rst.RecordCount - 1)
                Do While Not rst.EOF
                    strLabels(i) = rst(DataSourceLabelFieldName)
                    strData(i) = "{x:" & str$(rst(DataSourceDataFieldName(0))) & _
                                 ", y:" & str$(rst(DataSourceDataFieldName(1))) & "}"
                    i = i + 1
                    rst.MoveNext
                Loop
            Else
                ReDim strLabels(0)
                ReDim strData(0)
                strLabels(0) = "Demo"
                strData(0) = "{x:2,y:3,r:17}"
            End If
                mDataLabels = strLabels()
                mDataSetData = strData()
        Case Else
            ReDim strLabels(1)
            ReDim strData(UBound(DataSourceDataFieldName))
            If Not rst Is Nothing Then
                Do While Not rst.EOF
                    strLabels(0) = strLabels(0) & "'" & rst(DataSourceLabelFieldName) & "',"
                    i = 0
                    For Each field In DataSourceDataFieldName
                        strData(i) = strData(i) & str$(rst(DataSourceDataFieldName(i))) & ","
                        i = i + 1
                    Next field
                    rst.MoveNext
                Loop
            Else
                strLabels(0) = "'Berlin', 'London', 'Paris', 'Rome', 'Lisbon', 'Madrid',"
                strData(0) = "2, 13, 7, 9, 12, 5,"
            End If
            DetermineSingleBarChart
            ReDim mDataSet1Colors(UBound(DataSourceDataFieldName))
            mDataLabels = CutTrailingComma(strLabels(0))
            mDataSetLabel = DataSourceDataFieldName()
            i = 0
            For Each field In DataSourceDataFieldName()
                strData(i) = CutTrailingComma(strData(i))
                i = i + 1
            Next field
            mDataSetData = strData()
    End Select
Exit_Here:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set DB = Nothing
    Exit Sub
Handle_Error:
    Select Case Err.Number
        Case Else
            MsgBox Err.Description, vbExclamation, Err.Number
    End Select
    Resume Exit_Here
    Resume Next
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub Procedure
'   Name:            DeterminSingleBarChart()
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DetermineSingleBarChart()
    Select Case mChartType
        Case chChartType.Bar, chChartType.HorizontalBar
            If mIsSingleDataset = True Then
                mIsSingleBarChart = True
            Else
                mIsSingleBarChart = False
            End If
        Case Else
            mIsSingleBarChart = False
    End Select
End Sub




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetTitle
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetTitle() As String
    Dim strText As String
    If ShowTitle = False Then
        strText = strText & "            title: {" & vbCrLf
        strText = strText & "                display: false," & vbCrLf
        strText = strText & "            }," & vbCrLf
        GetTitle = strText
        Exit Function
    End If
   
    strText = strText & "            title: {" & vbCrLf
    strText = strText & "                display: true," & vbCrLf
    strText = strText & "                position: " & getPositionString(TitlePostition) & "," & vbCrLf
    strText = strText & "                text: '" & TitleText & "'" & vbCrLf
    strText = strText & "            }," & vbCrLf
    GetTitle = strText
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetLegend() As String
    Dim strText As String
    If ShowLegend = False Then
        strText = strText & "            legend: {" & vbCrLf
        strText = strText & "                display: false" & vbCrLf
        strText = strText & "            }," & vbCrLf
        GetLegend = strText
        Exit Function
    End If
    strText = strText & "            legend: {" & vbCrLf
    strText = strText & "                display: true," & vbCrLf
    strText = strText & "                position: " & getPositionString(LegendPostition) & vbCrLf
    strText = strText & "            }," & vbCrLf
    GetLegend = strText
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetScales() As String
    Dim strText As String
    strText = strText & "            scales: {" & vbCrLf
    strText = strText & "                yAxes: [{" & vbCrLf
    strText = strText & "                    id: 'first-y-Axis'," & vbCrLf
    If StackedChart = True Then
        Select Case ChartType
            Case chChartType.Line, chChartType.Bar, chChartType.HorizontalBar
                strText = strText & "                    stacked: true," & vbCrLf
            Case Else
        End Select
    End If
    If ShowYAxis = False Then
        strText = strText & "                    display: false" & vbCrLf
    Else
        strText = strText & "                    display: true," & vbCrLf
        If Len(YAxisLabelText) > 0 Then
            strText = strText & "                    scaleLabel: {" & vbCrLf
            strText = strText & "                       display: true," & vbCrLf
            strText = strText & "                       labelString: '" & YAxisLabelText & "'" & vbCrLf
            strText = strText & "                    }," & vbCrLf
        End If
        If YAxisBeginAtZero = True Then
            strText = strText & "                    ticks: {" & vbCrLf
            strText = strText & "                       beginAtZero: true" & vbCrLf
            strText = strText & "                    }," & vbCrLf
        End If
    End If
    strText = CutTrailingComma(strText)
    strText = strText & "                    }]," & vbCrLf
    strText = strText & "                xAxes: [{" & vbCrLf
    strText = strText & "                    id: 'first-x-Axis'," & vbCrLf
    If StackedChart = True Then
        Select Case ChartType
            Case chChartType.Line, chChartType.Bar, chChartType.HorizontalBar
                strText = strText & "                    stacked: true," & vbCrLf
            Case Else
        End Select
    End If
    If ShowXAxis = False Then
        strText = strText & "                    display: false" & vbCrLf
    Else
        strText = strText & "                    display: true," & vbCrLf
        If Len(XAxisLabelText) > 0 Then
            strText = strText & "                    scaleLabel: {" & vbCrLf
            strText = strText & "                    display: true," & vbCrLf
            strText = strText & "                    labelString: '" & XAxisLabelText & "'" & vbCrLf
            strText = strText & "                    }," & vbCrLf
        End If
        If XAxisBeginAtZero = True Then
            strText = strText & "                    ticks: {" & vbCrLf
            strText = strText & "                       beginAtZero: true" & vbCrLf
            strText = strText & "                    }," & vbCrLf
        End If
    End If
    strText = CutTrailingComma(strText)
    strText = strText & "                    }]," & vbCrLf
    strText = strText & "            }," & vbCrLf
    GetScales = strText
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetColorSchema() As String
    Dim strText As String
    strText = strText & "            plugins: {" & vbCrLf
    strText = strText & "                colorschemes: {" & vbCrLf
    strText = strText & "                    scheme: '" & ColorSchema & "'" & vbCrLf
    strText = strText & "                    }" & vbCrLf
    strText = strText & "                }," & vbCrLf
    GetColorSchema = strText
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function getPositionString(ByVal ThisPosition As chPosition) As String
    Select Case ThisPosition
        Case chPosition.posTop
            getPositionString = "'top'"
        Case chPosition.posLeft
            getPositionString = "'left'"
        Case chPosition.posBottom
            getPositionString = "'bottom'"
        Case chPosition.posRight
            getPositionString = "'right'"
        Case Else
            getPositionString = "'top'"
    End Select
End Function


'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetChartType
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function CutTrailingComma(ByVal strText As String) As String
    If Right$(strText, 1) = "," Then
        CutTrailingComma = Left$(strText, Len(strText) - 1)
    ElseIf Right$(strText, 3) = "," & vbCrLf Then
        CutTrailingComma = Left$(strText, Len(strText) - 3) & vbCrLf
    Else
        CutTrailingComma = strText
    End If
End Function







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



