VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_BudgetPath As BudgetPath
Private m_Name As String
Private m_Extension As String
Private m_FileSys As Scripting.FileSystemObject
Private m_Excel As Excel.Application
Private m_Workbook As Excel.Workbook
Private m_Worksheet As Excel.Worksheet
Private m_Chart As Excel.Chart
Private m_PivotTable As Excel.PivotTable
Private m_PivotFields As Excel.PivotFields
Private m_PivotItems As Excel.PivotItems
Private m_Error As String
Private m_Notification As String
Private m_Data As DAO.Database
Private m_DataSet As DAO.Recordset
Private m_Query As DAO.QueryDef
Private m_Connection As ADODB.Connection





'---------------------------------------------------------------------------------
'   Type:            Constructor
'   Name:
'   Parameters:      Void
'   Retval:          BudgetFile
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_Excel = New Excel.Application
    Set m_BudgetPath = New BudgetPath
End Sub

Private Sub Class_Terminate()
    If Not m_Excel Is Nothing Then
        Set m_Excel = Nothing
    End If
    If Not m_BudgetPath Is Nothing Then
        Set m_BudgetPath = Nothing
    End If
End Sub




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            WorkSheetExists
'   Parameters:      Void
'   Retval:          BudgetFile
'   Purpose:
'---------------------------------------------------------------------------------
Public Function WorksheetExists(pSheetName As String, pWorkbook As Workbook) As Boolean
    On Error GoTo ErrorHandler:
    Dim p_CheckSheet As Worksheet
    Set p_CheckSheet = pWorkbook.Worksheets(pSheetName)
    WorksheetExists = True
    Exit Function
ErrorHandler:
    ProcessError
    WorksheetExists = False
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Destructor
'   Name:
'   Parameters:      Void
'   Retval:          BudgetFile
'   Purpose:
'---------------------------------------------------------------------------------
Public Function WorkbookExists(pWorkbookName As String) As Boolean
    On Error GoTo ErrorHandler:
    Dim m_Test As Excel.Workbook
    Set m_Test = New Excel.Workbook
    Set m_Test = Workbooks(pWorkbookName)
    WorkbookExists = True
    Exit Function
ErrorHandler:
    ProcessError
    WorkbookExists = False
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            CreateWorkbook
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub CreateWorkbook(pFilePath As String)
    On Error GoTo ErrorHandler:
    Set m_Excel = CreateObject("Excel.Application")
    Set m_Workbook = m_Excel.Workbooks.add
    Set m_Worksheet = m_Workbook.Worksheets(1)
    m_Workbook.SaveAs filename:=pFilePath, _
        FileFormat:=xlOpenXMLWorkbook
    m_Workbook.Close SaveChanges:=False
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RecursiveDelete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetReportTemplate() As Excel.Workbook
    On Error GoTo ErrorHandler:
    Dim m_BudgetPath As BudgetPath
    Dim m_Template As String
    Dim m_Report As String
    Set m_BudgetPath = New BudgetPath
    m_Template = m_BudgetPath.ReportTemplate
    Set m_Workbook = m_Excel.Workbooks.Open(m_Template)
    Set GetReportTemplate = m_Workbook
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RecursiveDelete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBudgetTemplate() As Excel.Workbook
    On Error GoTo ErrorHandler:
    Dim m_BudgetPath As BudgetPath
    Dim m_Template As String
    Dim m_Report As String
    Set m_BudgetPath = New BudgetPath
    m_Template = m_BudgetPath.BudgetTemplate
    Set m_Workbook = m_Excel.Workbooks.Open(m_Template)
    m_Excel.Workbooks.add (m_Workbook)
    Set GetReportTemplate = m_Workbook
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub OpenFileDialog()
    On Error GoTo ErrorHandler:
    Dim m_Dialog As FileDialog
    Dim result As Integer
    Set m_Dialog = Application.FileDialog(msoFileDialogFilePicker)
    m_Dialog.AllowMultiSelect = False
    m_Dialog.Title = "Select an Excel file"
    m_Dialog.InitialFileName = "C:\"
    m_Dialog.Filters.Clear
    m_Dialog.Filters.add "Excel files", "*.xlsx"
    m_Dialog.Filters.add "All files", "*.*"
    If m_Dialog.Show = -1 Then
       Debug.Print m_Dialog.SelectedItems(1)
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RecursiveDelete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub OpenExcelFile(filename As String, FilePath As String)
    On Error GoTo ErrorHandler:
    Dim view As ADODB.Connection
    Dim m_DataSet As ADODB.Recordset
    Set view = New ADODB.Connection
    view.Open "Provider=Microsoft.ACE.OLEDB.12.0;" _
        & "Data Source=" & filename & _
        "Extended Properties=""Excel 12.0; HDR=YES"";"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            CreateCsvFile
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub CreateCsvFile(pFileName As String, pFilePath As String)
    On Error GoTo ErrorHandler:
    If Not pFileName & "" = "" And _
        Not pFilePath & "" = "" Then
            Dim p_FilePath As String
            Dim p_CsvName As String
            Dim p_Sheet As Worksheet
            Set p_Sheet = ActiveWorkbook.Sheets("Sheet1")
            p_CsvName = pFileName & ".csv"
            p_FilePath = Application.CurrentProject
            p_Sheet.Copy
            ActiveWorkbook.SaveAs filename:=p_FilePath & "\" & p_CsvName, _
                FileFormat:=xlCSV, CreateBackup:=False
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------------
' Type:         Function
' Name:         ExportQuery(String)
' Parameters:   Sting
' Retval:       Excel.Application
' Purpose:      Export recordset to Excel
'---------------------------------------------------------------------------------------
Function ExportQuery(ByVal pQuery As String)
    On Error GoTo ErrorHandler:
    Set m_Data = Application.CurrentDb
    Set m_DataSet = m_Data.OpenRecordset(pQuery)
    Dim m_IsOpen As Boolean
    Dim iCols As Integer
    Const xlCenter = -4108
    Set m_Excel = CreateObject("Excel.Application")
    m_IsOpen = True
    m_Excel.ScreenUpdating = False
    m_Excel.Visible = False   '
    Set m_Workbook = m_Excel.Workbooks.add()
    Set m_Worksheet = m_Workbook.Sheets(1)
    With m_DataSet
        If .RecordCount <> 0 Then
            For iCols = 0 To m_DataSet.Fields.count - 1
                m_Worksheet.Cells(1, iCols + 1).Value = m_DataSet.Fields(iCols).Name
            Next
            With m_Worksheet.Range(m_Worksheet.Cells(1, 1), _
                m_Worksheet.Cells(1, m_DataSet.Fields.count))
                .Font.Bold = True
                .Font.ColorIndex = 2
                .Interior.ColorIndex = 1
                .HorizontalAlignment = xlCenter
            End With
            m_Worksheet.Range("A2").CopyFromRecordset m_DataSet
            m_Worksheet.Range(m_Worksheet.Cells(1, 1), _
                m_Worksheet.Cells(1, m_DataSet.Fields.count)).Columns.AutoFit
            m_Worksheet.Range("A1").Select
        End If
    End With
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------------
' Type:         Function
' Name:         ExportData(DAO.Recordset)
' Parameters:   Sting
' Retval:       Excel.Application
' Purpose:      Export recordset to Excel
'---------------------------------------------------------------------------------------
Public Function ExportRecordset(pData As DAO.Recordset)
    On Error GoTo ErrorHandler:
    Set m_DataSet = pData
    Dim m_IsOpen As Boolean
    Dim iCols As Integer
    Const xlCenter = -4108
    Set m_Excel = CreateObject("Excel.Application")
    m_IsOpen = True
    m_Excel.ScreenUpdating = False
    m_Excel.Visible = False   '
    Set m_Workbook = m_Excel.Workbooks.add()
    Set m_Worksheet = m_Workbook.Sheets(1)
    With pData
        If .RecordCount <> 0 Then
            For iCols = 0 To pData.Fields.count - 1
                m_Worksheet.Cells(1, iCols + 1).Value = pData.Fields(iCols).Name
            Next
            With m_Worksheet.Range(m_Worksheet.Cells(1, 1), _
                m_Worksheet.Cells(1, pData.Fields.count))
                .Font.Bold = True
                .Font.ColorIndex = 2
                .Interior.ColorIndex = 1
                .HorizontalAlignment = xlCenter
            End With
            m_Worksheet.Range("A2").CopyFromRecordset pData
            m_Worksheet.Range(m_Worksheet.Cells(1, 1), _
                m_Worksheet.Cells(1, pData.Fields.count)).Columns.AutoFit
            m_Worksheet.Range("A1").Select
        End If
    End With
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub OpenSaveDialog()
    On Error GoTo ErrorHandler:
    Dim m_Dialog As FileDialog
    Dim result As Integer
    Set m_Dialog = Application.FileDialog(msoFileDialogSaveAs)
    m_Dialog.AllowMultiSelect = False
    m_Dialog.Title = "Select location to save file."
    m_Dialog.InitialFileName = "C:\"
    m_Dialog.Filters.Clear
    If m_Dialog.Show = -1 Then
       Debug.Print m_Dialog.SelectedItems(1)
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:         KillExcel()
' Parameters:   Void
' Retval:       Void
' Purpose:      Releases resources used by the Excel Application
'---------------------------------------------------------------------------------------
Private Sub KillExcel()
    On Error GoTo ErrorHandler:
    If Not m_Excel Is Nothing Then
        Set m_Excel = Nothing
    End If
    If Not m_Workbook Is Nothing Then
        Set m_Workbook = Nothing
    End If
    If Not m_Workbook Is Nothing Then
        Set m_Worksheet = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:         KillExcel()
' Parameters:   Void
' Retval:       Void
' Purpose:      Releases resources used by the Excel Application
'---------------------------------------------------------------------------------------
Private Sub KillChart()
    On Error GoTo ErrorHandler:
    If Not m_Chart Is Nothing Then
        Set m_Chart = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:         KillExcel()
' Parameters:   Void
' Retval:       Void
' Purpose:      Releases resources used by the Pivot Table
'---------------------------------------------------------------------------------------
Private Sub KillPivot()
    On Error GoTo ErrorHandler:
    If Not m_PivotTable Is Nothing Then
        Set m_PivotTable = Nothing
    End If
    If Not m_PivotFields Is Nothing Then
        Set m_PivotFields = Nothing
    End If
    If Not m_PivotItems Is Nothing Then
        Set m_PivotItems = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub






'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:         KillData()
' Parameters:   Void
' Retval:       Void
' Purpose:      Releases resources used by the Database
'---------------------------------------------------------------------------------------
Private Sub KillData()
    On Error GoTo ErrorHandler:
    If Not m_Data Is Nothing Then
        Set m_Data = Nothing
    End If
    If Not m_DataSet Is Nothing Then
        Set m_DataSet = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub







'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:         KillFiles()
' Parameters:   Void
' Retval:       Void
' Purpose:      Releases resources used by the BudgetPath class
'---------------------------------------------------------------------------------------
Private Sub KillFiles()
    On Error GoTo ErrorHandler:
    If Not m_BudgetPath Is Nothing Then
        Set m_BudgetPath = Nothing
    End If
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



