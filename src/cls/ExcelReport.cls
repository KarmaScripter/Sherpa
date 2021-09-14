VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_BudgetPath As BudgetPath
Private m_SavePath As String
Private m_Name As String
Private m_Extension As String
Private m_Location As String
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






'---------------------------------------------------------------------------------
'   Type:            Constructor/Finalizer
'   Name:            ExcelReport
'   Parameters:      Void
'   Retval:          BudgetFile
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_Data = CurrentDb
    Set m_Excel = New Excel.Application
    Set m_Workbook = New Excel.Workbook
    Set m_Worksheet = New Excel.Worksheet
    Set m_BudgetPath = New BudgetPath
End Sub

Private Sub Class_Terminate()
    KillExcel
    KillData
    KillFiles
End Sub





'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            CreateWorkbook
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub CreateWorkbook(pFilePath As String)
    On Error GoTo ErrorHandler:
    If Not pFilePath & "" = "" And _
        Not Dir(pFilePath) & "" = "" Then
            Set m_Excel = CreateObject("Excel.Application")
            Set m_Workbook = m_Excel.Workbooks.Open(pFilePath)
            Set m_Worksheet = m_Workbook.Worksheets(1)
            m_Workbook.SaveAs filename:=pFilePath, _
                FileFormat:=xlOpenXMLWorkbook
            m_Workbook.Close SaveChanges:=False
    End If
ErrorHandler:
    If Err.Number > 0 Then
        KillExcel
        m_Error = "Source:       ExceReport" _
            & vbCrLf & "Member:     CreateWorkbook(String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RecursiveDelete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetTemplate() As Excel.Workbook
    On Error GoTo ErrorHandler:
    Dim m_Template As String
    m_Template = m_BudgetPath.ReportTemplate
    Set m_Excel = CreateObject("Excel.Application")
    Set m_Workbook = m_Excel.Workbooks.Open(m_Template)
    Set GetTemplate = m_Workbook
ErrorHandler:
    If Err.Number > 0 Then
        KillExcel
        m_Error = "Source:       ExceReport" _
            & vbCrLf & "Member:     GetTemplate()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RecursiveDelete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetRange(pColumn As Integer, pRow As Integer) As Excel.Range
    On Error GoTo ErrorHandler:
    If pColumn > 0 And pRow > 0 Then
        Set m_Excel = CreateObject("Excel.Application")
        Set m_Workbook = m_Excel.Workbooks.add
        Set m_Worksheet = m_Workbook.Worksheets.add
        Dim m_Range As Excel.Range
        Dim m_Template As String
        Set m_BudgetPath = New BudgetPath
        m_Template = m_BudgetPath.ReportTemplate
        Set m_Workbook = m_Excel.Workbooks.Open(m_Template)
        Set m_Range = m_Workbook.Worksheets(1).Cells(pColumn, pRow)
        Set GetRange = m_Range
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       ExceReport" _
            & vbCrLf & "Member:     GetTemplate()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RecursiveDelete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub PopulateData(pData As DAO.Recordset)
    On Error GoTo ErrorHandler:
    Dim m_Template As String
    Dim m_Range As Excel.Range
    Set m_Excel = CreateObject("Excel.Application")
    Set m_Workbook = m_Excel.Workbooks.add
    Set m_Workbook = m_Excel.Workbooks.Open(m_Template)
    Set m_Range = m_Workbook.Worksheets(1).Cells(1, 2)
    m_Template = m_BudgetPath.ReportTemplate
    m_Workbook.Worksheets(2).Visible = False
    m_Range.CopyFromRecordset Data:=pData, MaxRows:=10, MaxColumns:=8
    m_Excel.Visible = True
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       ExceReport" _
            & vbCrLf & "Member:     PopulateData(DAO.Recordset)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub OpenFileDialog()
    On Error GoTo ErrorHandler:
    Dim p_Dialog As FileDialog
    Dim p_Result As Integer
    Set p_Dialog = Application.FileDialog(msoFileDialogFilePicker)
    p_Dialog.AllowMultiSelect = False
    p_Dialog.Title = "Select an Excel file"
    p_Dialog.InitialFileName = "C:\"
    p_Dialog.Filters.Clear
    p_Dialog.Filters.add "Excel files", "*.xlsx"
    p_Dialog.Filters.add "All files", "*.*"
    If p_Dialog.Show = -1 Then
       Debug.Print p_Dialog.SelectedItems(1)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       ExceReport" _
            & vbCrLf & "Member:     OpenFileDialog()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub OpenFolderDialog()
    On Error GoTo ErrorHandler:
    Dim p_Dialog As FileDialog
    Dim p_Result As Integer
    Set p_Dialog = Application.FileDialog(msoFileDialogFolderPicker)
    p_Dialog.AllowMultiSelect = False
    p_Dialog.Title = "Select an Folder"
    p_Dialog.InitialFileName = "C:\"
    p_Dialog.Filters.Clear
    If p_Dialog.Show = -1 Then
       Debug.Print p_Dialog.SelectedItems(1)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       ExceReport" _
            & vbCrLf & "Member:     OpenFolderDialog()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub OpenSaveDialog()
    On Error GoTo ErrorHandler:
    Dim p_Dialog As FileDialog
    Dim p_Result As Integer
    Set p_Dialog = Application.FileDialog(msoFileDialogSaveAs)
    p_Dialog.AllowMultiSelect = False
    p_Dialog.Title = "Select location to save file."
    p_Dialog.InitialFileName = "C:\"
    p_Dialog.Filters.Clear
    If p_Dialog.Show = -1 Then
       Debug.Print p_Dialog.SelectedItems(1)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       ExceReport" _
            & vbCrLf & "Member:     OpenFileDialog()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RecursiveDelete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub OpenExcelFile(pFileName As String, pFilePath As String)
    On Error GoTo ErrorHandler:
    Set m_Excel = CreateObject("Excel.Application")
    Set m_Workbook = m_Excel.Workbooks.Open(pFilePath)
    m_Excel.Visible = True
ErrorHandler:
    If Err.Number > 0 Then
        KillExcel
        m_Error = "Source:       ExceReport" _
            & vbCrLf & "Member:     OpenExcelFile(String, String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
        Not Dir(pFilePath) & "" = "" Then
            Set m_Excel = CreateObject("Excel.Application")
            Set m_Workbook = m_Excel.Workbooks.Open(pFilePath)
            Set m_Worksheet = m_Workbook.Worksheets.add
            m_Worksheet.Copy
            m_Workbook.SaveAs filename:=pFilePath & "\" & pFileName, _
                FileFormat:=xlCSV, CreateBackup:=False
    End If
ErrorHandler:
    If Err.Number > 0 Then
        KillExcel
        m_Error = "Source:       ExceReport" _
            & vbCrLf & "Member:     CreateCsvFile(String, String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------------
' Type:         Function
' Name:         ExcelExport(String)
' Parameters:   Sting
' Retval:       Excel.Application
' Purpose:      Export recordset to Excel
'---------------------------------------------------------------------------------------
Function ExportQuery(ByVal pQuery As String)
    On Error GoTo ErrorHandler:
    Set m_DataSet = m_Data.OpenRecordset(pQuery)
    Dim p_IsOpen As Boolean
    Dim p_Columns As Integer
    Const xlCenter = -4108
    Set m_Excel = CreateObject("Excel.Application")
    p_IsOpen = True
    m_Excel.ScreenUpdating = False
    m_Excel.Visible = False   '
    Set m_Workbook = m_Excel.Workbooks.add()
    Set m_Worksheet = m_Workbook.Sheets(1)
    With m_DataSet
        If .RecordCount <> 0 Then
            For p_Columns = 0 To m_DataSet.Fields.count - 1
                m_Worksheet.Cells(1, p_Columns + 1).Value = m_DataSet.Fields(p_Columns).Name
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
    If Err.Number > 0 Then
        KillExcel
        KillData
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     ExcelExport(String)" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------------
' Type:         Function
' Name:         ExcelExport(String)
' Parameters:   Sting
' Retval:       Excel.Application
' Purpose:      Export recordset to Excel
'---------------------------------------------------------------------------------------
Public Function ExportData(pData As DAO.Recordset)
    On Error GoTo ErrorHandler:
    If Not pData Is Nothing Then
        Set m_DataSet = pData
        Dim p_IsOpen As Boolean
        Dim p_Columns As Integer
        Const xlCenter = -4108
        Set m_Excel = CreateObject("Excel.Application")
        Set m_Workbook = m_Excel.Workbooks.add()
        Set m_Worksheet = m_Workbook.Sheets(1)
        p_IsOpen = True
        m_Excel.ScreenUpdating = False
        m_Excel.Visible = False   '
        With m_DataSet
            If .RecordCount <> 0 Then
                For p_Columns = 0 To m_DataSet.Fields.count - 1
                    m_Worksheet.Cells(1, p_Columns + 1).Value = m_DataSet.Fields(p_Columns).Name
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
    End If
ErrorHandler:
    If Err.Number > 0 Then
        KillExcel
        KillData
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     ExcelExport(String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function






'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:
' Parameters:   Sting
' Retval:       Void
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub SetLeftHeader(pText As String)
    On Error GoTo ErrorHandler:
    Dim p_LeftHeader As String
    p_LeftHeader = pText
    ActiveSheet.PageSetup.LeftHeader = p_LeftHeader
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     SetLeftHeader(String)" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub




'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:
' Parameters:   Sting
' Retval:       Void
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub SetCenterHeader(pText As String)
    On Error GoTo ErrorHandler:
    Dim p_CenterHeader As String
    p_CenterHeader = pText
    ActiveSheet.PageSetup.CenterHeader = p_CenterHeader
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     SetCenterHeader(String)" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub




'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:
' Parameters:   Sting
' Retval:       Void
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub SetRightHeader(pText As String)
    On Error GoTo ErrorHandler:
    Dim p_RightHeader As String
    p_RightHeader = pText
    ActiveSheet.PageSetup.RightFooter = p_RightHeader
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     SetRightHeader(String)" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub




'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:
' Parameters:   Sting
' Retval:       Void
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub SetLeftFooter(pText As String)
    On Error GoTo ErrorHandler:
    If Not pText & "" = "" Then
        Dim p_LeftFooter As String
        p_LeftFooter = pText
        ActiveSheet.PageSetup.LeftFooter = p_LeftFooter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     SetLeftFooter(String)" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub




'---------------------------------------------------------------------------------------
' Type:         Sub Procedure
' Name:
' Parameters:   Sting
' Retval:       Void
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub SetCenterFooter(pText As String)
    On Error GoTo ErrorHandler:
    Dim p_CenterFooter As String
    p_CenterFooter = pText
    ActiveSheet.PageSetup.CenterFooter = p_CenterFooter
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     SetCenterFooterString)" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Err.Number > 0 Then
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     KillExcel()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Err.Number > 0 Then
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     KillData()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    If Err.Number > 0 Then
        m_Error = "Source:   ExceReport" _
            & vbCrLf & "Member:     KillFiles()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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



