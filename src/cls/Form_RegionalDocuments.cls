VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RegionalDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private pRC As AccountArgs
Private DivisionNameFilter As String
Private DocumentNumberFilter As String
Private FiscalYearFilter As String
Private BocNameFilter As String
Private FundNameFilter As String
Private pFormFilter As String
Private pAnd As String
Private m_Error As String



'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    pAnd = " AND "
    FiscalYearFilter = "[BFY] = '2021'"
    Me.Filter = FiscalYearFilter
    Me.RecordSource = "SELECT * FROM Purchases WHERE NOT" _
        & " IsNull(Purchases.DocumentControlNumbers)" _
        & " AND Purchases.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:            Event / Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Section(acHeader).AutoHeight = False
    Me.Section(acHeader).Height = 2
    Me.Section(acDetail).AutoHeight = False
    Me.Section(acDetail).Height = 4
    Me.Section(acFooter).AutoHeight = False
    Me.Section(acFooter).Height = 1
    Set pRC = New AccountArgs
    Me.FiscalYearOptionGroup.Value = 2
    ClearComboBoxValues
    SetComboBoxColors
    SetButtonVisibility
    Me.DivisionNameComboBox.RowSource = "SELECT DISTINCT Purchases.DivisionName FROM Purchases" _
        & " WHERE NOT IsNull(Purchases.DocumentControlNumbers)" _
        & pAnd & "Purchases.BocCode NOT IN ('10', '17');"
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Purchases.BocName FROM Purchases" _
        & " WHERE NOT IsNull(Purchases.DocumentControlNumbers)" _
        & pAnd & "Purchases.BocCode NOT IN ('10', '17');"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            SetButtonVisibility
'   Parameters:      Void
'   Purpose:         Toggles button visibility on and off
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.AddButton.Visible = Not Me.AddButton.Visible
    Me.FirstButton.Visible = Not Me.FirstButton.Visible
    Me.PreviousButton.Visible = Not Me.PreviousButton.Visible
    Me.NextButton.Visible = Not Me.NextButton.Visible
    Me.LastButton.Visible = Not Me.LastButton.Visible
    Me.EditButton.Visible = Not Me.EditButton.Visible
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.DeleteButton.Visible = Not Me.DeleteButton.Visible
    Me.CalculatorButton.Visible = Not Me.CalculatorButton.Visible
    Me.ExcelButton.Visible = Not Me.ExcelButton.Visible
    Me.UndoButton.Visible = Not Me.UndoButton.Visible
    Me.SaveButton.Visible = Not Me.SaveButton.Visible
    Me.DataButton.Visible = Not Me.DataButton.Visible
    Me.CompassDocumentQueryButton.Visible = Not Me.CompassDocumentQueryButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            SetComboBoxColors
'   Parameters:      Void
'   Purpose:         Toggles button visibility ON and OFF
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.DocumentNumberComboBox.ForeColor = RGB(255, 255, 255)
    Me.DocumentNumberComboBox.BackColor = RGB(33, 33, 33)
    Me.DocumentNumberComboBox.BorderColor = RGB(68, 114, 196)
    Me.DivisionNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.DivisionNameComboBox.BackColor = RGB(33, 33, 33)
    Me.DivisionNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.DivisionNameComboBox.Value = vbNullString
    Me.DocumentNumberComboBox.Value = vbNullString
    Me.BocNameComboBox.Value = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearOptionGroupValues()
    On Error GoTo ErrorHandler:
    FiscalYearOptionGroup.Value = 2
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    DocumentNumberFilter = vbNullString
    FiscalYearFilter = vbNullString
    DivisionNameFilter = vbNullString
    BocNameFilter = vbNullString
    Me.RecordSource = "SELECT * FROM Purchases WHERE NOT" _
        & " IsNull(Purchases.DocumentControlNumbers)" _
        & pAnd & "Purchases.BocCode NOT IN ('10', '17');"
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    pRC.BFY = vbNullString
    FiscalYearFilter = vbNullString
    Me.Filter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            pRC.BFY = "2020"
            FiscalYearFilter = "[BFY] IN ('2019', '2020')"
        Case 2
            pRC.BFY = "2021"
            FiscalYearFilter = "[BFY] IN ('2021', '2022')"
    End Select
    Me.Filter = GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Purchases WHERE" _
        & " NOT IsNull(DocumentControlNumbers)" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFiscalYearFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetFiscalYearFilter()
    On Error GoTo ErrorHandler:
    If Not FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & DocumentNumberFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        DocumentNumberFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & DivisionNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        DivisionNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & DocumentNumberFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        DivisionNameFilter = vbNullString And _
        DocumentNumberFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DivisionNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    pRC.FundCode = Me.DivisionNameComboBox.Value
    DivisionNameFilter = "[DivisionName] = '" & Me.DivisionNameComboBox.Value & "'"
    Me.Filter = GetDivisionNameFilter
    Me.DocumentNumberComboBox.RowSource = "SELECT DISTINCT Purchases.DocumentControlNumbers" _
        & " FROM Purchases WHERE" _
        & " NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Purchases.BocName" _
        & " FROM Purchases WHERE" _
        & " NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
    Me.RecordSource = "SELECT * FROM Purchases" _
        & " WHERE NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetDivisionNameFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetDivisionNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not DivisionNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetDivisionNameFilter = DocumentNumberFilter & pAnd _
                & DivisionNameFilter & pAnd _
                & FiscalYearFilter & pAnd & BocNameFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetDivisionNameFilter = DocumentNumberFilter & pAnd _
                & DivisionNameFilter & pAnd _
                & FiscalYearFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetDivisionNameFilter = DocumentNumberFilter & pAnd _
                & DivisionNameFilter & pAnd & BocNameFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetDivisionNameFilter = DocumentNumberFilter & pAnd _
                & DivisionNameFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        DocumentNumberFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetDivisionNameFilter = DivisionNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        DocumentNumberFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetDivisionNameFilter = FiscalYearFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        DocumentNumberFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetDivisionNameFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function

'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DocumentNumberComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    DocumentNumberFilter = vbNullString
    BocNameFilter = vbNullString
    DocumentNumberFilter = "[DocumentControlNumbers] = '" & Me.DocumentNumberComboBox.Value & "'"
    Me.Filter = GetDocumentNumberFilter
    Me.RecordSource = "SELECT * FROM Purchases" _
        & " WHERE NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Purchases.BocName" _
        & " FROM Purchases WHERE" _
        & " NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetDocumentNumberFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetDocumentNumberFilter() As String
    On Error GoTo ErrorHandler:
    If Not DocumentNumberFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetDocumentNumberFilter = DocumentNumberFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & BocNameFilter
    End If
    If Not DocumentNumberFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetDocumentNumberFilter = DocumentNumberFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not DocumentNumberFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetDocumentNumberFilter = DocumentNumberFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not DocumentNumberFilter = vbNullString And _
        DivisionNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetDocumentNumberFilter = DocumentNumberFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & BocNameFilter
    End If
    If Not DocumentNumberFilter = vbNullString And _
        DivisionNameFilter = vbNullString And _
         FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetDocumentNumberFilter = DocumentNumberFilter _
                & pAnd & BocNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:            Event / Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    BocNameFilter = vbNullString
    BocNameFilter = "[BocName] = '" & Me.BocNameComboBox.Value & "'"
    Me.Filter = GetBocNameFilter
    Me.RecordSource = "SELECT * FROM Purchases" _
        & " WHERE NOT IsNull(DocumentControlNumbers)" _
        & pAnd & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:            Function
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not BocNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & DocumentNumberFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & DocumentNumberFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        DocumentNumberFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If BocNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetBocNameFilter = DocumentNumberFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter = vbNullString And _
        DocumentNumberFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter = vbNullString And _
        DocumentNumberFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not DocumentNumberFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & DocumentNumberFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter = vbNullString And _
        DocumentNumberFilter = vbNullString And _
        FundNameFilter = vbNullString And _
         FiscalYearFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    Me.Filter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ClearOptionGroupValues
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CompassDocumentQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetDetailedSpendingQuery", , , , , acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub OriginalActionDateTextBox_Enter()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.OriginalActionDateTextBox
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetColumnNames
'   Parameters:  Void
'   Retval:      Collection
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetColumnNames(pData As DAO.Recordset) As String()
    On Error GoTo ErrorHandler:
    Dim m_Fields As Collection
    Dim field As DAO.field
    Dim i As Integer
    Dim j As Integer
    Dim m_Array() As String
    Dim m_Data As DAO.Recordset
    Set m_Data = pData
    j = m_Data.Fields.count - 1
    ReDim m_Array(j)
    For i = LBound(m_Array()) To UBound(m_Array())
        If Not m_Data.Fields(i).Name & "" = "" Then
            m_Array(i) = m_Data.Fields(i).Name
        End If
    Next i
    GetColumnNames = m_Array()
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetSqlStatement()
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetSqlStatement() As String
    On Error GoTo ErrorHandler:
    Dim m_SQL As String
    m_SQL = "SELECT DISTINCT Purchases.BFY AS BFY, Purchases.FundCode AS Fund," _
        & " Purchases.AccountCode AS Account, Purchases.RcCode AS RC, Purchases.OrgCode AS ORG," _
        & " Purchases.DocumentControlNumbers AS DCN, Purchases.BocCode AS BOC, Purchases.OriginalActionDate AS [Date]," _
        & " SUM(Purchases.OpenCommitments) AS Commitments, SUM(Purchases.Obligations) AS Obligations" _
        & " FROM Purchases" _
        & " WHERE " & Me.Filter _
        & " AND NOT IsNull(Purchases.DocumentControlNumbers)" _
        & " GROUP BY Purchases.BFY, Purchases.FundCode, Purchases.AccountCode, Purchases.RcCode," _
        & " Purchases.OrgCode, Purchases.DocumentControlNumbers, Purchases.BocCode, Purchases.OriginalActionDate;"
    GetSqlStatement = m_SQL
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetReportData()
'   Parameters:  Void
'   Retval:      Collection
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetDataSet(pQuery As String) As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim m_Query As String
    m_Query = pQuery
    Dim m_Records As DAO.Recordset
    Set m_Records = CurrentDb.OpenRecordset(m_Query)
    Set GetDataSet = m_Records
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExcelButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExcelButton_Click()
    On Error GoTo ErrorHandler:
    Dim m_Data As DAO.Recordset
    Dim m_SqlQuery As String
    Dim m_Fields() As String
    Dim m_BudgetPath As BudgetPath
    Dim m_Excel As Excel.Application
    Dim m_Allocations As Excel.Workbook
    Dim m_Worksheet As Excel.Worksheet
    Dim m_List As Excel.ListObject
    Dim m_Range As Excel.Range
    Dim m_Cell As Object
    Dim m_Header As Excel.Range
    Dim m_Start As Excel.Range
    Dim m_End As Excel.Range
    Dim field As DAO.field
    Dim i As Integer
    Dim j As Integer
    m_SqlQuery = GetSqlStatement
    Set m_Data = CurrentDb.OpenRecordset(m_SqlQuery)
    m_Fields() = GetColumnNames(m_Data)
    Set m_BudgetPath = New BudgetPath
    Set m_Excel = CreateObject("Excel.Application")
    Set m_Allocations = m_Excel.Workbooks.Open(m_BudgetPath.ReportTemplate)
    m_Allocations.Worksheets(2).Visible = False
    Set m_Worksheet = m_Allocations.Worksheets(1)
    m_Worksheet.Name = "Regional Documents"
    m_Worksheet.Cells.HorizontalAlignment = xlHAlignLeft
    m_Worksheet.Cells.Font.Name = "Source Code Pro"
    m_Worksheet.Cells.Font.Size = 8
    Set m_Start = m_Worksheet.Cells(1, 1)
    Set m_End = m_Worksheet.Cells(1, UBound(m_Fields) + 1)
    Set m_Header = m_Worksheet.Range(m_Start, m_End)
    m_Header.Interior.Color = RGB(217, 217, 217)
    m_Header.HorizontalAlignment = xlHAlignLeft
    m_Header.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    m_Header.Borders(xlEdgeLeft).Color = RGB(133, 133, 133)
    m_Header.Borders(xlEdgeRight).Color = RGB(133, 133, 133)
    m_Header.Borders(xlEdgeTop).Color = RGB(133, 133, 133)
    m_Header.Font.Name = "Source Code Pro"
    m_Header.Font.Color = vbBlack
    m_Header.Font.Bold = True
    m_Header.Font.Size = 8
    For i = LBound(m_Fields) To UBound(m_Fields)
        j = i
        m_Header.Cells(1, i + 1).Value = m_Fields(j)
    Next i
    m_Worksheet.Cells(2, 1).CopyFromRecordset m_Data
    m_Header.Font.Color = vbBlack
    m_Excel.WindowState = xlMaximized
    m_Excel.Visible = True
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SaveButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UndoButton_Click()
    On Error GoTo ErrorHandler:
    If Me.Dirty Then
        DoCmd.RunCommand acCmdUndo
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SaveButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SaveButton_Click()
    On Error GoTo ErrorHandler:
    If Me.Dirty Then
        DoCmd.RunCommand acCmdSave
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DeleteButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DeleteButton_Click()
    On Error GoTo ErrorHandler:
    If Me.Dirty Then
        DoCmd.RunCommand acCmdDelete
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





