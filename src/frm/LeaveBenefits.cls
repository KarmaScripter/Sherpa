Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------  FIELDS      ----------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------

Private LastNameFilter As String
Private DivisionNameFilter As String
Private pAnd As String
Private mError As String
Private mNotification As String

'---------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------         METHODS      ---------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------





'----------------------------------------------------------------------------------
'   Type:       Sub-Procedure
'   Name:
'   Parameters: Void
'   Retval:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Section(acHeader).AutoHeight = False
    Me.Section(acHeader).Height = 2
    Me.Section(acDetail).AutoHeight = False
    Me.Section(acDetail).Height = 4
    Me.Section(acFooter).AutoHeight = False
    Me.Section(acFooter).Height = 0.5
    DivisionNameOptionGroup.Visible = True
    ClearComboBoxValues
    Me.FilterOn = True
    SetButtonVisibility
    pAnd = " AND "
    Me.RecordSource = "SELECT * FROM EmployeeLeave;"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            SetButtonVisibility
'   Parameters:      Void
'   Purpose:         Toggles button visibility ON and OFF
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
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   LeaveBenefits" _
            & vbCrLf & "Member:     SetButtonVisibility()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Timer()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Timer()
    On Error GoTo ErrorHandler:
    HideButtons
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   LeaveBenefits" _
            & vbCrLf & "Member: Form_Timer()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Hides buttons
'   Parameters:  Void
'   Purpose:     Toggles the toolbar button on/off
'---------------------------------------------------------------------------------
Private Sub HideButtons()
    On Error GoTo ErrorHandler:
    Me.AddButton.Visible = False
    Me.DataButton.Visible = False
    Me.FirstButton.Visible = False
    Me.PreviousButton.Visible = False
    Me.NextButton.Visible = False
    Me.LastButton.Visible = False
    Me.EditButton.Visible = False
    Me.RefreshButton.Visible = False
    Me.DeleteButton.Visible = False
    Me.CalculatorButton.Visible = False
    Me.ExcelButton.Visible = False
    Me.UndoButton.Visible = False
    Me.SaveButton.Visible = False
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     HideButtons()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.LastNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.LastNameComboBox.BackColor = RGB(100, 100, 100)
    Me.LastNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.DivisionNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.DivisionNameComboBox.BackColor = RGB(100, 100, 100)
    Me.DivisionNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:       Sub-Procedure
'   Name:
'   Parameters: Void
'   Retval:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.LastNameComboBox.Value = vbNullString
    Me.DivisionNameComboBox.Value = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     ClearComboBoxValues()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:       Sub-Procedure
'   Name:
'   Parameters: Void
'   Retval:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     CalculatorButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:       Sub-Procedure
'   Name:
'   Parameters: Void
'   Retval:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     MenuButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:       Sub-Procedure
'   Name:
'   Parameters: Void
'   Retval:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    LastNameFilter = vbNullString
    DivisionNameFilter = vbNullString
    Me.Filter = vbNullString
    Me.RecordSource = "SELECT * FROM EmployeeLeave;"
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:       Sub-Procedure
'   Name:
'   Parameters: Void
'   Retval:
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub LastNameComboBox_Change()
    On Error GoTo ErrorHandler:
    LastNameFilter = vbNullString
    Me.Filter = vbNullString
    LastNameFilter = "[LastName] = '" & Me.LastNameComboBox.Value & "'"
    Me.Filter = GetLastNameFilter
    Me.RecordSource = "SELECT * FROM EmployeeLeave WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     LastNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:       Function
'   Name:       GetLastNameFilter
'   Parameters: Void
'   Retval:     String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetLastNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not LastNameFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter & pAnd & DivisionNameFilter
    End If
    If Not LastNameFilter = vbNullString And _
        DivisionNameFilter = vbNullString Then
            GetLastNameFilter = LastNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     GetLastNameFilter" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            DivisionNameComboBox_Change
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DivisionNameComboBox_Change()
    On Error GoTo ErrorHandler:
    DivisionNameFilter = vbNullString
    Me.Filter = vbNullString
    Dim Code As String
    Code = DLookup("Code", "Divisions", "Divisions.Title = '" & Me.DivisionNameComboBox.Value & "'")
    DivisionNameFilter = "[RcCode] = '" & Code & "'"
    Me.Filter = GetDivisionNameFilter
    Me.RecordSource = "SELECT * FROM EmployeeLeave WHERE " & Me.Filter
    Me.Requery
    Me.LastNameComboBox.RowSource = "SELECT DISTINCT EmployeeLeave.LastName" _
        & " FROM EmployeeLeave" _
        & " WHERE " & DivisionNameFilter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     DivisionNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:       Function
'   Name:
'   Parameters: Void
'   Retval:     String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetDivisionNameFilter() As String
    If Not DivisionNameFilter = vbNullString And _
        Not LastNameFilter = vbNullString Then
            GetDivisionNameFilter = DivisionNameFilter & pAnd & LastNameFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        LastNameFilter = vbNullString Then
            GetDivisionNameFilter = DivisionNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      LeaveBenefits" _
            & vbCrLf & "Member:     GetDivisionNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetColumnNames
'   Parameters:  Void
'   Retval:      Collection
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetColumnNames() As String()
    On Error GoTo ErrorHandler:
    Dim mFields As Collection
    Dim field As DAO.field
    Dim mData As DAO.Recordset
    Set mData = Me.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim mArray() As String
    j = mData.Fields.count - 1
    ReDim mArray(j)
    For i = LBound(mArray()) To UBound(mArray())
        If Not mData.Fields(i).Name & "" = "" Then
            mArray(i) = mData.Fields(i).Name
        End If
    Next i
    GetColumnNames = mArray()
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:       LeaveBenefits" _
            & vbCrLf & "Member:     GetColumnNames" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetReportData
'   Parameters:  Void
'   Retval:      DAO Recordset
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetReportData() As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim mData As DAO.Recordset
    Set mData = Me.Recordset
    mData.Filter = Me.Filter
    Set GetReportData = mData
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   LeaveBenefits" _
            & vbCrLf & "Member:     GetReportData" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Dim mFields() As String
    mFields() = GetColumnNames
    Dim mBudgetPath As BudgetPath
    Dim mExcel As Excel.Application
    Dim mAllocations As Excel.Workbook
    Dim mWorksheet As Excel.Worksheet
    Dim mList As Excel.ListObject
    Dim mRange As Excel.Range
    Dim mCell As Object
    Dim mHeader As Excel.Range
    Dim mStart As Excel.Range
    Dim mEnd As Excel.Range
    Dim field As DAO.field
    Dim mData As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Set mBudgetPath = New BudgetPath
    Set mExcel = CreateObject("Excel.Application")
    Set mAllocations = mExcel.Workbooks.Open(mBudgetPath.ReportTemplate)
    mAllocations.Worksheets(2).Visible = False
    Set mWorksheet = mAllocations.Worksheets(1)
    mWorksheet.Name = "Employee Leave"
    mWorksheet.Cells.HorizontalAlignment = xlHAlignLeft
    mWorksheet.Cells.Font.Name = "Source Code Pro"
    mWorksheet.Cells.Font.Size = 8
    Set mStart = mWorksheet.Cells(1, 1)
    Set mEnd = mWorksheet.Cells(1, UBound(mFields) - 1)
    Set mHeader = mWorksheet.Range(mStart, mEnd)
    mHeader.HorizontalAlignment = xlHAlignLeft
    mHeader.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    mHeader.Font.Name = "Source Code Pro"
    mHeader.Font.Color = vbBlack
    mHeader.Font.Bold = True
    mHeader.Font.Size = 8
    For i = LBound(mFields) To UBound(mFields)
        mHeader.Cells(i + 1).Value = mFields(i)
    Next i
    Set mData = Me.Recordset
    mData.Filter = Me.Filter
    mWorksheet.Cells(2, 1).CopyFromRecordset mData
    mHeader.Font.Color = vbBlack
    mExcel.WindowState = xlMaximized
    mExcel.Visible = True
ErrorHandler:
    If Err.Number <> 0 Then
        mError = "Source:   LeaveBenefits" _
            & vbCrLf & "Member:     ExcelButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
        Set mExcel = Nothing
        Set mAllocations = Nothing
        Set mWorksheet = Nothing
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   LeaveBenefits" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   LeaveBenefits" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



