Option Compare Database
Option Explicit



Private PRC As AccountArgs
Private DivisionNameFilter As String
Private DocumentNumberFilter As String
Private FiscalYearFilter As String
Private BocNameFilter As String
Private pFormFilter As String
Private pAnd As String



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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     Form_Open()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Set PRC = New AccountArgs
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     SetButtonVisibility()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.DivisionNameComboBox.Value = vbNullString
    Me.DocumentNumberComboBox.Value = vbNullString
    Me.BocNameComboBox.Value = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     ClearComboBoxValues()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Private Sub ClearOptionGroupValues()
    On Error GoTo ErrorHandler:
    FiscalYearOptionGroup.Value = 2
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     ClearOptionGroupValues()" _
            & vbCrLf & "Descript:   " & Err.Description
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     ClearFilterValues()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.BFY = vbNullString
    FiscalYearFilter = vbNullString
    Me.Filter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            PRC.BFY = "2020"
            FiscalYearFilter = "[BFY] IN ('2019', '2020')"
        Case 2
            PRC.BFY = "2021"
            FiscalYearFilter = "[BFY] IN ('2021', '2022')"
    End Select
    Me.Filter = GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Purchases WHERE" _
        & " NOT IsNull(DocumentControlNumbers)" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     FiscalYearOptionGroup_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     GetFiscalYearFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    PRC.FundCode = Me.DivisionNameComboBox.Value
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     DivisionNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     GetDivisionNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     DocumentNumberComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     GetDocumentNumberFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   DocumentObligations" _
            & vbCrLf & "Member: BocNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   DocumentObligations" _
            & vbCrLf & "Member: GetBocNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     Form_Close()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ClearOptionGroupValues
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Private Sub CompassDocumentQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetDetailedSpendingQuery", , , , , acDialog
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     CompassDocumentQueryButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     CalculatorButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Private Sub OriginalActionDateTextBox_Enter()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.OriginalActionDateTextBox
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     OriginalActionDateTextBox_Enter()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member:     MenuButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
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
    Dim mFields As Collection
    Dim field As DAO.field
    Dim i As Integer
    Dim j As Integer
    Dim mArray() As String
    Dim mData As DAO.Recordset
    Set mData = pData
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
        mError = "Source:       RegionalDocuments" _
            & vbCrLf & "Member:     GetColumnNames" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    Dim mSQL As String
    mSQL = "SELECT DISTINCT Purchases.BFY AS BFY, Purchases.FundCode AS Fund," _
        & " Purchases.AccountCode AS Account, Purchases.RcCode AS RC, Purchases.OrgCode AS ORG," _
        & " Purchases.DocumentControlNumbers AS DCN, Purchases.BocCode AS BOC, Purchases.OriginalActionDate AS [Date]," _
        & " SUM(Purchases.OpenCommitments) AS Commitments, SUM(Purchases.Obligations) AS Obligations" _
        & " FROM Purchases" _
        & " WHERE " & Me.Filter _
        & " AND NOT IsNull(Purchases.DocumentControlNumbers)" _
        & " GROUP BY Purchases.BFY, Purchases.FundCode, Purchases.AccountCode, Purchases.RcCode," _
        & " Purchases.OrgCode, Purchases.DocumentControlNumbers, Purchases.BocCode, Purchases.OriginalActionDate;"
    GetSqlStatement = mSQL
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:       RegionalDocuments" _
            & vbCrLf & "Member:     GetSqlStatement()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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
    Dim mQuery As String
    mQuery = pQuery
    Dim mRecords As DAO.Recordset
    Set mRecords = CurrentDb.OpenRecordset(mQuery)
    Set GetReportData = mRecords
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:       RegionalDocuments" _
            & vbCrLf & "Member:     GetReportData()" _
            & vbCrLf & "Descript:   " & Err.Description
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
    Dim mData As DAO.Recordset
    Dim mSqlQuery As String
    Dim mFields() As String
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
    Dim i As Integer
    Dim j As Integer
    mSqlQuery = GetSqlStatement
    Set mData = CurrentDb.OpenRecordset(mSqlQuery)
    mFields() = GetColumnNames(mData)
    Set mBudgetPath = New BudgetPath
    Set mExcel = CreateObject("Excel.Application")
    Set mAllocations = mExcel.Workbooks.Open(mBudgetPath.ReportTemplate)
    mAllocations.Worksheets(2).Visible = False
    Set mWorksheet = mAllocations.Worksheets(1)
    mWorksheet.Name = "Regional Documents"
    mWorksheet.Cells.HorizontalAlignment = xlHAlignLeft
    mWorksheet.Cells.Font.Name = "Source Code Pro"
    mWorksheet.Cells.Font.Size = 8
    Set mStart = mWorksheet.Cells(1, 1)
    Set mEnd = mWorksheet.Cells(1, UBound(mFields) + 1)
    Set mHeader = mWorksheet.Range(mStart, mEnd)
    mHeader.Interior.Color = RGB(217, 217, 217)
    mHeader.HorizontalAlignment = xlHAlignLeft
    mHeader.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    mHeader.Borders(xlEdgeLeft).Color = RGB(133, 133, 133)
    mHeader.Borders(xlEdgeRight).Color = RGB(133, 133, 133)
    mHeader.Borders(xlEdgeTop).Color = RGB(133, 133, 133)
    mHeader.Font.Name = "Source Code Pro"
    mHeader.Font.Color = vbBlack
    mHeader.Font.Bold = True
    mHeader.Font.Size = 8
    For i = LBound(mFields) To UBound(mFields)
        j = i
        mHeader.Cells(1, i + 1).Value = mFields(j)
    Next i
    mWorksheet.Cells(2, 1).CopyFromRecordset mData
    mHeader.Font.Color = vbBlack
    mExcel.WindowState = xlMaximized
    mExcel.Visible = True
ErrorHandler:
    If Err.Number <> 0 Then
        mError = "Source:   RegionalDocuments" _
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
        mError = "Source:   RegionalDocuments" _
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
        mError = "Source:   RegionalDocuments" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



