Option Compare Database
Option Explicit



Public Args As PayrollArgs
Private pAnd As String
Private mFundCode As String
Private mFundName As String
Private FundNameFilter As String
Private mBFY As String
Private FiscalYearFilter As String
Private mAnd As String
Private mFilter As String
Private mError As String
Private mNotification As String


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    SetComboBoxColors
    pAnd = " AND "
    Set Args = New PayrollArgs
    Me.FiscalYearOptionGroup.DefaultValue = 2
    FiscalYearFilter = "[BFY] = '2021'"
    Args.BFY = "2021"
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT UCase(StatusOfFunds.FundName)" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND StatusOfFunds.BocCode = '10';"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   PayrollStatusDialog" _
            & vbCrLf & "Member:      Form_Load()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetComboBoxColors
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   FundDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        AcceptButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AcceptButton_Click()
    On Error GoTo ErrorHandler:
    Me.Visible = Not Me.Visible
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   FundDialog" _
            & vbCrLf & "Member:     AcceptButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExitButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    ClearFilterValues
    DoCmd.Close acForm, Me.Name
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   FundDialog" _
            & vbCrLf & "Member: ExitButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    mFundName = vbNullString
    Args.FundName = vbNullString
    Me.Filter = vbNullString
    FundNameFilter = vbNullString
    Args.FundName = Me.FundNameComboBox.Value
    Args.FundCode = DLookup("FundCode", "StatusOfFunds", "StatusOfFunds.FundName = '" _
        & Me.FundNameComboBox.Value & "'")
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   PayrollStatusDialog" _
            & vbCrLf & "Member:     FundNameComboBox_Change()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFundNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetFundNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   FundDialog" _
            & vbCrLf & "Member:     GetFundNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.BFY = vbNullString
    FiscalYearFilter = vbNullString
    Me.Filter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            Args.BFY = "2020"
            FiscalYearFilter = "[BFY] = '2020'"
        Case 2
            Args.BFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
    End Select
    Me.Filter = GetFiscalYearFilter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   FundDialog" _
            & vbCrLf & "Member:      FiscalYearOptionGroup_AfterUpdate()" _
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
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   FundDialog" _
            & vbCrLf & "Member: GetFiscalYearFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    Me.FiscalYearOptionGroup.Value = 2
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   FundDialog" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ClearComboBoxValues
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    ClearFilterValues
    Me.FundNameComboBox.Value = vbNullString
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT UCase(StatusOfFunds.FundName)" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BocCode = '10';"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   FundDialog" _
            & vbCrLf & "Member:     ClearComboBoxValues()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearFilterValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    FundNameFilter = vbNullString
    mFundName = vbNullString
    Args.FundName = vbNullString
    Args.FundCode = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   FundDialog" _
            & vbCrLf & "Member:     ClearFilterValues()" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


