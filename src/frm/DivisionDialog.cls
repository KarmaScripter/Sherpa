Option Compare Database
Option Explicit


Public Args As AccountArgs
Public FiscalYearFilter As String
Public DivisionNameFilter As String
Private mRcCode As String
Private mBFY As String
Private pAnd As String
Private mError As String


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set Args = New AccountArgs
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member:     Form_Open()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    SetComboBoxColors
    pAnd = " AND "
    DivisionNameFilter = vbNullString
    Me.DivisionNameComboBox.RowSource = "SELECT DISTINCT Allocations.DivisionName" _
        & " FROM Allocations" _
        & " WHERE RcCode NOT IN ('06', '06R');"
    Me.FiscalYearOptionGroup.Value = 2
    Me.DivisionNameComboBox.Value = vbNullString
    Args.BFY = "2021"
    FiscalYearFilter = "[BFY] = '2021'"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member:     Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub






'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    mBFY = vbNull
    Me.Filter = vbNullString
    FiscalYearFilter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            mBFY = "2020"
            FiscalYearFilter = "[BFY] = '2020'"
            Args.BFY = mBFY
            Me.Filter = GetFiscalYearFilter
        Case 2
            mBFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
            Args.BFY = "2021"
            Me.Filter = GetFiscalYearFilter
    End Select
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member: Form_Load()" _
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
Private Function GetFiscalYearFilter() As String
    On Error GoTo ErrorHandler:
    If Not FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & DivisionNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & DivisionNameFilter
    End If
    If FiscalYearFilter = vbNullString And _
        Not DivisionNameFilter = vbNullString Then
            GetFiscalYearFilter = DivisionNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member:     Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DivisionNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DivisionNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    DivisionNameFilter = vbNullString
    DivisionNameFilter = "[DivisionName] = '" _
        & Me.DivisionNameComboBox.Value & "'"
    Args.DivisionName = Me.DivisionNameComboBox.Value
    Args.RcCode = DLookup("RcCode", "Allocations", "Allocations.DivisionName =" _
        & "'" & Me.DivisionNameComboBox.Value & "'")
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member:     DivisionNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetRcCodeFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetRcCodeFilter() As String
    On Error GoTo ErrorHandler:
    If DivisionNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetRcCodeFilter = FiscalYearFilter
    End If
    If Not DivisionNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetRcCodeFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member: GetRcCodeFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExecuteButton_OnClick
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AcceptButton_Click()
    On Error GoTo ErrorHandler:
    Me.Visible = Not Me.Visible
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member:     AcceptButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ClearOptionGroupValues
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
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
    Me!DivisionNameComboBox.ForeColor = RGB(255, 255, 255)
    Me!DivisionNameComboBox.BackColor = RGB(33, 33, 33)
    Me!DivisionNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member: SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearOptionGroupValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearOptionGroupValues()
    On Error GoTo ErrorHandler:
    Me.FiscalYearOptionGroup.Value = 2
    FiscalYearFilter = "[BFY] = '2021'"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member: ClearOptionGroupValues()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearComboBoxValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.DivisionNameComboBox.RowSource = "SELECT DISTINCT Allocations.DivisionName" _
        & " FROM Allocations" _
        & " WHERE RcCode NOT IN ('06', '06R');"
    Me.DivisionNameComboBox.Value = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member:     ClearComboBoxValues()" _
            & vbCrLf & "Descript:   " & Err.Description
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
    mBFY = vbNullString
    FiscalYearFilter = vbNullString
    mRcCode = vbNullString
    DivisionNameFilter = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member: ClearFilterValues()" _
            & vbCrLf & "Descript:   " & Err.Description
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
    DoCmd.Close acForm, Me.Name
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionSelector" _
            & vbCrLf & "Member: ExitButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



