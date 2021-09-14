Option Compare Database
Option Explicit



Public PRC As AccountArgs
Private mDialog As Form_DivisionSelector
Private pFilter As String
Private FiscalYearFilter As String
Private FundNameFilter As String
Private ProgramProjectNameFilter As String
Private BocNameFilter As String
Private RcCodeFilter As String
Private pAnd As String
Private mError As String
Private mNotification As String


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler::
    Set mDialog = New Form_DivisionSelector
    Set PRC = New AccountArgs
    DoCmd.OpenForm FormName:="DivisionSelector", WindowMode:=acDialog
    Set mDialog = Forms("DivisionSelector")
    Set PRC = mDialog.Args
    RcCodeFilter = "[RcCode] = '" & PRC.RcCode & "'"
    Me.RecordSource = "SELECT * FROM Purchases WHERE NOT" _
        & " IsNull(Purchases.DocumentControlNumbers)" _
        & " AND Purchases.BocCode NOT IN ('10', '17')" _
        & " AND Purchases.RcCode = '" & PRC.RcCode & "';"
    SetDivisionIcon
    DoCmd.Close ObjectType:=acForm, ObjectName:=mDialog.Name, Save:=acSaveNo
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DocumentObligations" _
            & vbCrLf & "Member:     Form_Open()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Load
'   Retval:      Void
'   Parameters:  Void
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
    ClearFilterValues
    ClearComboBoxValues
    SetComboBoxColors
    Me.FilterOn = True
    SetButtonVisibility
    pAnd = " AND "
    Me.Filter = RcCodeFilter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT Purchases.ProgramProjectName" _
        & " FROM Purchases" _
        & " WHERE" & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Purchases.FundName" _
        & " FROM Purchases" _
        & " WHERE" & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Purchases.BocName" _
        & " FROM Purchases" _
        & " WHERE" & Me.Filter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------


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
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: Form_Timer()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
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
    Me.CompassSpendingQueryButton.Visible = False
    Me.SaveButton.Visible = False
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: HideButtons()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearComboBoxValues
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.BocNameComboBox.Value = vbNullString
    Me.ProgramProjectNameComboBox.Value = vbNullString
    Me.FundNameComboBox.Value = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: ClearComboBoxValues()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearFilterValues
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    FiscalYearFilter = vbNullString
    FundNameFilter = vbNullString
    BocNameFilter = vbNullString
    ProgramProjectNameFilter = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: ClearFilterValues()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetDivisionIcon
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetDivisionIcon()
    On Error GoTo ErrorHandler::
    If Not PRC.RcCode = vbNullString Then
        Select Case PRC.RcCode
            Case "06A"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\ORA.png"
            Case "06B"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\LCARD.png"
            Case "06C"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\MSD.png"
            Case "06D"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\ORC.png"
            Case "06F"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\EJ.png"
            Case "06G"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\WCF.png"
            Case "06H"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\LSASD.png"
            Case "06J"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\ARD.png"
            Case "06K"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\WD.png"
            Case "06L"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\SEMD.png"
            Case "06M"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\ECAD.png"
            Case "06N"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\WSA.png"
            Case "06R"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\MSR.png"
            Case "06X"
                Me.DivisionIcon.Picture = CurrentProject.path & "\etc\png\DivisionLogo\XA.png"
        End Select
    Else
        Me.DivisionIcon.Picture = _
            CurrentProject.path & "\etc\png\AppIcons\interface\ui\Reports.png"
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      ProgrammaticObligations" _
            & vbCrLf & "Member:     SetDivisionIcon()" _
            & vbCrLf & "Descript: " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetComboBoxColors
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.ProgramProjectNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.ProgramProjectNameComboBox.BackColor = RGB(33, 33, 33)
    Me.ProgramProjectNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.BocNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.BocNameComboBox.BackColor = RGB(33, 33, 33)
    Me.BocNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetButtonVisibility
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
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
    Me.CompassSpendingQueryButton.Visible = Not Me.CompassSpendingQueryButton.Visible
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: SetButtonVisibility()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Close
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: Form_Close()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Retval:      Void
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
            FiscalYearFilter = "[BFY] = '2020'"
        Case 2
            PRC.BFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
    End Select
    Me.Filter = RcCodeFilter & pAnd & GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Purchases WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: FiscalYearOptionGroup_AfterUpdate()" _
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
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: GetFiscalYearFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProgramProjectNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProgramProjectNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    ProgramProjectNameFilter = vbNullString
    BocNameFilter = vbNullString
    PRC.ProgramProjectName = Me.ProgramProjectNameComboBox.Value
    ProgramProjectNameFilter = "[ProgramProjectName] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    Me.Filter = RcCodeFilter & pAnd & GetProgramProjectNameFilter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT BocName" _
        & " FROM Purchases" _
        & " WHERE Purchases.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.RecordSource = "SELECT * FROM Purchases WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: ProgramProjectNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetProgramProjectNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetProgramProjectNameFilter()
    On Error GoTo ErrorHandler:
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: GetProgramProjectNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        BocNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    BocNameFilter = vbNullString
    PRC.BocCode = Me.BocNameComboBox.Value
    BocNameFilter = "[BocName] = '" & Me.BocNameComboBox.Value & "'"
    Me.Filter = RcCodeFilter & pAnd & GetBocNameFilter
    Me.RecordSource = "SELECT * FROM Purchases WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: BocNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetBocNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = FiscalYearFilter & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetBocNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: GetBocNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    FundNameFilter = vbNullString
    Me.Filter = vbNullString
    PRC.FundName = Me.FundNameComboBox.Value
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = RcCodeFilter & pAnd & GetFundNameFilter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT Purchases.ProgramProjectName" _
        & " FROM Purchases" _
        & " WHERE Purchases.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Purchases.BocName" _
        & " FROM Purchases" _
        & " WHERE Purchases.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.RecordSource = "SELECT * FROM Purchases WHERE Purchases.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: FundNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Private Function GetFundNameFilter()
    On Error GoTo ErrorHandler:
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            Me.Filter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member:     GetFundNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        CalculatorButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: CalculatorButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        CompassSpendingQueryButton_Click()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'----------------------------------------------------------------------------------
Private Sub CompassSpendingQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetDetailedSpendingQuery"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: CompassSpendingQueryButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        MenuButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: MenuButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ResetOptionGroupValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
    On Error GoTo ErrorHandler:
    Me.FiscalYearOptionGroup.Value = 2
    FiscalYearFilter = "[BFY] = '2021'"
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionUtilization" _
            & vbCrLf & "Member:         ResetOptionGroupValues()" _
            & vbCrLf & "Descript:       " & Err.Description
            Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ResetOptionGroupValues
    Me.Filter = RcCodeFilter & pAnd & GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Purchases" _
        & " WHERE Purchases.BocCode IN ('36', '37', '38', '41')" _
        & pAnd & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Purchases.FundName" _
        & " FROM Purchases" _
        & " WHERE Purchases.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT Purchases.ProgramProjectName" _
        & " FROM Purchases" _
        & " WHERE Purchases.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Purchases.BocName" _
        & " FROM Purchases" _
        & " WHERE Purchases.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: RefreshButton_Click()" _
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
        mError = "Source:       ProgrammaticObligations" _
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
        mError = "Source:   ProgrammaticObligations" _
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
    mWorksheet.Name = "Programmatic Obligations"
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
        mError = "Source:   ProgrammaticObligations" _
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
        mError = "Source:   ProgrammaticObligations" _
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
        mError = "Source:   ProgrammaticObligations" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


