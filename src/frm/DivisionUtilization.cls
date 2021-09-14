Option Compare Database
Option Explicit


Public DbData As DAO.Database
Public DbSet As DAO.Recordset
Public PRC As AccountArgs
Public mDialog As Form_DivisionSelector
Private BocNameFilter As String
Private FiscalYearFilter As String
Private FundNameFilter As String
Private ProgramProjectNameFilter As String
Private RcCodeFilter
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
    On Error GoTo ErrorHandler:
    Set mDialog = New Form_DivisionSelector
    Set PRC = New AccountArgs
    DoCmd.OpenForm FormName:="DivisionSelector", WindowMode:=acDialog
    Set mDialog = Forms("DivisionSelector")
    Set PRC = mDialog.Args
    RcCodeFilter = "[RcCode] = '" & PRC.RcCode & "'"
    Me.RecordSource = "SELECT * FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" _
        & " AND DivisionExecution.RcCode = '" & PRC.RcCode & "';"
    DoCmd.Close ObjectType:=acForm, ObjectName:="DivisionSelector", Save:=acSaveNo
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionUtilization" _
            & vbCrLf & "Member:     Form_Open()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    pAnd = " AND "
    HideButtons
    SetDivisionIcon
    ClearComboBoxValues
    Me.FiscalYearOptionGroup.DefaultValue = 2
    PRC.BFY = "2021"
    FiscalYearFilter = "[BFY] = '2021'"
    Me.Filter = RcCodeFilter & pAnd & GetFiscalYearFilter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.FundName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.ProgramProjectName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.BocName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionUtilization" _
            & vbCrLf & "Member:     Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Close
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    Me.Filter = ""
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     Form_Close()" _
            & vbCrLf & "Descript: " & Err.Description
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
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     SetButtonVisibility()" _
            & vbCrLf & "Descript: " & Err.Description
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
    On Error GoTo ErrorHandler:
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
        mError = "Source:      DivisionAllocations" _
            & vbCrLf & "Member:     SetDivisionIcon()" _
            & vbCrLf & "Descript: " & Err.Description
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
        mError = "Source:   DivisionUtilization" _
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
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     HideButtons()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    FiscalYearFilter = vbNullString
    Me.Filter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            FiscalYearFilter = "[BFY] = '2020'"
            PRC.BFY = "2020"
        Case 2
            FiscalYearFilter = "[BFY] = '2021'"
            PRC.BFY = "2021"
    End Select
    Me.Filter = RcCodeFilter & pAnd & GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     FiscalYearOptionGroup_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFiscalYear
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetFiscalYearFilter()
    On Error GoTo ErrorHandler:
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     GetFiscalYearFilter()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Sub Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    FundNameFilter = vbNullString
    Me.Filter = vbNullString
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = RcCodeFilter & pAnd & GetFundNameFilter
    Me.RecordSource = "SELECT * FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" _
        & pAnd & Me.Filter
    Me.Requery
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.BocName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" _
        & pAnd & Me.Filter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     FundNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
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
        ProgramProjectNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     GetFundNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Sub Procedure
'   Name:        ProgramProjectNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProgramProjectNameComboBox_Change()
    On Error GoTo ErrorHandler:
    ProgramProjectNameFilter = vbNullString
    Me.Filter = vbNullString
    ProgramProjectNameFilter = "[ProgramProjectName] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    Me.Filter = RcCodeFilter & pAnd & GetProgramProjectNameFilter
    Me.RecordSource = "SELECT * FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" _
        & pAnd & Me.Filter
    Me.Requery
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.FundName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.BocName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" _
        & pAnd & Me.Filter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     ProgramProjectNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
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
        BocNameFilter = vbNullString Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        BocNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not FundNameFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        FundNameFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetProgramProjectNameFilter = BocNameFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & BocNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     GetProgramProjectNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function






'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        BocNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    BocNameFilter = vbNullString
    Me.Filter = vbNullString
    BocNameFilter = "[BocName] = '" & Me.BocNameComboBox.Value & "'"
    Me.Filter = RcCodeFilter & pAnd & GetBocNameFilter
    Me.RecordSource = "SELECT * FROM DivisionExecution WHERE" & Me.Filter
    Me.Requery
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.FundName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     BocNameComboBox_Change()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetBocNameFilter
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter & pAnd & FiscalYearFilter _
                & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter & pAnd & FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString And _
        Not BocNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & BocNameFilter
    End If
    If Not BocNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString And _
        BocNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     GetBocNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearFilters
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilters()
    On Error GoTo ErrorHandler:
    ProgramProjectNameFilter = vbNullString
    FiscalYearFilter = vbNullString
    BocNameFilter = vbNullString
    FundNameFilter = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionUtilization" _
            & vbCrLf & "Member:     ClearFilters()" _
            & vbCrLf & "Descript:       " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearComboBoxes
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.ProgramProjectNameComboBox.Value = vbNullString
    Me.FundNameComboBox.Value = vbNullString
    Me.BocNameComboBox.Value = vbNullString
    ClearFilters
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   DivisionUtilization" _
            & vbCrLf & "Member:     ClearComboBoxes" _
            & vbCrLf & "Descript:       " & Err.Description
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
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



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
    ClearFilters
    ResetOptionGroupValues
    Me.Filter = RcCodeFilter & pAnd & GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" _
        & pAnd & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.FundName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.ProgramProjectName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.BocName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:       DivisionUtilization" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript:       " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
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
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     MenuButton_Click()" _
            & vbCrLf & "Descript:       " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
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
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     CalculatorButton_Click()" _
            & vbCrLf & "Descript:       " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ExitButton_Click()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, Me.Name, acSaveNo
    DoCmd.OpenForm FormName:="DataAnalysis", WindowMode:=acDialog
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      DivisionUtilization" _
            & vbCrLf & "Member:     ExitButton_Click()" _
            & vbCrLf & "Descript:       " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



