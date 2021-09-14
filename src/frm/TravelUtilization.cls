Option Compare Database
Option Explicit



Public DbData As DAO.Database
Public DbSet As DAO.Recordset
Public Args As AccountArgs
Public mDialog As Form_DivisionSelector
Private DivisionNameFilter As String
Private FiscalYearFilter As String
Private FundNameFilter As String
Private ProgramProjectNameFilter As String
Private BocNameFilter As String
Private pAnd As String
Private mError As String
Private mNotification As String
Private SQL As String




'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        Form_Open
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set mDialog = New Form_DivisionSelector
    DoCmd.OpenForm FormName:="DivisionSelector", WindowMode:=acDialog
    pAnd = " AND "
    FiscalYearOptionGroup.Value = 2
    FiscalYearFilter = "[BFY] = '2021'"
    Set Args = New AccountArgs
    Set mDialog = Forms("DivisionSelector")
    Set Args = mDialog.Args
    Args.BFY = "2021"
    If Not Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
            FiscalYearFilter = "[BFY] = '" & Args.BFY & "'"
            FundNameFilter = "[FundName] = '" & Args.FundName & "'"
            Me.Filter = DivisionNameFilter & _
                pAnd & FiscalYearFilter & pAnd & FundNameFilter
            SQL = "SELECT * FROM DivisionExecution" _
                & " WHERE DivisionExecution.BocCode IN ('21','28')" & _
                pAnd & Me.Filter
    End If
    If Args.DivisionName & "" = "" And _
        Not Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            FiscalYearFilter = "[BFY] = '" & Args.BFY & "'"
            FundNameFilter = "[FundName] = '" & Args.FundName & "'"
            Me.Filter = FiscalYearFilter & pAnd & FundNameFilter
            SQL = "SELECT * FROM DivisionExecution" _
                & " WHERE DivisionExecution.BocCode IN ('21','28')" & _
                pAnd & Me.Filter
    End If
    If Not Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            DivisionNameFilter = "[DivisionName] = '" & Args.DivisionName & "'"
            FiscalYearFilter = "[BFY] = '" & Args.BFY & "'"
            Me.Filter = DivisionNameFilter & _
                pAnd & FiscalYearFilter
            SQL = "SELECT * FROM DivisionExecution" _
                & " WHERE DivisionExecution.BocCode IN ('21','28')" & _
                pAnd & Me.Filter
    End If
    If Args.DivisionName & "" = "" And _
        Args.FundName & "" = "" And _
        Not Args.BFY & "" = "" Then
            FiscalYearFilter = "[BFY] = '" & Args.BFY & "'"
            Me.Filter = FiscalYearFilter
            SQL = "SELECT * FROM DivisionExecution" _
                & " WHERE DivisionExecution.BocCode IN ('21','28')" & _
                pAnd & Me.Filter
    End If
    Me.RecordSource = SQL
    DoCmd.Close ObjectType:=acForm, ObjectName:=mDialog.Name, Save:=acSaveNo
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelUtilization" _
            & vbCrLf & "Member:     Form_Open()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:        Sub
'   Name:        Form_Load()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    SetDivisionIcon
    Me.DivisionTitle.Caption = Args.DivisionName
    Me.Caption = "Travel Utilization"
    SetButtonVisibility
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.ProgramProjectName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('21', '28')" & pAnd & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.FundName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('21', '28')" & pAnd & Me.Filter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      TravelUtilization" _
            & vbCrLf & "Member:     Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub
'   Name:        Form_Close()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    DivisionNameFilter = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      TravelUtilization" _
            & vbCrLf & "Member:     Form_Close()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub
'   Name:        SetButtonVisibility()
'   Parameters:  Void
'   Retval:      Void
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
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      TravelUtilization" _
            & vbCrLf & "Member:     SetButtonVisibility()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub
'   Name:        FiscalYearOptionGroup_AfterUpdate()
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
    Me.Filter = DivisionNameFilter & pAnd & GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('21', '28')" & _
            pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      TravelUtilization" _
            & vbCrLf & "Member:     FiscalYearOptionGroup_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetFiscalYearFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetFiscalYearFilter() As String
    On Error GoTo ErrorHandler:
    If Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FundNameFilter = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelUtilization" _
            & vbCrLf & "Member:     GetFiscalYearFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        FundNameComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    FundNameFilter = vbNullString
    Me.ProgramProjectNameComboBox.Value = vbNullString
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetFundNameFilter
    Me.RecordSource = "SELECT * FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('21', '28')" _
        & pAnd & Me.Filter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.ProgramProjectName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('21', '28')" _
        & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravellUtilization" _
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
Private Function GetFundNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = vbNullString And _
        Not ProgramProjectNameFilter = vbNullString And _
        FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter = vbNullString And _
        ProgramProjectNameFilter = vbNullString And _
        Not FiscalYearFilter = vbNullString Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelUtilization" _
            & vbCrLf & "Member:     GetFundNameFilter()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        ProgramProjectNameComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ProgramProjectNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    ProgramProjectNameFilter = vbNullString
    ProgramProjectNameFilter = "[ProgramProjectName] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    Me.Filter = DivisionNameFilter & pAnd & GetProgramProjectNameFilter
    Me.RecordSource = "SELECT * FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('21', '28')" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelUtilization" _
            & vbCrLf & "Member:     ProgramProjectNameComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub






'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetProgramProjectNameFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetProgramProjectNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not ProgramProjectNameFilter = "" And _
        Not FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        FundNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        Not FundNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        Not FundNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        FundNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter = "" And _
        Not FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      TravelUtilization" _
            & vbCrLf & "Member:     GetProgramProjectNameFilter()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        RefreshButton_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    Me.FiscalYearOptionGroup.Value = 2
    Me.Filter = DivisionNameFilter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.ProgramProjectName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('21', '28')" & pAnd & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT DivisionExecution.FundName" _
        & " FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('21', '28')" & pAnd & Me.Filter
    Me.RecordSource = "SELECT * FROM DivisionExecution" _
        & " WHERE DivisionExecution.BocCode IN ('21', '28')" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelUtilization" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        ClearFilterValues
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    FundNameFilter = vbNullString
    ProgramProjectNameFilter = vbNullString
    FiscalYearFilter = "[BFY] = '2021'"
    Me.Filter = DivisionNameFilter & pAnd & FiscalYearFilter
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelUtilization" _
            & vbCrLf & "Member:     ClearFilterValues()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        ClearComboBoxValues
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.Value = vbNullString
    Me.ProgramProjectNameComboBox.Value = vbNullString
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   TravelUtilization" _
            & vbCrLf & "Member:     ClearComboBoxValues()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub
'   Name:        MenuButton_Click()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      TravelUtilization" _
            & vbCrLf & "Member:     MenuButton_Click()" _
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
    If Not Args.RcCode = vbNullString Then
        Select Case Args.RcCode
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
        mError = "Source:      TravelUtilization" _
            & vbCrLf & "Member:     SetDivisionIcon()" _
            & vbCrLf & "Descript: " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


