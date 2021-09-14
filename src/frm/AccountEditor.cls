Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------  FIELDS      ----------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------

Public PRC As AccountArgs
Private mBudgetLevelFilter As String
Private mFiscalYearFilter As String
Private mFundNameFilter As String
Private mProgramProjectNameFilter As String
Private pAnd As String
Private mError As String


'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------  METHODS     ----------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            SetButtonVisibility
'   Parameters:      Void
'   Retval:          Void
'   Purpose:         Toggles button visibility ON and OFF
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.CopyButton.Visible = Not Me.CopyButton.Visible
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.DataWarehouseButton.Visible = Not Me.DataWarehouseButton.Visible
    Me.CalculatorButton.Visible = Not Me.CalculatorButton.Visible
    Me.UndoButton.Visible = Not Me.UndoButton.Visible
    Me.UpdateButton.Visible = Not Me.UpdateButton.Visible
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: SetButtonVisibility" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name            Update
'   Parameters      Void
'   Purpose         Defines and assigns the insert statement
'                   used to update the fields of a PRC
'---------------------------------------------------------------------------------
Public Sub Update()
    On Error GoTo ErrorHandler:
    Dim SQL As String
    Dim DB As DAO.Database
    Set DB = CurrentDb
    Dim Data As DAO.Recordset
    If Not Me.IdTextBox.Value & "" = "" Then
            PRC.ID = CInt(Me.IdTextBox.Value)
    End If
    If Not IsNull(Me.BudgetLevelTextBox) _
        And Not Me.BudgetLevelTextBox.Value & "" = "" Then
            PRC.BudgetLevel = Me.BudgetLevelTextBox.Value
    End If
    If Not IsNull(Me.BfyTextBox) _
        And Not Me.BfyTextBox.Value & "" = "" Then
            PRC.BFY = Me.BfyTextBox.Value
    End If
    If Not IsNull(Me.AhCodeTextBox) _
        And Not Me.AhCodeTextBox.Value & "" = "" Then
            PRC.AhCode = Me.AhCodeTextBox.Value
    End If
    If Not IsNull(Me.FundCodeTextBox) _
        And Not Me.FundCodeTextBox.Value & "" = "" Then
            PRC.FundCode = Me.FundCodeTextBox.Value
    End If
    If Not IsNull(Me.OrgCodeTextBox) _
        And Not Me.OrgCodeTextBox.Value & "" = "" Then
            PRC.OrgCode = Me.OrgCodeTextBox.Value
    End If
    If Not IsNull(Me.AccountCodeTextBox) _
        And Not Me.AccountCodeTextBox.Value = "" Then
            PRC.AccountCode = Me.AccountCodeTextBox.Value
    End If
    If Not IsNull(Me.BocCodeTextBox) _
        And Not Me.BocCodeTextBox.Value & "" = "" Then
            PRC.BocCode = Me.BocCodeTextBox.Value
    End If
    If Not IsNull(Me.RcCodeTextBox) _
        And Not Me.RcCodeTextBox.Value & "" = "" Then
            PRC.RcCode = Me.RcCodeTextBox.Value
    End If
    SQL = "UPDATE Allocations SET Allocations.AhCode = '" & PRC.AhCode & "'" _
        & pAnd & "Allocations.BudgetLevel = '" & PRC.BudgetLevel & "'" _
        & pAnd & "Allocations.BFY = '" & PRC.BFY & "'" _
        & pAnd & "Allocations.FundCode = '" & PRC.FundCode & "'" _
        & pAnd & "Allocations.OrgCode = '" & PRC.OrgCode & "'" _
        & pAnd & "Allocations.AccountCode = '" & PRC.AccountCode & "'" _
        & pAnd & "Allocations.BocCode = '" & PRC.BocCode & "'" _
        & pAnd & "Allocations.RcCode = '" & PRC.RcCode & "'" _
        & pAnd & "Allocations.Amount = '" & Me.ChangedAmountTextBox.Value & "'" _
        & "WHERE Allocations.PrcId = '" & PRC.ID & "'"
    Set Data = DB.OpenRecordset(SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: Update()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me!BudgetLevelComboBox.ForeColor = RGB(255, 255, 255)
    Me!BudgetLevelComboBox.BackColor = RGB(33, 33, 33)
    Me!BudgetLevelComboBox.BorderColor = RGB(68, 114, 196)
    Me!RpioCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!RpioCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!RpioCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!BfyComboBox.ForeColor = RGB(255, 255, 255)
    Me!BfyComboBox.BackColor = RGB(33, 33, 33)
    Me!BfyComboBox.BorderColor = RGB(68, 114, 196)
    Me!FundCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!FundCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!FundCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!AhCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!AhCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!AhCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!OrgCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!OrgCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!OrgCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!AccountCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!AccountCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!AccountCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!RcCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!RcCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!RcCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!ActivityCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!ActivityCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!ActivityCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!BocCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!BocCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!BocCodeComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub SetTextBoxColors()
    On Error GoTo ErrorHandler:
    Me!BudgetLevelTextBox.ForeColor = RGB(255, 255, 255)
    Me!BudgetLevelTextBox.BackColor = RGB(10, 10, 10)
    Me!BudgetLevelTextBox.BorderColor = RGB(52, 52, 52)
    Me!RpioCodeTextBox.ForeColor = RGB(255, 255, 255)
    Me.RpioCodeTextBox.BorderColor = RGB(52, 52, 52)
    Me.RpioCodeTextBox.BackColor = RGB(10, 10, 10)
    Me!BfyTextBox.ForeColor = RGB(255, 255, 255)
    Me!BfyTextBox.BackColor = RGB(10, 10, 10)
    Me!BfyTextBox.BorderColor = RGB(52, 52, 52)
    Me!FundCodeTextBox.ForeColor = RGB(255, 255, 255)
    Me!FundCodeTextBox.BackColor = RGB(10, 10, 10)
    Me!FundCodeTextBox.BorderColor = RGB(52, 52, 52)
    Me!AhCodeTextBox.ForeColor = RGB(255, 255, 255)
    Me!AhCodeTextBox.BackColor = RGB(10, 10, 10)
    Me!AhCodeTextBox.BorderColor = RGB(52, 52, 52)
    Me!OrgCodeTextBox.ForeColor = RGB(255, 255, 255)
    Me!OrgCodeTextBox.BackColor = RGB(10, 10, 10)
    Me!OrgCodeTextBox.BorderColor = RGB(52, 52, 52)
    Me!AccountCodeTextBox.ForeColor = RGB(255, 255, 255)
    Me!AccountCodeTextBox.BackColor = RGB(10, 10, 10)
    Me!AccountCodeTextBox.BorderColor = RGB(52, 52, 52)
    Me!RcCodeTextBox.ForeColor = RGB(255, 255, 255)
    Me!RcCodeTextBox.BackColor = RGB(10, 10, 10)
    Me!RcCodeTextBox.BorderColor = RGB(52, 52, 52)
    Me!ActivityCodeTextBox.ForeColor = RGB(255, 255, 255)
    Me!ActivityCodeTextBox.BackColor = RGB(10, 10, 10)
    Me!ActivityCodeTextBox.BorderColor = RGB(52, 52, 52)
    Me!BocCodeTextBox.ForeColor = RGB(255, 255, 255)
    Me!BocCodeTextBox.BackColor = RGB(10, 10, 10)
    Me!BocCodeTextBox.BorderColor = RGB(52, 52, 52)
    Me!AmountTextBox.ForeColor = RGB(255, 255, 255)
    Me!AmountTextBox.BackColor = RGB(10, 10, 10)
    Me!AmountTextBox.BorderColor = RGB(52, 52, 52)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: SetTextBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ClearFilters()
    On Error GoTo ErrorHandler:
    mBudgetLevelFilter = ""
    mFiscalYearFilter = ""
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: ClearFilters()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ClearTextBoxValues()
    On Error GoTo ErrorHandler:
    SetTextBoxColors
    Me.BfyTextBox = Null
    Me.AccountCodeTextBox = Null
    Me.BudgetLevelTextBox = Null
    Me.RpioCodeTextBox = Null
    Me.BocCodeTextBox = Null
    Me.RcCodeTextBox = Null
    Me.BocCodeTextBox = Null
    Me.AhCodeTextBox = Null
    Me.ActivityCodeTextBox = Null
    Me.FundCodeTextBox = Null
    Me.OrgCodeTextBox = Null
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: ClearTextBoxValues()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    SetComboBoxColors
    Me.BfyComboBox.Value = ""
    Me.AccountCodeComboBox.Value = ""
    Me.BudgetLevelComboBox.Value = ""
    Me.RpioCodeComboBox.Value = ""
    Me.BocCodeComboBox.Value = ""
    Me.RcCodeComboBox.Value = ""
    Me.BocCodeComboBox.Value = ""
    Me.AhCodeComboBox.Value = ""
    Me.ActivityCodeComboBox.Value = ""
    Me.FundCodeComboBox.Value = ""
    Me.OrgCodeComboBox.Value = ""
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: ClearComboBoxValues()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name            ExitButton_Click
'   Parameters      Void
'   Retval
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    Dim mOwner As String
    mOwner = Me.Tag
    DoCmd.Close ObjectType:=acForm, _
        ObjectName:="AccountEditor", Save:=acSaveYes
    DoCmd.OpenForm FormName:=mOwner
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member:     ExitButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name            Event to initializes members
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Set PRC = New AccountArgs
    pAnd = " AND "
    SetButtonVisibility
    ClearComboBoxValues
    Me.IdTextBox.Value = Me.OpenArgs
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ActivityCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!ActivityCodeTextBox.Value = Me.ActivityCodeComboBox.SelText
    Me.ActivityCodeTextBox.BorderColor = RGB(68, 114, 196)
    Me.ActivityCodeTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: ActivityCodeComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub AhCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!AhCodeTextBox.Value = Me.AhCodeComboBox.SelText
    PRC.AhCode = Me!AhCodeTextBox.Value
    Me.AhCodeTextBox.BorderColor = RGB(68, 114, 196)
    Me.AhCodeTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: AhCodeComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: MenuButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub OrgCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!OrgCodeTextBox.Value = Me.OrgCodeComboBox.SelText
    PRC.OrgCode = Me!OrgCodeTextBox.Value
    Me.OrgCodeTextBox.BorderColor = RGB(68, 114, 196)
    Me.OrgCodeTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: OrgCodeComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub RcCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!RcCodeTextBox.Value = Me.RcCodeComboBox.SelText
    PRC.RcCode = Me!RcCodeTextBox.Value
    Me.RcCodeTextBox.BorderColor = RGB(68, 114, 196)
    Me.RcCodeTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: RcCodeComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub BfyComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!BfyTextBox.Value = Me.BfyComboBox.SelText
    PRC.BFY = Me!BfyTextBox.Value
    Me.BfyTextBox.BorderColor = RGB(68, 114, 196)
    Me.BfyTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: BfyComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub BocCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!BocCodeTextBox.Value = Me.BocCodeComboBox.SelText
    PRC.BocCode = Me!BocCodeTextBox.Value
    Me.BocCodeTextBox.BorderColor = RGB(68, 114, 196)
    Me.BocCodeTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: BocCodeComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub BudgetLevelComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!BudgetLevelTextBox.Value = Me.BudgetLevelComboBox.SelText
    PRC.BudgetLevel = Me!BudgetLevelTextBox.Value
    Me.BudgetLevelTextBox.BorderColor = RGB(68, 114, 196)
    Me.BudgetLevelTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: BudgetLevelComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub RpioCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!RpioCodeTextBox.Value = Me.RpioCodeComboBox.SelText
    PRC.RpioCode = Me!RpioCodeTextBox.Value
    Me.RpioCodeTextBox.BorderColor = RGB(68, 114, 196)
    Me.RpioCodeTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: RpioCodeComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FundCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!FundCodeTextBox.Value = Me.FundCodeComboBox.SelText
    PRC.FundCode = Me!FundCodeTextBox.Value
    Me.FundCodeTextBox.BorderColor = RGB(68, 114, 196)
    Me.FundCodeTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: FundCodeComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub AccountCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me!AccountCodeTextBox.Value = Me.AccountCodeComboBox.SelText
    PRC.AccountCode = Me!AccountCodeTextBox.Value
    Me.AccountCodeTextBox.BorderColor = RGB(68, 114, 196)
    Me.AccountCodeTextBox.BackColor = RGB(33, 33, 33)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: AccountCodeComboBox_AfterUpdate()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------

Private Sub UpdateButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.SetWarnings (False)
    DoCmd.OpenQuery ("UpdateAccountEditorData")
    DoCmd.SetWarnings (True)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearTextBoxValues
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: RefreshButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: CalculatorButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Retval          Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub CopyButton_Click()
    On Error GoTo ErrorHandler:
    Dim pSQL As String
    Dim pAhName As String
    Dim pOrgName As String
    Dim pFundName As String
    Dim pBocName As String
    Dim pProgramProjectCode As String
    Dim pProgramAreaCode As String
    Dim pProgramProjectName As String
    Dim pDivisionName As String
    Dim pNpmCode As String
    Dim pNpmName As String
    Dim pGoalCode As String
    Dim pGoalName As String
    Dim pObjectiveCode As String
    Dim pObjectiveName As String
    pProgramProjectCode = mID(Me.AccountCodeTextBox, 5, 2)
    pGoalCode = Left(Me.AccountCodeTextBox, 1)
    pObjectiveCode = mID(Me.AccountCodeTextBox, 1, 2)
    pNpmCode = mID(Me.AccountCodeTextBox, 4, 1)
    pProgramProjectName = DLookup("ProgramProjectName", "Allocations", "[ProgramProjectCode]= '" & pProgramProjectCode & "'")
    pFundName = DLookup("FundName", "Allocations", "[FundCode]= '" & Me.FundCodeTextBox & "'")
    pBocName = DLookup("BocName", "Allocations", "[BocCode]= '" & Me.BocCodeTextBox & "'")
    pNpmName = DLookup("BocName", "Allocations", "[NpmCode]= '" & pNpmCode & "'")
    pDivisionName = DLookup("DivisionName", "Allocations", "[RcCode]= '" & Me.RcCodeTextBox & "'")
    pGoalName = DLookup("GoalName", "Allocations", "[GoalCode]= '" & pGoalCode & "'")
    pObjectiveName = DLookup("ObjectiveName", "Allocations", "[ObjectiveCode]= '" & pObjectiveCode & "'")
    pAhName = DLookup("AhName", "Allocations", "[AhCode]= '" & Me.AhCodeTextBox & "'")
    pOrgName = DLookup("ORgName", "Allocations", "[OrgCode]= '" & Me.OrgCodeTextBox & "'")
    pSQL = "INSERT INTO Allocations (BudgetLevel, BFY, RPIO, AhCode, FundCode, OrgCode," _
        & " AccountCode, BocCode, RcCode, ActivityCode)" _
        & " VALUES (Me.BudgetLevelTextBox.Text, Me.BfyTextBox.Text, Me.RpioTextBox.Text," _
        & " Me.AhCodeTextBox.Text, Me.FundCodeTextBox.Text, Me.OrgTextBox.Text," _
        & " Me.AccountCodeTextBox.Text, Me.BocCodeTextBox.Text, Me.RcCodeTextBox.Text, Me.ActivityCodeTextBox.Text)"
    DoCmd.RunSQL (pSQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: CopyButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub InsertButton_Click()
    On Error GoTo ErrorHandler:
    Dim pSQL As String
    Dim pAhName As String
    Dim pOrgName As String
    Dim pFundName As String
    Dim pBocName As String
    Dim pProgramProjectCode As String
    Dim pProgramAreaCode As String
    Dim pProgramProjectName As String
    Dim pDivisionName As String
    Dim pNpmCode As String
    Dim pNpmName As String
    Dim pGoalCode As String
    Dim pGoalName As String
    Dim pObjectiveCode As String
    Dim pObjectiveName As String
    pProgramProjectCode = mID(Me.AccountCodeTextBox, 5, 2)
    pGoalCode = Left(Me.AccountCodeTextBox, 1)
    pObjectiveCode = mID(Me.AccountCodeTextBox, 1, 2)
    pNpmCode = mID(Me.AccountCodeTextBox, 4, 1)
    pProgramProjectName = DLookup("ProgramProjectName", "Allocations", "[ProgramProjectCode]= '" & pProgramProjectCode & "'")
    pFundName = DLookup("FundName", "Allocations", "[FundCode]= '" & Me.FundCodeTextBox & "'")
    pBocName = DLookup("BocName", "Allocations", "[BocCode]= '" & Me.BocCodeTextBox & "'")
    pNpmName = DLookup("BocName", "Allocations", "[NpmCode]= '" & pNpmCode & "'")
    pDivisionName = DLookup("DivisionName", "Allocations", "[RcCode]= '" & Me.RcCodeTextBox & "'")
    pGoalName = DLookup("GoalName", "Allocations", "[GoalCode]= '" & pGoalCode & "'")
    pObjectiveName = DLookup("ObjectiveName", "Allocations", "[ObjectiveCode]= '" & pObjectiveCode & "'")
    pAhName = DLookup("AhName", "Allocations", "[AhCode]= '" & Me.AhCodeTextBox & "'")
    pOrgName = DLookup("ORgName", "Allocations", "[OrgCode]= '" & Me.OrgCodeTextBox & "'")
    pSQL = "INSERT INTO Allocations (BudgetLevel, BFY, RPIO, AhCode, FundCode, OrgCode, AccountCode, BocCode, RcCode, ActivityCode) " _
        & "VALUES (Me.BudgetLevelTextBox.Text, Me.BfyTextBox.Text, Me.RpioTextBox.Text, Me.AhCodeTextBox.Text, Me.FundCodeTextBox.Text, Me.OrgTextBox.Text, " _
        & "Me.AccountCodeTextBox.Text, Me.BocCodeTextBox.Text, Me.RcCodeTextBox.Text, Me.ActivityCodeTextBox.Text)"
    DoCmd.RunSQL (pSQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: InsertButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub UpdateBudgetLevel()
    On Error GoTo ErrorHandler:
    Me.IdTextBox.SetFocus
    PRC.ID = CInt(Me!IdTextBox.text)
    Me.BudgetLevelTextBox.SetFocus
    PRC.BudgetLevel = Me.BudgetLevelTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.BudgetLevel = '" _
        & PRC.BudgetLevel & "'" & _
        " WHERE Allocations.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateBudgetLevel()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub UpdateBfy()
    On Error GoTo ErrorHandler:
    Me.IdTextBox.SetFocus
    PRC.ID = CInt(Me!IdTextBox.text)
    Me.BfyTextBox.SetFocus
    PRC.BFY = Me.BfyTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.BFY = '" _
        & PRC.BFY & "'" & _
        " WHERE Allocations.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateBfy()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub UpdateRpioCode()
    On Error GoTo ErrorHandler:
    Me.IdTextBox.SetFocus
    PRC.ID = CInt(Me!IdTextBox.text)
    Me.RpioCodeTextBox.SetFocus
    PRC.RpioCode = Me.RpioCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.RPIO = '" _
        & PRC.RpioCode & "'" & _
        " WHERE Allocations.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateRpioCode()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub UpdateAhCode()
    On Error GoTo ErrorHandler:
    Me.IdTextBox.SetFocus
    PRC.ID = CInt(Me!IdTextBox.text)
    Me.AhCodeTextBox.SetFocus
    PRC.AhCode = Me.AhCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.AhCode = '" _
        & PRC.AhCode & "'" & _
        " WHERE Allocations.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateAhCode()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub UpdateOrgCode()
    On Error GoTo ErrorHandler:
    Me.IdTextBox.SetFocus
    PRC.ID = CInt(Me!IdTextBox.text)
    Me.OrgCodeTextBox.SetFocus
    PRC.OrgCode = Me.OrgCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.OrgCode = '" _
        & PRC.OrgCode & "'" & _
        " WHERE Allocations.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateOrgCode()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub UpdateAccountCode()
    On Error GoTo ErrorHandler:
    Me.IdTextBox.SetFocus
    PRC.ID = CInt(Me!IdTextBox.text)
    Me.AccountCodeTextBox.SetFocus
    PRC.AccountCode = Me.AccountCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.AccountCode = '" _
        & PRC.AccountCode & "'" & _
        " WHERE Allocations.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateAccountCode()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub UpdateBocCode()
    On Error GoTo ErrorHandler:
    Me.IdTextBox.SetFocus
    PRC.ID = CInt(Me!IdTextBox.text)
    Me.BocCodeTextBox.SetFocus
    PRC.BocCode = Me.BocCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.BocCode = '" _
        & PRC.BocCode & "'" & _
        " WHERE Allocations.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateBocCode()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub UpdateRcCode()
    On Error GoTo ErrorHandler:
    Me.IdTextBox.SetFocus
    PRC.ID = CInt(Me!IdTextBox.text)
    Me.RcCodeTextBox.SetFocus
    PRC.RcCode = Me.RcCodeTextBox.text
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.RcCode = '" _
        & PRC.RcCode & "'" & _
        " WHERE Allocations.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateRcCode()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type            Sub-Procedure
'   Name
'   Parameters      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub UpdateAmount()
    On Error GoTo ErrorHandler:
    Me.IdTextBox.SetFocus
    PRC.ID = CInt(Me!IdTextBox.text)
    Me.ChangedAmountTextBox.SetFocus
    Dim Change As Double
    Change = CDbl(Me.ChangedAmountTextBox.text)
    Dim SQL As String
    SQL = "UPDATE Allocations SET Allocations.Amount = " _
        & Change & _
        " WHERE Allocations.PrcId = " & PRC.ID
    Dim DB As DAO.Database
    Set DB = CurrentDb
    DB.Execute (SQL)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   AccountEditor" _
            & vbCrLf & "Member: UpdateAmount()" _
            & vbCrLf & "Descript:   " & Err.Description
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
        mError = "Source:   AllocationForm" _
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
        mError = "Source:   AllocationForm" _
            & vbCrLf & "Member: SaveButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


