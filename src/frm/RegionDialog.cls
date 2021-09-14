Option Compare Database
Option Explicit



Private mBFY As String
Private mBudgetLevel As String
Private mFundCode As String
Private mRcCode As String
Private mDivisionName As String
Private mAccountCode As String
Private mBocCode As String
Public DialogArgs As BudgetArgs
Private mOptionGroupFilter As String
Private mFirstFilter As String
Private mSecondFilter As String
Private mDialogFilter As String
Private pAnd As String




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    SetComboBoxColors
    pAnd = " AND "
    mBudgetLevel = "8"
    SetOptionGroupFilter
    Me.FirstComboBox.RowSource = "SELECT DISTINCT Divisions.Title FROM Divisions " _
        & "WHERE Divisions.Code NOT IN ('06', '06R')"
    Me.FirstComboBox.Value = ""
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetComboBoxColors
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    Me!FirstComboBox.ForeColor = RGB(255, 255, 255)
    Me!FirstComboBox.BackColor = RGB(33, 33, 33)
    Me!FirstComboBox.BorderColor = RGB(68, 114, 196)
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ResetOptionGroupValues
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
    If Not IsNull(Me.DialogOptionGroup) And _
        Not Me.DialogOptionGroup.Value & "" = "" Then
        Me.DialogOptionGroup.Value = ""
        mOptionGroupFilter = ""
    End If
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetOptionGroupFilter
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetOptionGroupFilter()
    mOptionGroupFilter = ""
    mBFY = ""
    Select Case Me.DialogOptionGroup.Value
        Case 1
            mOptionGroupFilter = "[BFY] = '2020'"
            mBFY = "2020"
        Case 2
            mOptionGroupFilter = "[BFY] = '2021'"
            mBFY = "2021"
    End Select
End Sub

'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DialogOptionGroup_AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DialogOptionGroup_AfterUpdate()
    mOptionGroupFilter = ""
    mBFY = ""
    DialogArgs.BFY = ""
       
    'Set the PRC parameters
    Select Case Me.DialogOptionGroup.Value
        Case 1
            mOptionGroupFilter = "[BFY] = '2020'"
            mBFY = "2020"
            DialogArgs.BFY = mBFY
        Case 2
            mOptionGroupFilter = "[BFY] = '2021'"
            mBFY = "2021"
            DialogArgs.BFY = mBFY
    End Select
End Sub


Private Sub Form_Open(Cancel As Integer)
    Set DialogArgs = New BudgetArgs
    Me.FirstComboBox.Value = ""
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        RefreshButton_OnClick
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    mFirstFilter = ""
    mSecondFilter = ""
    Me.FirstComboBox.Value = ""
    Me.FirstComboBox.RowSource = "SELECT DISTINCT Divisions.Title FROM Divisions " _
        & "WHERE Divisions.Code NOT IN ('06', '06R')"
End Sub


'----------------------------------------------------------------------------------
'   Type:        Property-Get
'   Name:        FiscalYear
'   Parameters:  Void
'   Purpose:
'----------------------------------------------------------------------------------
Public Property Get FiscalYear()
    If Not mBFY & "" = "" Then
        FiscalYear = mBFY
    End If
End Property


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FirstFilterComboBox_Change
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FirstComboBox_AfterUpdate()
    mFirstFilter = ""
    SetOptionGroupFilter
    mFirstFilter = "[DivisionName] = '" & Me.FirstComboBox.SelText & "'"
    mDivisionName = Me.FirstComboBox.SelText
    DialogArgs.DivisionName = mDivisionName
End Sub

'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExecuteButton_OnClick
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AcceptButton_Click()
    Me.Visible = Not Me.Visible
End Sub

