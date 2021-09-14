Option Compare Database
Option Explicit




Public Args As BudgetArgs
Private pAnd As String
Private mDatabase As DAO.Database
Private mData As DAO.Recordset
Private mQuery As DAO.QueryDef
Private mBudget As Report
Private mError As String




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Set Args = New BudgetArgs
    pAnd = " AND "
    Set mDatabase = CurrentDb
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     Form_Load()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Property-Get
'   Name:        SqlQuery
'   Parameters:  Void
'   Retval:      Void
'   Purpose:     Returns the sql query text used by the form
'---------------------------------------------------------------------------------
Public Property Get SqlQuery() As QueryDef
    On Error GoTo ErrorHandler:
    If IsObject(mQuery) Then
        Set SqlQuery = mQuery
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     SqlQuery()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearComboBoxValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.BfyComboBox.Value = ""
    Me.RcCodeComboBox.Value = ""
    Me.FundCodeComboBox.Value = ""
    Me.AhCodeComboBox.Value = ""
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     ClearComboBoxValues()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetButtonVisibility
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.ExitButton.SetFocus
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.AcceptButton.Visible = Not Me.AcceptButton.Visible
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     SetButtonVisibility()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        m_AccetpButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub m_AcceptButton_Click()
    On Error GoTo ErrorHandler:
    Dim SQL As String
    SQL = "SELECT Allocations.*, Appropriations.Title AS [Bill Title]," _
        & " Funds.Title AS [Fund Title], Funds.TreasurySymbol, Allocations.BFY, *" _
        & " FROM Funds INNER JOIN (Appropriations" _
        & " INNER JOIN Allocations ON Appropriations.BFY = Allocations.BFY)" _
        & " ON Funds.Code = Allocations.FundCode" _
        & " WHERE (((Allocations.BocCode)<>'10'" _
        & " AND (Allocations.BocCode)<>'17')" _
        & " AND ((Allocations.Amount)<>0) " _
        & " AND ((Allocations.AhCode)= '" & Args.AhCode & "'" _
        & " AND ((Appropriations.BFY)= '" & Args.BFY & "'" _
        & " AND ((Allocations.RcCode)= '" & Args.RcCode & "'" _
        & " AND ((Allocations.FundCode)= '" & Args.FundCode & "'" _
        & " AND ((Funds.Code)=[Allocations].[FundCode])" _
        & " AND ((Allocations.BudgetLevel)='8'));"
    Set mBudget = New Report_DivisionBudgetAuthority
    mBudget.RecordSource = SQL
    DoCmd.OpenReport mBudget.Name
    'DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSaveNo
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     m_AcceptButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AcceptButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:     Button even procedure that hides the dialog
'---------------------------------------------------------------------------------
Private Sub AcceptButton_Click()
    On Error GoTo ErrorHandler:
    Me.Visible = Not Me.Visible
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     AcceptButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ExitButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:     Closes the dialog
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, Me.Name, acSaveYes
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     ExitButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_AfterRender(ByVal drawObject As Object, ByVal chartObject As Object)
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     Form_AfterRender()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        AhCodeComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AhCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    If IsNull(Args) Then
        Set Args = New BudgetArgs
    End If
    If Not Me.AhCodeComboBox.SelText & "" = "" Then
        Args.AhCode = Me.AhCodeComboBox.SelText
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     AhCodeComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        BfyComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BfyComboBox_Change()
    On Error GoTo ErrorHandler:
    If Not Me.BfyComboBox.SelText & "" = "" Then
        Args.BFY = Me.BfyComboBox.SelText
        Me.RcCodeComboBox.RowSource = ""
        Me.RcCodeComboBox.RowSource = "SELECT DISTINCT RcCode" _
            & " FROM BudgetParameters" _
            & " WHERE RcCode NOT IN ('06', '06R')" _
            & " AND BudgetParameters.BFY = '" _
            & Args.BFY & "'" & _
            pAnd & "AhCode = '" & Args.AhCode & "'"
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     BfyComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        RcCodeComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RcCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    If Not Me.RcCodeComboBox.SelText & "" = "" Then
        Args.RcCode = Me.RcCodeComboBox.SelText
        Me.FundCodeComboBox.RowSource = "SELECT DISTINCT BudgetParameters.FundCode" & _
            " FROM BudgetParameters" & _
            " WHERE RcCode NOT IN ('06', '06R') AND BudgetParameters.BFY = '" & _
            Args.BFY & "'" & _
            pAnd & "RcCode = '" & _
            Args.RcCode & "'" & _
            pAnd & "AhCode = '" & _
            Args.AhCode & "'"
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     RcCodeComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FundCodeComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    If Not Me.FundCodeComboBox.SelText & "" = "" Then
        Args.FundCode = Me.FundCodeComboBox.SelText
        SetButtonVisibility
    End If
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     FundCodeComboBox_Change()" _
            & vbCrLf & "Descript:   " & Err.Description
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
'   Purpose:     Reset filters, control items, control visibility
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    Me.ExitButton.SetFocus
    ClearComboBoxValues
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:      BudgetDialog" _
            & vbCrLf & "Member:     RefreshButton_Click()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub
