VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AccountDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Args As AccountArgs
Private pAnd As String
Private m_BudgetLevel As String
Private m_FiscalYear As String
Private m_FundCode As String
Private m_AhCode As String
Private m_OrgCode As String
Private m_BocCode As String
Private m_AccountCode As String
Private m_ActivityCode As String
Private m_RcCode As String
Private m_Error As String




'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Set Args = New AccountArgs
    ClearComboBoxValues
    SetButtonVisibility
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
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
    Me!ActivityCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!ActivityCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!ActivityCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!RcCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!RcCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!RcCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!BocCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!BocCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!BocCodeComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    SetComboBoxColors
    Me.BfyComboBox.Value = ""
    Me.AccountCodeComboBox.Value = ""
    Me.BocCodeComboBox.Value = ""
    Me.RcCodeComboBox.Value = ""
    Me.AhCodeComboBox.Value = ""
    Me.ActivityCodeComboBox.Value = ""
    Me.FundCodeComboBox.Value = ""
    Me.OrgCodeComboBox.Value = ""
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.AcceptButton.Visible = Not Me.AcceptButton.Visible
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub SetGoButtonVisibility()
    On Error GoTo ErrorHandler:
    If IsNull(Args.BFY) And _
        IsNull(Args.FundCode) And _
        Me!GoButton.Visible = False Then
            Me!GoButton.Visible = True
        Else: Me!GoButton.Visible = False
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ActivityCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.ActivityCode = Me!ActivityCodeComboBox.SelText
    m_ActivityCode = Args.ActivityCode
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AhCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.AhCode = Me!AhCodeComboBox.SelText
    AH = Args.AhCode
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub OrgCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.OrgCode = Me!OrgCodeComboBox.SelText
    ORG = Args.OrgCode
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RcCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.RcCode = Me!RcCodeComboBox.SelText
    RC = Args.RcCode
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BfyComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.BFY = Me!BfyComboBox.SelText
    FiscalYear = Args.BFY
    Me!AcceptButton.Visible = Not Me!AcceptButton.Visible
    Me!RefreshButton.Visible = Not Me!RefreshButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.BocCode = Me!BocCodeComboBox.SelText
    BOC = Args.BocCode
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.FundCode = Me!FundCodeComboBox.SelText
    m_FundCode = Args.FundCode
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    Me!ExitButton.SetFocus
    ClearComboBoxValues
    Me!GoButton.Visible = Not Me!GoButton.Visible
    Me!RefreshButton.Visible = Not Me!RefreshButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Level(bl As String)
    On Error GoTo ErrorHandler:
    If (bl & "") = "" Then
        m_BudgetLevel = bl
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Level() As String
    On Error GoTo ErrorHandler:
    If Not (m_BudgetLevel & "") = "" Then
        m_BudgetLevel = m_BudgetLevel
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FiscalYear(pCode As String)
    On Error GoTo ErrorHandler:
    If Not pCode & "" = "" Then
        m_FiscalYear = pCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FiscalYear() As String
    On Error GoTo ErrorHandler:
    If Not m_FiscalYear & "" = "" Then
        FiscalYear = m_FiscalYear
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property






'---------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Fund(pCode As String)
    On Error GoTo ErrorHandler:
    If Not pCode & "" = "" Then
        m_FundCode = pCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Fund() As String
    On Error GoTo ErrorHandler:
    If Not m_FundCode & "" = "" Then
        Fund = m_FundCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BOC(pBOC As String)
    On Error GoTo ErrorHandler:
    If Not (pBOC & "") = "" Then
        m_BocCode = pBOC
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BOC() As String
    On Error GoTo ErrorHandler:
    If Not m_BocCode & "" = "" Then
        BOC = m_BocCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let RC(pRC As String)
    On Error GoTo ErrorHandler:
    If Not (pRC = "") = "" Then
        m_RcCode = pRC
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RC() As String
    On Error GoTo ErrorHandler:
    If Not m_RcCode & "" = "" Then
        RC = m_RcCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Code(pCode As String)
    On Error GoTo ErrorHandler:
    If Not pCode & "" = "" Then
        m_AccountCode = pCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Code() As String
    On Error GoTo ErrorHandler:
    If Not m_AccountCode & "" = "" Then
        Code = m_AccountCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let AH(pAH As String)
    On Error GoTo ErrorHandler:
    If Not (pAH & "" = "") Then
        m_AhCode = pAH
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get AH() As String
    On Error GoTo ErrorHandler:
    If Not m_AhCode & "" = "" Then
        AH = m_AhCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Activity() As String
    On Error GoTo ErrorHandler:
    If Not m_ActivityCode & "" = "" Then
        Activity = m_ActivityCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Activity(pActivity As String)
    On Error GoTo ErrorHandler:
    If Not pActivity & "" = "" Then
        m_ActivityCode = pActivity
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ORG(pORG As String)
    On Error GoTo ErrorHandler:
    If Not pORG & "" = "" Then
        m_OrgCode = pORG
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ORG() As String
    On Error GoTo ErrorHandler:
    If Not m_OrgCode & "" = "" Then
        ORG = m_OrgCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProcessError
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProcessError()
    If Err.Number <> 0 Then
        m_Error = "Source:      " & Err.Source _
            & vbCrLf & "Number:     " & Err.Number _
            & vbCrLf & "Issue:      " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
End Sub





