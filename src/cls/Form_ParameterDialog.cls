VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ParameterDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Args As AccountArgs
Private pAnd As String
Private m_Level As String
Private m_FiscalYear As String
Private m_FundCode As String
Private m_AhCode As String
Private m_OrgCode As String
Private m_BocCode As String
Private m_AccountCode As String
Private m_ActivityCode As String
Private m_RcCode As String
Private m_And As String
Private m_Filter As String
Private m_Error As String










'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Set Args = New AccountArgs
    m_And = " AND "
    Me.AhCodeComboBox.RowSource = "SELECT DISTINCT Allocations.AhCode FROM Allocations;"
    Me.FiscalYearComboBox.RowSource = "SELECT DISTINCT Allocations.BFY FROM Allocations;"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me!AhCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!AhCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!AhCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!FiscalYearComboBox.ForeColor = RGB(255, 255, 255)
    Me!FiscalYearComboBox.BackColor = RGB(33, 33, 33)
    Me!FiscalYearComboBox.BorderColor = RGB(68, 114, 196)
    Me!FundCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!FundCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!FundCodeComboBox.BorderColor = RGB(68, 114, 196)
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
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    SetComboBoxColors
    Me.FiscalYearComboBox.Value = ""
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
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AhCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.AhCode = Me!AhCodeComboBox.SelText
    AH = Args.AhCode
    m_Filter = "Allocations.AhCode = '" & AH & "'"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.BFY = Me!FiscalYearComboBox.SelText
    m_Filter = m_Filter & m_And & "Allocations.BFY = '" & Args.BFY & "'"
    Me.FundCodeComboBox.RowSource = "SELECT DISTINCT Allocations.FundCode FROM Allocations WHERE " & m_Filter
    MsgBox m_Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.FundCode = Me!FundCodeComboBox.SelText
    m_Filter = m_Filter & m_And & "Allocations.FundCode = '" & Args.FundCode & "'"
    MsgBox m_Filter
    Me.OrgCodeComboBox.RowSource = "SELECT DISTINCT Allocations.OrgCode FROM Allocations WHERE " & m_Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub OrgCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.OrgCode = Me!OrgCodeComboBox.SelText
    m_Filter = m_Filter & m_And & "Allocations.OrgCode = '" & Args.OrgCode & "'"
    MsgBox m_Filter
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT Allocations.AccountCode FROM Allocations WHERE " & m_Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ActivityCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.ActivityCode = Me!ActivityCodeComboBox.SelText
    Activity = Args.ActivityCode
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RcCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.RcCode = Me!RcCodeComboBox.SelText
    RC = Args.RcCode
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.BocCode = Me!BocCodeComboBox.SelText
    BOC = Args.BocCode
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    m_Filter = ""
    ClearComboBoxValues
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Level(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        m_Level = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Level() As String
    On Error GoTo ErrorHandler:
    If (m_Level <> "" And Not IsNull(m_Level)) Then
        Level = m_Level
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property






'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let FiscalYear(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        m_FiscalYear = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FiscalYear() As String
    On Error GoTo ErrorHandler:
    If (m_FiscalYear <> "" And Not IsNull(m_FiscalYear)) Then
        FiscalYear = m_FiscalYear
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property






'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Fund(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        m_FundCode = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Fund() As String
    On Error GoTo ErrorHandler:
    If (m_FundCode <> "" And Not IsNull(m_FundCode)) Then
        Fund = m_FundCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property






'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let BOC(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        m_BocCode = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BOC() As String
    On Error GoTo ErrorHandler:
    If (m_BocCode <> "" And Not IsNull(m_BocCode)) Then
        BOC = m_BocCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let RC(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        m_RcCode = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get RC() As String
    On Error GoTo ErrorHandler:
    If (m_RcCode <> "" And Not IsNull(m_RcCode)) Then
        RC = m_RcCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Account() As String
    On Error GoTo ErrorHandler:
    If Not m_AccountCode & "" = "" Then
        Account = m_AccountCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property






'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Account(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        m_AccountCode = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let AH(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        m_AhCode = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get AH() As String
    On Error GoTo ErrorHandler:
    If (m_AhCode <> "" And Not IsNull(m_AhCode)) Then
        AH = m_AhCode
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
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
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Activity(ac As String)
    On Error GoTo ErrorHandler:
    If Not ac & "" = "" Then
        m_ActivityCode = ac
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ORG(Code As String)
    On Error GoTo ErrorHandler:
    If (Code <> "") Then
        m_OrgCode = Code
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ORG() As String
    On Error GoTo ErrorHandler:
    If (m_OrgCode <> "" And Not IsNull(m_OrgCode)) Then
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





