VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Systems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_Error As String



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Caption = "System Menu"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        BfsButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BfsButton_Click()
    On Error GoTo ErrorHandler:
    Application.FollowHyperlink "https://bfs.epa.gov/Account/Login?ReturnUrl=/"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        BusinessObjectsButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BusinessObjectsButton_Click()
    On Error GoTo ErrorHandler:
    Application.FollowHyperlink "https://ocfoboprodw.epa.gov/BOE/BI"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        CompassButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CompassButton_Click()
    On Error GoTo ErrorHandler:
    Application.FollowHyperlink "https://compassmomentum.epa.gov/authservice/login.jsp;"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ConcurButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ConcurButton_Click()
    On Error GoTo ErrorHandler:
    Application.FollowHyperlink "https://idp.epa.gov/SecureAuth97/"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ConcurButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub EBusinessButton_Click()
    On Error GoTo ErrorHandler:
    Application.FollowHyperlink "https://ebusiness.epa.gov/ebusiness/index.cfm?event=home"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ConcurButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub EpaAcquisitionSystemButton_Click()
    On Error GoTo ErrorHandler:
    Application.FollowHyperlink "https://eas.epa.gov/prism.net/Security/logon.aspx"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        PeoplePlusButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PeoplePlusButton_Click()
    On Error GoTo ErrorHandler:
    Application.FollowHyperlink "https://peopleplus.epa.gov/psp/ots92prd/"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ProcurementDataSystemButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProcurementDataSystemButton_Click()
    On Error GoTo ErrorHandler:
    Application.FollowHyperlink "https://www.fpds.gov/fpdsng_cms/index.php/en/"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ScorpiosButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ScorpiosButton_Click()
    On Error GoTo ErrorHandler:
    Application.FollowHyperlink "https://usepa.sharepoint.com/sites/OCFO/ots/Pages/SCORPIOS%20e-Recovery.aspx"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub







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




