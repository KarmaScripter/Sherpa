VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_EmployeeDataByDivision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public mBudgetArgs As BudgetArgs
Public m_Dialog As Form_RegionDialog
Public RcCodeArgs As String
Public DivisionNameFilter As String
Private m_Error As String







'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set mBudgetArgs = New BudgetArgs
    DoCmd.OpenForm FormName:="RegionDialog", view:=acNormal, WindowMode:=acDialog
    Set m_Dialog = Forms("RegionDialog")
     m_Dialog.FundNameComboBoxLabel.Visible = False
     m_Dialog.FundNameComboBox.Visible = False
     m_Dialog.FundNameGroupBox.Visible = False
     m_Dialog.FiscalYearOptionGroup.Visible = False
     m_Dialog.FiscalYearOptionGroupLabel.Visible = False
     m_Dialog.FirstRadioButton.Visible = False
     m_Dialog.FirstRadioButtonLabel.Visible = False
     m_Dialog.SecondRadioButton.Visible = False
     m_Dialog.SecondRadioButtonLabel.Visible = False
    Set mBudgetArgs = m_Dialog.Args
    RcCodeArgs = mBudgetArgs.RcCode
    Dim SQL As String
    SQL = "SELECT DivisionPersonnel.*," _
        & "  EmployeeData.*, * FROM EmployeeData" _
        & " INNER JOIN DivisionPersonnel ON EmployeeData.EmployeeId = DivisionPersonnel.EmployeeId" _
        & " WHERE DivisionPersonnel.RcCode = '" & mBudgetArgs.RcCode & "';"
    Me.RecordSource = SQL
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

