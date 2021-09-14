VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_BudgetCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Dim mActiveControl As control ' Open
Dim mPickDate As Date ' Open
Dim gBooTime As Boolean ' Load
Public SelectedDate As Date
Private mError As String




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:     FormName -- .txtDate = pDate, .ProcedureName CDate(pDate)
'                CALLED BY DayClick, buttons to change calendar day,
'                and from code behind other forms (FormName)
'                runs ProcedureName in code behind FormName
'                me.txtCalendarDate holds the calendar date
'                me.txtHr, me.txtMin, Me.txtAP --> time
'---------------------------------------------------------------------------------
Public Sub Update_ExternalForms(pDate As Variant)
   
'    On Error GoTo Proc_Err
    
'    If CurrentProject.AllForms(FormName).IsLoaded Then
'        With Forms!FormName
'            .txtDate = pDate
'            DoEvents
'            .ProcedureName CDate(pDate)
'        End With
'    End If
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   Update_ExternalForms : " & Me.Name

   Resume Proc_Exit
   Resume
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_Cancel_Click()
   On Error Resume Next
   DoCmd.Close acForm, Me.Name, acSaveNo
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_Close_Click()
   On Error Resume Next

   Dim nDate As Date

   nDate = DateValue(Me.txtCalendarDate)
   
   'add time to date if time controls are showing
   If Me.chkUseTime Then
'      If CInt(Nz(Me.txtHr, "0")) > 0 Or CInt(Nz(Me.txtMin, "0")) > 0 Then
         nDate = nDate _
            + TimeSerial(CInt(Nz(Me.txtHr, "0")) + IIf(InStr(Me.txtAP, "p") > 0, 12, 0) _
            , CInt(Nz(Me.txtMin, "0")), 0)
'      End If
   End If
   
   If Not Len(Nz(Me.OpenArgs, "")) > 0 Then
      'will throw an error if mActiveControl is not defined
      '   ie: maybe there was no active form when the date picker was launched
      '       or there was an unhandled error and the object variable was lost
      'mActiveControl = nDate
      SelectedDate = nDate
      'If mActiveControl <> nDate Then
         'form was opened independently
         'tell user how to get this feature into another database
         'ShowDatePickerMessage
    End If
    Me.Visible = False
   'DoCmd.Close acForm, Me.Name, acSaveNo
   
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function UseTheTime(pBoo As Boolean)
   On Error Resume Next
   Me.chkUseTime = pBoo
   Me.Label_chkUseTime.FontBold = pBoo
   
   If Me.txtAP.Visible <> pBoo Then
      cal_ShowHideControlsTag pBoo, "Time"
   End If
   
End Function



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_CurrentTime_Click()
   On Error Resume Next
   
   UseTheTime True
   Me.txtHr.Value = Hour(Now()) Mod 12
   Me.txtMin.Value = Minute(Now())
   Me.txtAP = IIf(DatePart("h", Now()) >= 12, "pm", "am")
   Me.txtHr.SetFocus
End Sub



'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdMonth_Click()
   On Error Resume Next
   MsgBox cal_GetBirthstone(Month(Me.txtCalendarDate)), , "Birthstone for " & Me.cmdMonth.Caption
   
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdYr_Click()
   On Error Resume Next
   'year in roman numbers
   'get Chinese zodiac animal?
   Dim nYear As Integer
   nYear = CInt(Me.cmdYr.Caption)
   MsgBox cal_GetRoman(nYear), , nYear & " in Roman Numbers"
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
   On Error Resume Next
   Dim sStr As String _
      , booTimeVisible As Boolean
   
   booTimeVisible = False
   mPickDate = Date
   
   If Len(Me.OpenArgs) > 0 Then
      sStr = Me.OpenArgs
      If IsDate(sStr) Then
         mPickDate = CDate(sStr)
      End If
   Else
      sStr = Trim(Screen.ActiveForm.Caption & "")
      If Err.Number > 0 Then GoTo Proc_Exit
      If Len(sStr) = 0 Then
         sStr = Screen.ActiveForm.Name
      End If
      If Not Err.Number > 0 Then
         Set mActiveControl = Screen.ActiveControl
         
         Select Case True

         Case IsDate(mActiveControl)
            If Not mActiveControl = 0 Then
               mPickDate = mActiveControl.Value
            End If
            
         End Select
   
         With mActiveControl
            sStr = .Controls(0).Caption
            If Err.Number > 0 Then
               sStr = mActiveControl.Name
            End If
         End With
         sStr = Trim(sStr)
         If Right(sStr, 1) = ":" Then sStr = Left(sStr, Len(sStr) - 1)
         
      End If

   End If
   
Proc_Exit:
   Exit Sub
   
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
   
' CALLS
   ' cal_cal_GetRow4Calendar
   ' cal_cal_GetCol4Calendar
   ' Set_Calendar
   ' cal_ShowHideControlsTag
   
    On Error GoTo Proc_Err
    
    Dim nRow As Integer _
      , nCol As Integer _
         , iPos As Integer _
         , nDate As Date _
         , sStr As String
      
   nDate = mPickDate 'set in Open event
   
   'openArgs
   
   nRow = cal_GetRow4Calendar(nDate)
   nCol = cal_GetCol4Calendar(nDate)
   
   'keep track so colors can be set back to normal
   
   Me.txtRowPick = nRow
   Me.txtColPick = nCol
   Me.txtRowCur = nRow
   Me.txtColCur = nCol
   Me.txtCalendarDate = nDate
   
   'see if date includes a time component
   gBooTime = IIf(nDate <> DateValue(nDate), True, False)
   
   Me.chkUseTime = gBooTime
   
   If gBooTime Then
      Me.txtMin = Minute(nDate)
      If Hour(nDate) > 12 Then
         Me.txtHr = Hour(nDate) - 12
         Me.txtAP = "pm"
      Else
         Me.txtHr = Hour(nDate)
         Me.txtAP = "am"
      End If
   End If
   
   cal_ShowHideControlsTag gBooTime, "Time"

   Set_Calendar nDate
       
   'draw the circle for Reset
   
Proc_Exit:
   On Error Resume Next
   Exit Sub
  
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   Form_Load : " & Me.Name

   Resume Proc_Exit
   Resume
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub DayClick()
' CALLS
   ' Set_Calendar
   ' Update_ExternalForms
   
    On Error GoTo Proc_Err
    If Me.ActiveControl.Caption = "" Then
        'user clicked on a day with no number - do nothing
        Exit Sub
    End If
    
    Dim nRow As Integer _
        , nCol As Integer
        
    Dim nDate As Date _
        , nOldDate As Date _
        , nDay As Integer
        
    nDay = Me.ActiveControl.Caption
    
    nOldDate = Me.txtCalendarDate
    
    nDate = DateSerial(year(nOldDate), Month(nOldDate), nDay)
    
    Set_Calendar nDate
    Update_ExternalForms nDate
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
  
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   DayClick : " & Me.Name

   Resume Proc_Exit
   Resume
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_Reset_Click()
'reset date back to original pick
   On Error Resume Next
   
   Dim nDate As Date
   nDate = mPickDate
   
   Me.txtCalendarDate = nDate
   
   UseTheTime gBooTime
    
   cal_ShowHideControlsTag gBooTime, "Time"
   
   If gBooTime Then
      Me.txtMin = Minute(nDate)
      If Hour(nDate) > 12 Then
         Me.txtHr = Hour(nDate) - 12
         Me.txtAP = "am"
      Else
         Me.txtHr = Hour(nDate)
         Me.txtAP = "pm"
      End If
   End If
   
   Add_SetCalendar nDate, 0, 1, 0
   Update_ExternalForms nDate
   
   
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_AddDays_Click()
   Dim nDays As Long
   If IsNull(Me.txtDays) Then
      Me.txtDays.SetFocus
      MsgBox "Specify number of days to add or subtract", , "Can't add days, no number specified"
      Exit Sub
   End If
   nDays = Me.txtDays
   If nDays = 0 Then
      Me.txtDays.SetFocus
      MsgBox "Specify number of days to add or subtract", , "Can't add days, no number specified"
      Exit Sub
   End If
   Dim nDate As Date
   nDate = Me.txtCalendarDate
   nDate = DateSerial(year(nDate), Month(nDate), Day(nDate) + nDays)
   Set_Calendar nDate
   Update_ExternalForms nDate
End Sub






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_Now_Click()
   On Error Resume Next
   
   Dim nDate As Date
   nDate = Date
   
   Me.txtCalendarDate = nDate
   
   Set_Calendar nDate
   Update_ExternalForms nDate
   
   UseTheTime True
   cmd_CurrentTime_Click
   
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdMonthAdd_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
   
    On Error GoTo Proc_Err
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 0, 1, 0
    Update_ExternalForms nDate
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdMonthSub_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
   On Error Resume Next
   
   Dim nDate As Date
   nDate = Me.txtCalendarDate
   Add_SetCalendar nDate, 0, -1, 0
   Update_ExternalForms nDate

End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdYrAdd_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
 
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 1, 0, 0
    Update_ExternalForms nDate

Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdYrSub_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
 
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, -1, 0, 0
    Update_ExternalForms nDate
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_M6add_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
   
    On Error GoTo Proc_Err
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 0, 6, 0
    Update_ExternalForms nDate
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_M6sub_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
   
    On Error GoTo Proc_Err
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 0, -6, 0
    Update_ExternalForms nDate
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_Today_Click()
' CALLS
   ' Set_Calendar
   ' Update_ExternalForms
   
    On Error GoTo Proc_Err
    Set_Calendar Date
    Update_ExternalForms Date
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_W1add_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
   
    On Error GoTo Proc_Err
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 0, 0, 7
    Update_ExternalForms nDate
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_W1sub_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
   
    On Error GoTo Proc_Err
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 0, 0, -7
    Update_ExternalForms nDate
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_Q1add_Click()
    On Error GoTo Proc_Err
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 0, 3, 0
    Update_ExternalForms nDate
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_Q1sub_Click()
    On Error GoTo Proc_Err
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 0, -3, 0
    Update_ExternalForms nDate
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_Y10add_Click()
    On Error GoTo Proc_Err
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 10, 0, 0
    Update_ExternalForms nDate
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd_Y10sub_Click()
    On Error GoTo Proc_Err
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, -10, 0, 0
    Update_ExternalForms nDate
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub txtCalendarDate_AfterUpdate()
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Set_Calendar nDate
    Update_ExternalForms nDate
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdDayAdd_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
 
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 0, 0, 1
    Update_ExternalForms nDate
    
Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmdDaySub_Click()
' CALLS
   ' Add_SetCalendar
   ' Update_ExternalForms
 
    Dim nDate As Date
    nDate = Me.txtCalendarDate
    Add_SetCalendar nDate, 0, 0, -1
    Update_ExternalForms nDate

Proc_Exit:
   On Error Resume Next
   Exit Sub
Proc_Err:
   Resume Proc_Exit
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd11_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd12_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd13_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd14_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd15_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd16_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd17_Click()
    DayClick
End Sub

Private Sub cmd21_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd22_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd23_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd24_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd25_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd26_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd27_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd31_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd32_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd33_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd34_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd35_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd36_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd37_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd41_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd42_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd43_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd44_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd45_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd46_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd47_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd51_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd52_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd53_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd54_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd55_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd56_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd57_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd61_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd62_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd63_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd64_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd65_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd66_Click()
    DayClick
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub cmd67_Click()
    DayClick
End Sub






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub hDn_Click()
   UseTheTime True
   If Nz(Me.txtHr.Value, 1) = 1 Then
       Me.txtHr.Value = 12
     Else
       Me.txtHr.Value = Me.txtHr.Value - 1
   End If
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub hUp_Click()
   UseTheTime True
    If Nz(Me.txtHr.Value, 12) = 12 Then
        Me.txtHr.Value = 1
      Else
        Me.txtHr.Value = Me.txtHr.Value + 1
    End If
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function HrUpDn(pStep As Integer)
   UseTheTime True
   
   Dim nHr As Integer _
      , nHrOld As Integer
      
   nHrOld = Nz(Me.txtHr, 12)
     
   nHr = (Nz(Me.txtHr.Value, 0) + pStep) Mod 12
   If nHr = 0 Then nHr = 12
   
   Select Case True
   Case nHr = 12 And pStep > 0
      AmPm
   Case nHrOld = 12 And pStep > 0
   Case nHr < 1
      nHr = 12 + nHr
      If nHr <> 12 And pStep <> -1 Then AmPm
   
   Case nHr >= 13
      nHr = nHr - 12
      If nHrOld <> 12 Then AmPm
   
   Case nHr = 11 And pStep = -1
       AmPm
   Case nHrOld = 12 And pStep < 0
      AmPm
   Case nHrOld = 12, nHr = 12
   Case pStep < 0
      If nHr > nHrOld Then AmPm
   Case pStep > 0
      If nHr < nHrOld Then AmPm
   End Select

Proc_Exit:
   Me.txtHr.Value = nHr

End Function





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function MinUpDn(pStep As Integer)
   UseTheTime True
   Dim nMin As Integer
   
   'move up or down on even increments of 30
   nMin = ((Nz(Me.txtMin.Value, 0) + 1) \ 30) * 30
   
   nMin = nMin + pStep
   
   If nMin < 0 Then
      nMin = 60 + nMin
      HrUpDn -1
   End If
   If nMin >= 60 Then
      nMin = nMin - 60
      HrUpDn 1
   End If
   
   Me.txtMin.Value = nMin
End Function





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function AmPm()
    If Me.txtAP.Value = "am" Then
        Me.txtAP.Value = "pm"
    Else
        Me.txtAP.Value = "am"
    End If
End Function






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Add_SetCalendar( _
      pDate As Date _
    , Optional pYearAdd As Integer = 0 _
    , Optional pMonthAdd As Integer = 0 _
    , Optional pDayAdd As Integer = 0 _
    )

   On Error GoTo Proc_Err

    If pMonthAdd <> 0 Or pYearAdd <> 0 Or pDayAdd <> 0 Then
      pDate = DateSerial(year(pDate) + pYearAdd, Month(pDate) + pMonthAdd, Day(pDate) + pDayAdd)
    End If
    
    Set_Calendar pDate
    
Proc_Exit:
    On Error Resume Next
   Exit Sub
  
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   Add_SetCalendar"

   Resume Proc_Exit
   Resume
End Sub






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Set_Calendar( _
      pDate As Date _
    )
'---------- CUSTOMIZE Defaults for -- Set_Calendar
    

    
'PARAMETERS
   'pDate - optional. if specified and > 1900, calendar will be set to the date
   
' CALLS
'  cal_GetRow4Calendar
'  cal_GetCol4Calendar
'  Set_DefaultFormat
'  Mark_TodayAndDate
   
' CALLED BY
   ' Form Load and buttons to change calendar day
   ' FormName: txtDate_AfterUpdate, FindTheDay, DayAddSub
    
   On Error GoTo Proc_Err
    
    '----- dimension variables
    Dim ctl As control _
      , DB As DAO.Database _
      , rs As DAO.Recordset
        
    Dim nMonth As Integer _
        , nYear As Integer _
        , nFirstCol As Integer _
        , nLastRow As Integer _
        , nLastCol As Integer _
        , iRow As Integer _
        , iCol As Integer _
        , nRowPick As Integer _
        , nColPick As Integer _
        , nRowCur As Integer _
        , nColCur As Integer _
        , sSQL As String _
        , sStr As String _
        , iDay As Integer

    '----- set variables
    
   nMonth = Month(pDate)
   nYear = year(pDate)
   
   If year(Date) = nYear And Month(Date) = nMonth Then
      'calendar is showing the current month
      nRowCur = cal_GetRow4Calendar(Date)
      nColCur = cal_GetCol4Calendar(Date)
   Else
      'calendar is not showing the current month
      nRowCur = 0
      nColCur = 0
   End If
   
   nRowPick = cal_GetRow4Calendar(pDate)
   nColPick = cal_GetCol4Calendar(pDate)

   nLastCol = Weekday(DateSerial(nYear, nMonth + 1, 0))
   nLastRow = cal_GetRow4Calendar(DateSerial(nYear, nMonth + 1, 0))
   nFirstCol = Weekday(DateSerial(nYear, nMonth, 1))
      
   'keep track of picked day so colors can be set back to normal
   'when the date is changed
   
   If Me.txtRowPick <> nRowPick Then
      Set_DefaultFormat Me("cmd" & Me.txtRowPick & Me.txtColPick), , False
      Me.txtRowPick = nRowPick
      Me.txtColPick = nColPick
   End If
   
   If Me.txtRowCur <> nRowCur Then
      'reset previous current date if is was showing
      If Me.txtRowCur <> 0 And Me.txtColCur <> 0 Then
         Set_DefaultFormat Me("cmd" & Me.txtRowCur & Me.txtColCur), , False
      End If
      Me.txtRowCur = nRowCur
      Me.txtColCur = nColCur
   End If
   
   Me.txtCalendarDate = pDate
   Me.txtCalendarDate.Tag = "cmd" & nRowPick & nColPick
   
    If nLastRow = 0 Or nLastCol = 0 Then
        MsgBox "Error getting last row or column for calendar", , "Aborting"
        Exit Sub
    End If
    
    'caption for cmdMonth
    Me.cmdMonth.Caption = Format(pDate, "mmmm")
    Me.cmdYr.Caption = nYear
        
    'hide unused squares in the first row
    For iCol = 1 To (nFirstCol - 1)
      Set ctl = Me("cmd1" & iCol)
      With ctl
         .Visible = False
      End With
    Next iCol
    
   '-----------------------------------------------------------------------
   ' reset visible cells to default format
   
   iDay = 1
   
   iRow = 1
   iCol = 1
   
   For iRow = 1 To 6
      For iCol = 1 To 7
   
         Set ctl = Me("cmd" & iRow & iCol)
      
         Select Case iRow
         Case 1
            If iCol < nFirstCol Then
               ctl.Visible = False
               GoTo NextDay
            Else
               Set_DefaultFormat ctl, iDay, iCol, False
               iDay = iDay + 1
            End If
            
         Case nLastRow
            If iCol <= nLastCol Then
               Set_DefaultFormat ctl, iDay, iCol, False
               iDay = iDay + 1
            Else
               ctl.Visible = False
               GoTo NextDay
            End If
            
         Case Is < nLastRow
               Set_DefaultFormat ctl, iDay, iCol, False
               iDay = iDay + 1
               
         Case Is > nLastRow
            ctl.Visible = False
            GoTo NextDay
               
         End Select
NextDay:
      Next iCol
   Next iRow
      
   Call Mark_TodayAndDate(pDate)
 
Proc_Exit:
    On Error Resume Next
    Set ctl = Nothing
   Exit Sub
  
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   Set_Calendar"

   Resume Proc_Exit
   Resume
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Set_DefaultFormat(pCtl As control _
   , Optional iDay As Integer = 0 _
   , Optional iCol As Integer = 0 _
   , Optional BoldWkend As Boolean = True)
'Private Sub Set_DefaultFormat(pCtl As Control, Optional iDay As Integer = 0)
'120623, 120627, 120701
   Dim booBold As Boolean
   
   With pCtl
      .Visible = True
      .FontSize = 10
      .ForeColor = 0 'black
      booBold = True
      If iDay > 0 Then
         .Caption = iDay & Chr(160) & Chr(160) & Chr(160) & Chr(160) & vbCrLf & Chr(160)
         If Not BoldWkend _
               And (iCol = 1 Or iCol = 7) Then
            booBold = False
         End If
      End If
      .BorderColor = Me.Detail.BackColor
      .BorderWidth = 2
      .FontBold = booBold
   End With
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Mark_TodayAndDate(pDate As Date)
'120623, 120627, 120701
   Dim nRow As Integer _
      , nCol As Integer _
      , sCtlName As String
      
   'set format back for last current date
   
   ' clear current date
   If Not Format(pDate, "yyyymm") = Format(Date, "yyyymm") Then
      'calendar is not showing current month
      If Nz(Me.txtRowCur, 0) <> 0 And Nz(Me.txtColCur, 0) <> 0 Then
         sCtlName = "cmd" & Me.txtRowCur & Me.txtColCur
         Set_DefaultFormat Me(sCtlName)
      End If
      GoTo MarkScheduleDate
   Else
      nRow = cal_GetRow4Calendar(Date)
      nCol = cal_GetCol4Calendar(Date)
      sCtlName = "cmd" & nRow & nCol
      With Me(sCtlName)
         .ForeColor = RGB(255, 0, 0) 'red
         .BorderWidth = 0 'hairline
         .BorderColor = RGB(255, 0, 0)
      End With
   End If
   
'   ' clear pick date date
'   If Not Format(pDate, "yyyymm") = Format(Me.txtCalendarDate, "yyyymm") Then
'      'pick date is different
'      If Nz(Me.txtRowPick, 0) <> 0 And Nz(Me.txtColPick, 0) <> 0 Then
'         sCtlName = "cmd" & Me.txtRowPick & Me.txtColPick
'         Set_DefaultFormat Me(sCtlName)
'      End If
'   End If
   
   
   If pDate = Date Then
      'make control purple if Pick = Today
      With Me(sCtlName)
         .ForeColor = RGB(150, 0, 250) 'purple
         .BorderWidth = 0 'hairline
         .BorderColor = RGB(150, 0, 250)
      End With
      GoTo Proc_Exit
   End If

MarkScheduleDate:
   'mark schedule date
   nRow = cal_GetRow4Calendar(pDate)
   nCol = cal_GetCol4Calendar(pDate)
   sCtlName = "cmd" & nRow & nCol
   
   With Me(sCtlName)
      .ForeColor = RGB(0, 0, 255) 'blue
         .BorderColor = RGB(0, 0, 255)
         .BorderWidth = 0 'hairline
'      If IsSubform(Me) Then '120623
'         Me.Parent.Label_DayDesc.Caption = .ControlTipText
'      End If
   End With
   
   
Proc_Exit:
   On Error Resume Next
   Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ShowDatePickerMessage()
'120701
   MsgBox "To use this popup calendar in a form," _
      & " assign the DOUBLE-CLICK event " _
      & " of date control on a form to" & vbCrLf & vbCrLf _
      & "   DoCmd.OpenForm ""f_PopupCalendar""" & vbCrLf & vbCrLf _
      & vbCrLf & vbCrLf _
      & "To use this in another database, " _
      & "import form f_PopupCalendar" _
      , , "Popup Calendar by Crystal"

End Sub






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Function cal_GetCardinalNumber(Optional pNumber) As String
'11-24-08
   'written by fdcusa (John)
   'modified by Crystal

   'returns the string from a number in this form:
   '1st, 2nd, 3rd, 10th, 301st, 1000th

   If IsMissing(pNumber) Or IsNull(pNumber) Or (Not IsNumeric(pNumber)) Then Exit Function
   
   Dim strEnding As String
   
   'convert to string, get the last character
   'then turn back into an integer for comparison
   
    Select Case CInt(Right(CStr(pNumber), 1))
        Case 1: strEnding = "st"
        Case 2: strEnding = "nd"
        Case 3: strEnding = "rd"
        Case Else: strEnding = "th"
    End Select
 
    cal_GetCardinalNumber = CStr(pNumber) & strEnding
    
End Function





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function cal_GetCol4Calendar(pDate As Date) As Integer
   cal_GetCol4Calendar = 0
   cal_GetCol4Calendar = Weekday(pDate, vbSunday)
End Function





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function cal_GetRow4Calendar(pDate As Date) As Integer
'120623 Crystal

   On Error GoTo Proc_Err
   cal_GetRow4Calendar = 0
        
   Dim nCol_First As Integer _
      , nDate_First As Date _
      , nRow As Integer _
      , nCol As Integer _
      , nNumDaysRow1 As Integer
   
   nDate_First = DateSerial(year(pDate), Month(pDate), 1)
   nCol_First = Weekday(nDate_First, vbSunday)
   nNumDaysRow1 = 7 - nCol_First + 1
   
   nCol = Weekday(pDate, vbSunday)
   
   nRow = (Day(pDate)) \ 7 + 1
   
   If Day(pDate) Mod 7 > nNumDaysRow1 Then nRow = nRow + 1
   If Day(pDate) Mod 7 = 0 And nCol >= nCol_First Then nRow = nRow - 1
   
   cal_GetRow4Calendar = nRow
   
Proc_Exit:
   On Error Resume Next
   Exit Function
  
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   cal_GetRow4Calendar"

   Resume Proc_Exit
   Resume
End Function





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function cal_GetDowN4Calendar(pDate As Date) As Integer
'120623 Crystal

   On Error GoTo Proc_Err
   cal_GetDowN4Calendar = 0
        
   Dim nDowN As Integer
         
   nDowN = (Day(pDate)) \ 7 + 1
   If Day(pDate) Mod 7 = 0 Then nDowN = nDowN - 1
      
   cal_GetDowN4Calendar = nDowN
   
Proc_Exit:
   On Error Resume Next
   Exit Function
  
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   cal_GetDowN4Calendar"

   Resume Proc_Exit
   Resume
End Function







'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function cal_IsSubform(pForm As Form) As Boolean
'8-29-07
   'return:
   ' TRUE is specified form reference is being used as a subform
   ' FALSE if it is not
   
   'example useage: in code before parent controls are used
   'If IsSubform(Me) then ...
   
    On Error Resume Next
    cal_IsSubform = _
       Not IsError(Len(pForm.Parent.Name) > 0)
       
End Function






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function cal_ShowHideControlsTag( _
   pBoo As Boolean _
   , pTag As String)
 
 
   On Error GoTo Proc_Err
   
   Dim ctl As control
      
   On Error Resume Next
   For Each ctl In Me.Detail.Controls
      If InStr(ctl.Tag, pTag) > 0 Then
         ctl.Visible = pBoo
      End If
   Next ctl
   
Proc_Exit:
   If Not ctl Is Nothing Then Set ctl = Nothing
   Exit Function
  
Proc_Err:
   MsgBox Err.Description, , _
        "ERROR " & Err.Number _
        & "   ShowHideControlsTag"
        
   'press F8 to step through code
   'comment next line when debugged
   Stop: Resume
   
   Resume Proc_Exit

End Function





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function cal_GetRoman(ByVal pNumber As Integer) As String
'120627
'modified from Microsoft Support
' OFF97: VBA Procedure to Convert Numbers to Roman Numerals
' http://support.microsoft.com/kb/184657

   On Error GoTo Proc_Err
   
   Const ROMAN = "IVXLCDM" 'I=1,V=5, X=10, L=100, C=1,000, D=500   M=1,000
   
   Dim i As Integer, Digit As Integer, sStr As String
   
   i = 1
   sStr = ""
   Do While pNumber > 0
      Digit = pNumber Mod 10
      pNumber = pNumber \ 10
      Select Case Digit
         Case 1
           sStr = mID(ROMAN, i, 1) & sStr
         Case 2
           sStr = mID(ROMAN, i, 1) & mID(ROMAN, i, 1) & sStr
         Case 3
           sStr = mID(ROMAN, i, 1) & mID(ROMAN, i, 1) & _
                  mID(ROMAN, i, 1) & sStr
         Case 4
           sStr = mID(ROMAN, i, 2) & sStr
         Case 5
           sStr = mID(ROMAN, i + 1, 1) & sStr
         Case 6
           sStr = mID(ROMAN, i + 1, 1) & mID(ROMAN, i, 1) & sStr
         Case 7
           sStr = mID(ROMAN, i + 1, 1) & mID(ROMAN, i, 1) & _
                  mID(ROMAN, i, 1) & sStr
         Case 8
           sStr = mID(ROMAN, i + 1, 1) & mID(ROMAN, i, 1) & _
                  mID(ROMAN, i, 1) & mID(ROMAN, i, 1) & sStr
         Case 9
           sStr = mID(ROMAN, i, 1) & mID(ROMAN, i + 2, 1) & sStr
      End Select
      i = i + 2
   Loop
   cal_GetRoman = sStr
   
Proc_Exit:
   Exit Function
Proc_Err:
   Resume Proc_Exit
   
End Function





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Function cal_GetBirthstone(pMonth As Integer) As String
   'birthstone for month -- CUSTOMIZE FOR YOUR CULTURE

   Select Case pMonth
   Case 1: cal_GetBirthstone = "Garnet"
   Case 2: cal_GetBirthstone = "Amethyst"
   Case 3: cal_GetBirthstone = "Aquamarine, Bloodstone"
   Case 4: cal_GetBirthstone = "Diamond, Rock Crystal"
   Case 5: cal_GetBirthstone = "Emerald, Chrysoprase"
   Case 6: cal_GetBirthstone = "Pearl, Moonstone, Alexandrite"
   Case 7: cal_GetBirthstone = "Ruby, Cornelian"
   Case 8: cal_GetBirthstone = "Peridot, Sardonyx"
   Case 9: cal_GetBirthstone = "Sapphire, Lapis Lazuli"
   Case 10: cal_GetBirthstone = "Opal, Yourmaline"
   Case 11: cal_GetBirthstone = "Topaz, Citrine"
   Case 12: cal_GetBirthstone = "Turquoise, Zircon, Tanzanite"
   End Select
End Function




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub txtCalendarDate_BeforeUpdate(Cancel As Integer)
'120701
   On Error Resume Next
   If IsNull(Me.ActiveControl) Then
      Me.ActiveControl.Undo
      Cancel = True
      Exit Sub
   End If
   If Not IsDate(Me.ActiveControl) Then
      MsgBox Me.ActiveControl & " is not a valid date", , "Cannot change"
      Me.ActiveControl.Undo
      Cancel = True
      Exit Sub
   End If
End Sub





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub txtDays_DblClick(Cancel As Integer)
'120627
   On Error Resume Next
   If IsNull(Me.ActiveControl) Then Exit Sub
   Me.txtDays = -Me.txtDays
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
        mError = "Source:      " & Err.Source _
            & vbCrLf & "Number:     " & Err.Number _
            & vbCrLf & "Issue:      " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (mError)
End Sub




