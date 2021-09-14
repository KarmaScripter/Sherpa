Attribute VB_Name = "Calculator"
Option Compare Database

'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------   FIELDS   ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------

Public pid As Variant


'----------------------------------------------------------------------------------
'   Type        Function
'   Name        IsProcessRunning
'   Parameters  String - AppName
'   Purpose     Function to check for running application by its process name
'---------------------------------------------------------------------------------


Function IsProcessRunning(AppName As String)
    On Error GoTo Skip
        Dim objList As Object
        Set objList = GetObject("winmgmts:").ExecQuery("Select ProcessID from Win32_Process where Name='" & AppName & "'")
        If objList.count > 0 Then
            IsProcessRunning = True
            For Each objProcess In objList
                If pid <> objProcess.ProcessID Then
                    pid = objProcess.ProcessID
                    Exit Function
                End If
            Next
        Else
            IsProcessRunning = False
            Exit Function
        End If
Skip:
    IsProcessRunning = False
End Function


'----------------------------------------------------------------------------------
'   Type        SubProcedure
'   Name        Calculate
'   Parameters  Void
'   Purpose     Launches the Windows 10 calculator 'calc.exe'
'----------------------------------------------------------------------------------

Sub Calculate()
    If IsProcessRunning("calc.exe") = True Then
        On Error GoTo Reload                ' Open new instance of calculator in event of error
        AppActivate (vPID)                  ' Reactivate calculator process using Public declared variant
        SendKeys "%{Enter}"                 ' Bring it back into focus if user minimises it
    Else
Reload:
        vPID = Shell("calc.exe", 1)         ' Run Calculator
    End If
    On Error GoTo 0
End Sub
