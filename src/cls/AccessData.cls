VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_BudgetPath As BudgetPath
Private m_Error As String
Private m_Notification As String
Private m_DataModels As Collection
Private m_Connection As ADODB.Connection
Private m_Command As ADODB.Command
Private m_Data As ADODB.Recordset
Private m_Record As ADODB.Record

Private Sub Class_Initialize()
    Set m_BudgetPath = New BudgetPath
    Set m_Data = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
    KillData
End Sub
 



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name        Data
'   Parameters  Void
'   Retval      ADODB.Connection
'   Purpose     Function to create and return the current ADO connection
'---------------------------------------------------------------------------------
Public Property Get Data() As ADODB.Connection
    On Error GoTo ErrorHandler:
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    cnn.Open m_BudgetPath.Data
    Set Data = cnn
ErrorHandler:
    ProcessError
    Set cnn = Nothing
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        References
'   Parameters  Void
'   Retval      ADODB.Connection
'   Purpose     Function to create property getter and return the current ADO connection
'---------------------------------------------------------------------------------
Public Property Get References() As ADODB.Connection
    On Error GoTo ErrorHandler:
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    cnn.Open m_BudgetPath.References
    Set References = cnn
ErrorHandler:
    ProcessError
    Set cnn = Nothing
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        Divisions
'   Parameters  Void
'   Retval      ADODB.Connection
'   Purpose     Function to create and return the current ADO connection
'---------------------------------------------------------------------------------
Public Property Get Divisions() As ADODB.Connection
    On Error GoTo ErrorHandler:
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    cnn.Open m_BudgetPath.Divisions
    Set Divisions = cnn
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetReferenceModel
'   Parameters  Void
'   Retval      ADODB.Recordset
'   Purpose     Function to create and return the pTableName
'---------------------------------------------------------------------------------
Public Function GetReferenceModel(pTableName As String) As ADODB.Recordset
    On Error GoTo ErrorHandler:
    If Not pTableName & "" = "" Then
        Dim rst As ADODB.Recordset
        Dim SQL As String
        SQL = "SELECT * FROM pTableName"
        Set rst = New ADODB.Recordset
        rst.Open SQL, References, adOpenStatic, adLockOptimistic
        Set GetReferenceModel = rst
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function





'----------------------------------------------------------------------------------
'   Type        Function
'   Name        CreateSqlCommand
'   Parameters  String - sql
'   Purpose     Function to create and return a command
'               based on a sql string argument
'---------------------------------------------------------------------------------
Public Function CreateCommand(pSQL As String) As ADODB.Command
    On Error GoTo ErrorHandler:
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = CurrentProject.Connection
    If Not pSQL & "" = "" Then
        cmd.CommandText = pSQL
    End If
    If Not IsNull(cmd) Then
        Set CreateCommand = cmd
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetModel
'   Parameters  String - Table
'   Purpose     Function to create and return a recordset
'               based on a pTableName name provided as parameter argument
'---------------------------------------------------------------------------------
Public Function GetModel(pTableName As String) As ADODB.Recordset
    On Error GoTo ErrorHandler:
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open (pTableName)
    If rst.RecordCount > 0 Then
        Set GetModel = rst
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type        Sub-Procedure
'   Name        ImportExcelData
'   Parameters  String
'   Retval      Void
'   Purpose     Function to Import Excel data into MS Access
'---------------------------------------------------------------------------------
Public Sub ImportExcelData(pFilePath As String)
    On Error GoTo ErrorHandler:
    Dim p_FilePath As String
    Dim p_List() As String
    Dim p_Count As Integer
    Dim p_Worksheet As Excel.Worksheet
    p_FilePath = pFilePath
    Dim p_Excel As Excel.Application
    Set p_Excel = CreateObject("Excel.Application")
    p_Excel.Visible = False
    p_Excel.DisplayAlerts = False
    p_Excel.Workbooks.Open p_FilePath
    If p_Excel.Worksheets.count > 1 Then
        Dim i As Integer
        For Each p_Worksheet In p_Excel.Worksheets
            ReDim Preserve p_List(p_Count)
            p_List(p_Count) = p_Worksheet.Name
            p_Count = p_Count + 1
        Next
        For i = LBound(p_List()) To UBound(p_List())
            DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, p_List(i), _
                p_FilePath, True, p_List(i) & "!A1:IU9999"
        Next i
    Else
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, p_List(0), _
            p_FilePath, True, p_List(i) & "!A1:IU9999"
    End If
    p_Excel.ActiveWorkbook.Close (True)
    p_Excel.Quit
    Set p_Excel = Nothing
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type        Sub-Procedure
'   Name        KillData
'   Parameters  Void
'   Retval      Void
'   Purpose     Frees resources used by the Database
'---------------------------------------------------------------------------------
Private Sub KillData()
    On Error GoTo ErrorHandler:
    If Not m_Data Is Nothing Then
        Set m_Data = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type        Sub-Procedure
'   Name        KillFiles
'   Parameters  Void
'   Retval      Void
'   Purpose     Frees resources used by the BudgetPath object
'---------------------------------------------------------------------------------
Private Sub KillFiles()
    On Error GoTo ErrorHandler:
    If Not m_BudgetPath Is Nothing Then
        Set m_BudgetPath = Nothing
    End If
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



