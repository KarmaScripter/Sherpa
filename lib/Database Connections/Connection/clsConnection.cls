VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Type vars
    sql              As String
    connectionString As String
    connectionPath   As String
    connection       As ADODB.connection
    recordset        As ADODB.recordset
End Type

Dim this As vars
'-----------------------------------------Construct and Destruct-------------------------------------------'
Public Sub create(ByVal returnRecordSet As Boolean)
    createConnection
    If returnRecordSet Then createRecordSet
End Sub
'------------------------------------------Props---------------------------------------------------'
Property Let sql(ByVal value As String): this.sql = value: End Property
Property Get sql() As String: sql = this.sql: End Property
Property Let connectionPath(ByVal value As String): this.connectionPath = value: End Property
Property Let connectionString(ByVal value As String): this.connectionString = value: End Property
Property Get recordset() As ADODB.recordset: Set recordset = this.recordset: End Property
Property Get connection() As ADODB.connection: Set connection = this.connection: End Property
'------------------------------------------DEFAULTS METHODS -------------------------------------------------'
Private Function createConnection()
    Set this.connection = New ADODB.connection
    this.connection.Open this.connectionString
    If this.connection.State = adStateOpen Then Debug.Print "Connection succeds..."
End Function

Private Function createRecordSet()
    Set this.recordset = New ADODB.recordset
    this.recordset.Open this.sql, this.connection, adOpenKeyset, adLockOptimistic
    Debug.Print "Data found..."
End Function

Public Function execute(): this.connection.execute this.sql: End Function

Public Function executeWithParameter(ParamArray params() As Variant)
    Dim cmd As New ADODB.Command
    Dim param As Variant
    
    Set this.recordset = New ADODB.recordset
    With cmd
        .CommandText = this.sql
        .CommandType = adCmdText
        .NamedParameters = True
        .Prepared = True
    End With
    
    For Each param In params
         cmd.Parameters.Append newParams(cmd, param)
    Next param

    cmd.activeConnection = this.connection
    Set this.recordset = cmd.execute
End Function

Private Function newParams(ByVal cmd As ADODB.Command, ByVal value As String) As ADODB.Parameter
     Dim typeInput As Long
    
     If IsNumeric(value) Then typeInput = 3
     If IsDate(value) Then typeInput = 7
     If value = vbNullString Then typeInput = 200
     Set newParams = cmd.CreateParameter(, typeInput, adParamInput, , value)
End Function

Public Function killConnection(Optional ByVal connection As ADODB.connection, Optional ByVal recordset As ADODB.recordset)
    On Error Resume Next
    If this.recordset.State = 1 Then this.recordset.Close: Set this.recordset = Nothing: Debug.Print "Recordset closed."
    If this.connection.State = adStateOpen Then this.connection.Close: Set this.connection = Nothing: Debug.Print "Connection closed."
    If Not connection Is Nothing Then connection.Close: Set connection = Nothing
    If Not recordset Is Nothing Then recordset.Close: Set recordset = Nothing
End Function
'------------------------------------------COMMANDS------------------------------------------------------s-----'
Public Function fieldToArray(Optional FirstBlank As Boolean = False) As Variant
    Dim field       As ADODB.field
    Dim matriz      As Variant
    Dim IncControl  As Byte
    
    ReDim matriz(this.recordset.Fields.Count)
    
    If FirstBlank = True Then matriz(0) = "": IncControl = 1
    
    For Each field In this.recordset.Fields
        matriz(IncControl) = field.Name
        IncControl = IncControl + 1
       ' If IncControl >= UBound(matriz) - 1 Then Exit For
    Next field
    
    fieldToArray = matriz
End Function
'---------------------------------------------REPORTS--------------------------------------------------------'
Public Function recordsetToExcel(Optional ByVal handlerExcel As Boolean) As Variant
    Dim wks       As Excel.Worksheet
    Dim header    As Variant
    
    Set wks = Workbooks.Add.Sheets(1)

    header = fieldToArray
    wks.Activate
    wks.Cells(1, 1).Resize(1, UBound(header) + 1) = header
    wks.Cells(2, 1).CopyFromRecordset this.recordset
    wks.ListObjects.Add xlSrcRange, Range("a1").CurrentRegion, xlGuess, tablestylename:="TableStyleMedium20"
    wks.Columns.AutoFit
    
    If handlerExcel Then Set recordsetToExcel = wks
    Set wks = Nothing
End Function

