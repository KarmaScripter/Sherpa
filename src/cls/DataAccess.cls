VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private mBudgetExecution As DAO.Workspace
Private mDatabase As DAO.Database
Private m_Error As String
Private m_Notification As String





Private Sub Class_Initialize()
    Set mDatabase = CurrentDb
    Set mBudgetExecution = DBEngine.Workspaces(0)
End Sub

Private Sub Class_Terminate()
    If Not IsNull(mDatabase) Then
        Set mDatabase = Nothing
    End If
    If Not IsNull(mBudgetExecution) Then
        Set mBudgetExecution = Nothing
    End If
End Sub





'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        IsConnected
'   Parameters:  Void
'   RetVal:      Boolean
'   Purpose:     Function returns boolean value indicating if Workspace is open
'---------------------------------------------------------------------------------
Public Property Get IsConnected() As Boolean
    On Error GoTo ErrorHandler:
    If Not IsObject(mDatabase) Then
        IsConnected = True
    End If
    If IsObject(mDatabase) Then
        IsConnected = False
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     IsConnected" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name        IsOpen
'   Parameters  Void
'   RetVal      Boolean
'   Purpose     Function returns boolean value indicating if Workspace is open
'---------------------------------------------------------------------------------
Public Property Get IsOpen() As Boolean
    On Error GoTo ErrorHandler:
    If Not IsObject(mBudgetExecution) Then
        IsOpen = True
    End If
    If IsObject(mBudgetExecution) Then
        IsOpen = False
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     IsOpen()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetDataSet
'   Parameters  String - table;
'   RetVal      DAO.Recordset
'   Purpose     Function returns the records from the Data Table using the table's
'               name or a sql select query
'---------------------------------------------------------------------------------
Public Function GetDataSet(table As String) As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim rs As DAO.Recordset
    If Not IsNull(mDatabase) And _
        Not table & "" = "" Then
           Set rs = mDatabase.OpenRecordset(table, dbOpenDynaset)
    End If
    Set GetDataSet = rs
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     GetDataSet" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetRecord
'   Parameters  String - table;
'   RetVal      DAO.Recordset
'   Purpose     Function returns record using ID parameter and Primary Key
'               defined in a sql select query
'---------------------------------------------------------------------------------
Public Function GetDataRecord(table As String, primarykey As String, ID As Integer) As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Set DB = mDatabase
    Dim rs As DAO.Recordset
    Dim SQL As String
    If Not IsNull(mDatabase) And _
        Not table & "" = "" And _
        Not primarykey & "" = "" And _
        ID > 0 Then
            SQL = "SELECT * FROM table WHERE key = 'id'"
            Set rs = mDatabase.OpenRecordset(SQL, dbOpenDynaset)
            Set GetDataRecord = rs
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     GetDataRecord" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        CloseDatabase
'   Parameters  Void
'   RetVal      Void
'   Purpose     Method used to close the database
'---------------------------------------------------------------------------------
Public Sub CloseDatabase()
    On Error GoTo ErrorHandler:
    If Not mDatabase Is Nothing Then
        Set mDatabase = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     CloseDatabase" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        CloseWorkspace
'   Parameters  Void
'   RetVal      Void
'   Purpose     Method used to close the database
'---------------------------------------------------------------------------------
Public Sub CloseWorkspace()
    On Error GoTo ErrorHandler:
    If Not IsNull(mBudgetExecution) Then
        Set mBudgetExecution = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     CloseWorkspace" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        UpdateDate
'   Parameters  table - String, field - String, value - Variant
'   RetVal      Void
'   Purpose     Method used to update data in current recordset
'---------------------------------------------------------------------------------
Public Sub UpdateData(table As String, field As String, Value As Variant)
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Set DB = mDatabase
    Dim rs As DAO.Recordset
    If Not table & "" = "" Then
        Set rs = GetDataSet(table)
    End If
    If Not field & "" = "" _
        And Not IsNull(Value) Then
            rs.Edit
            rs.Fields(field) = Value
            rs.Update
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     UpdateData" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        AddRecord
'   Parameters  table - String, field - String, value - Variant
'   RetVal      Void
'   Purpose     Method used to add data to recordset
'---------------------------------------------------------------------------------
Public Sub AddRecord(table As String, field As String, Value As Variant)
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Dim rs As DAO.Recordset
    Set DB = mDatabase
    Set rs = GetDataSet(table)
    With rs
      .AddNew
      !Fields(field) = Value
      .Update
    End With
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     AddRecord" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        UpdateRecord
'   Parameters  table - String, field - String, value - Variant
'   RetVal      Void
'   Purpose     Method used to add data to recordset
'---------------------------------------------------------------------------------
Public Sub UpdateRecord(table As String, field As String, NewValue As Variant, ID As Integer)
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Dim rs As DAO.Recordset
    Set DB = mDatabase
    Set rs = GetDataRecord(table, field, ID)
    If Not IsNull(rs) Then
        With rs
          .AddNew
          !Fields(field) = NewValue
          .Update
        End With
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     UpdateRecord" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        DeleteRecord
'   Parameters  table - String, pk - String, id - integer
'   RetVal      Void
'   Purpose     Method used to delete records from the data set
'---------------------------------------------------------------------------------
Public Sub DeleteRecord(table As String, pk As String, ID As Integer)
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Dim rs As DAO.Recordset
    Set DB = mDatabase
    Dim SQL As String
    SQL = "DELETE FROM table WHERE pk = id"
    Set rs = DB.OpenRecordset(SQL)
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataAccess" _
            & vbCrLf & "Member:     DeleteRecord" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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



