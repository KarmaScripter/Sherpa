VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_CreateDatabase As String
Private m_CreateTable As String
Private m_CreateView As String
Private m_Select As String
Private m_Insert As String
Private m_Update As String
Private m_Delete As String
Private m_Drop As String
Private m_AlterTable As String
Private m_AlterColumn As String
Private m_Source As String
Private m_Tables() As String
Private m_Criteria As New Scripting.Dictionary
Private m_Command As String
Private m_Commands() As String
Private pAnd As String
Private pLike As String
Private pWhere As String
Private m_Error As String
Private m_Notification As String



Private Sub Class_Initialize()
    pAnd = " AND "
    m_Commands(0) = "CREATE"
    m_Commands(1) = "SELECT"
    m_Commands(2) = "UPDATE"
    m_Commands(3) = "INSERT"
    m_Commands(4) = "DELETE"
    m_Commands(7) = "DROP"
    m_Commands(8) = "ALTER"
    Set m_Criteria = CreateObject("Scripting.Dicitionary")
End Sub
Private Sub Class_Terminate()
    If Not m_Criteria Is Nothing Then
        Set m_Criteria = Nothing
    End If
End Sub






'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        Replace
'   Parameters  Replace
'   Retval      Void
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Property Let Source(pSource As String)
    On Error GoTo ErrorHandler:
    If Not pSource & "" = "" Then
        m_Source = pSource
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   SqlBuilder" _
            & vbCrLf & "Member:     Source" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name        Replace
'   Parameters  Void
'   Retval      Replace
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Property Get Source() As String
    On Error GoTo ErrorHandler:
    If Not m_Source & "" = "" Then
        Source = m_Source
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   SqlBuilder" _
            & vbCrLf & "Member:     Source" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        Command
'   Parameters  String
'   Retval      Void
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Property Let Command(pCommand As String)
    On Error GoTo ErrorHandler:
    If Not pCommand & "" = "" Then
        Dim i As Integer
        For i = LBound(m_Commands) To i = UBound(m_Commands)
            If m_Commands(i) = pCommand Then
                m_Command = pCommand
                Exit Property
            End If
        Next i
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   SqlBuilder" _
            & vbCrLf & "Member:     Command" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        Command
'   Parameters  String
'   Retval      Void
'   Purpose     Replace
'---------------------------------------------------------------------------------
Private Function GetNames(pData As Scripting.Dictionary) As String()
    On Error GoTo ErrorHandler:
    If Not pData Is Nothing Then
        Dim i As Integer
        Dim p_Names() As String
        For i = 0 To i = pData.count
            p_Names(i) = pData.Keys(i)
        Next i
        GetNames = p_Names
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   SqlBuilder" _
            & vbCrLf & "Member:     GetNames(Scripting.Dictionary)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function





'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetValues
'   Parameters  String
'   Retval      Void
'   Purpose     Replace
'---------------------------------------------------------------------------------
Private Function GetValues(pData As Scripting.Dictionary) As Variant()
    On Error GoTo ErrorHandler:
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   SqlBuilder" _
            & vbCrLf & "Member:     GetValues(Scripting.Dictionary)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name        Command
'   Parameters  Void
'   Retval      String
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Property Get Command() As String
    On Error GoTo ErrorHandler:
    If Not m_Command & "" = "" Then
        Command = m_Command
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   SqlBuilder" _
            & vbCrLf & "Member:     Command" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        CreateSelectCommand
'   Parameters  Scripting.Dictionary
'   Retval      String
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Function CreateSelectCommand(pCriteria As Scripting.Dictionary) As String
    On Error GoTo ErrorHandler:
    If Not pCriteria Is Nothing And _
        pCriteria.count > 0 Then
            Dim p_Names() As String
            Dim p_Values() As Variant
            Dim p_NameValuePairs As String
            Dim p_Sql As String
            Dim p_Predicate As String
            Dim i As Integer
            p_Sql = "SELECT * FROM m_Source" & pWhere
            For i = 0 To i = pCriteria.count
                p_Names(i) = pCriteria.Keys(i)
                p_Values(i) = pCriteria.Items(i)
            Next i
            For i = 0 To i = pCriteria.count
                Dim p_Pair As String
                p_Pair = p_Names(i) & " = '" & p_Values(i) & "'"
                p_NameValuePairs = p_NameValuePairs & p_Pair & pAnd
            Next i
            p_Predicate = Replace(Expression:=p_NameValuePairs, Find:=pAnd, _
                Replace:=";", Start:=-1)
            CreateSelectCommand = p_Sql & p_Predicate
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   SqlBuilder" _
            & vbCrLf & "Member:     CreateSelectCommand(Scripting.Dictionary)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function







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



