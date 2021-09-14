VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CriteriaBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Names() As String
Private m_Values() As Variant
Private m_Criteria As New Scripting.Dictionary
Private pWhere As String
Private pAnd As String
Private pLike As String
Private m_Error As String
Private m_Notification As String


Private Sub Class_Initialize()
    pAnd = " AND "
    pWhere = " WHERE "
    Set m_Criteria = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
    If Not m_Criteria Is Nothing Then
        Set m_Criteria = Nothing
    End If
End Sub



'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        Names
'   Parameters  String()
'   Retval      Void
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Property Let Names(pNames() As String)
    On Error GoTo ErrorHandler:
    Dim i As Integer
    If IsArray(pNames) And _
        UBound(pNames) > 0 Then
            For i = LBound(pNames) To i = UBound(pNames)
                If Not pNames(i) & "" = "" Then
                    m_Names(i) = pNames(i)
                End If
            Next i
            Names = m_Names
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        Names
'   Parameters  Void
'   Retval      String()
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Property Get Names() As String()
    On Error GoTo ErrorHandler:
    Dim i As Integer
    If IsArray(m_Names) And _
        UBound(m_Names) > 0 Then
            Names = m_Names
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        Values(Variant())
'   Parameters  Replace
'   Retval      Void
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Property Let Values(pValues() As Variant)
    On Error GoTo ErrorHandler:
    Dim i As Integer
    If IsArray(pValues) And _
        UBound(pValues) > 0 Then
            For i = LBound(pValues) To i = UBound(pValues)
                If Not pValues(i) & "" = "" Then
                    m_Values(i) = pValues(i)
                End If
            Next i
            Values = m_Values
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        Values()
'   Parameters  Replace
'   Retval      Void
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Property Get Values() As Variant()
    On Error GoTo ErrorHandler:
    Dim i As Integer
    If IsArray(m_Values) Then
        Values = m_Values
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property






'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        CreateCriteria()
'   Parameters  Replace
'   Retval      Void
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Function GetNameValuePairs() As Scripting.Dictionary
    On Error GoTo ErrorHandler:
    Dim i As Integer
    Dim p_Criteria As New Scripting.Dictionary
    Set p_Criteria = CreateObject("Scripting.Dictionary")
    If IsArray(m_Names) And _
        IsArray(m_Values) And _
        LBound(m_Names) = LBound(m_Values) And _
        UBound(m_Names) = UBound(m_Values) Then
            For i = LBound(m_Names) To UBound(m_Names)
                p_Criteria.add Item:=m_Values(i), Key:=m_Names(i)
            Next i
            Set GetNameValuePairs = p_Criteria
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function






'----------------------------------------------------------------------------------
'   Type        Property Let
'   Name        Values()
'   Parameters  DAO.Recordset
'   Retval      Sciprting.Dictionary
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Function GetDataAccessCriteria(pData As DAO.Recordset) As Scripting.Dictionary
    On Error GoTo ErrorHandler:
    Dim i As Integer
    Dim p_Criteria As New Scripting.Dictionary
    Set p_Criteria = CreateObject("Scripting.Dictionary")
    If Not pData Is Nothing Then
        For i = 0 To i = pData.Fields.count
            p_Criteria.add pData.Fields(i).Name, pData.Fields(i).Value
        Next i
        Set GetDataAccessCriteria = p_Criteria.Items
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function







'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetAccessDataCriteria
'   Parameters  ADODB.Recordset
'   Retval      Scripting.Dictionary
'   Purpose     Replace
'---------------------------------------------------------------------------------
Public Function GetAccessDataCriteria(pData As ADODB.Recordset) As Scripting.Dictionary
    On Error GoTo ErrorHandler:
    Dim i As Integer
    Dim p_Criteria As New Scripting.Dictionary
    Set p_Criteria = CreateObject("Scripting.Dictionary")
    If Not pData Is Nothing Then
        For i = 0 To i = pData.Fields.count
            p_Criteria.add pData.Fields(i).Name, pData.Fields(i).Value
        Next i
        Set GetAccessDataCriteria = p_Criteria.Items
    End If
ErrorHandler:
    ProcessError
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


