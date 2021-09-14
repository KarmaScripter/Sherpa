Attribute VB_Name = "ListObjectHelper"
' Dependencies:
'   - Module: SheetObjectHelper.bas
'   - Class: ValidateHeadersResult.cls
Public Function GetNewList(targetSheet As Worksheet, Optional ColumnHeaders As Collection) As ListObject
    If ColumnHeaders Is Nothing Then _
        Set ColumnHeaders = New Collection
    
    Dim myList As ListObject
    Set myList = targetSheet.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=getSrcRange(targetSheet, ColumnHeaders), _
        XLListObjectHasHeaders:=xlYes _
    )
    
    Dim headerIndex As Integer
    For headerIndex = 1 To ColumnHeaders.Count
        myList.ListColumns(headerIndex).Name = ColumnHeaders(headerIndex)
    Next headerIndex
    
    Set GetNewList = myList
End Function

Private Function getSrcRange(theSheet As Worksheet, headers As Collection) As Range
    If headers.Count = 0 Then
        Set getSrcRange = theSheet.Range("A1:A2")
    Else
        Set getSrcRange = theSheet.Range( _
            theSheet.Cells(1, 1), _
            theSheet.Cells(2, headers.Count) _
        )
    End If
End Function

Public Function GetList(SearchWorkbook As Workbook, ListName As String) As ListObject
    On Error GoTo Catch
    
    Dim iSheet As Worksheet
    Dim iList As ListObject
    Dim result As ListObject
    
    For Each iSheet In SearchWorkbook.Sheets
        For Each iList In iSheet.ListObjects
            If iList.Name = ListName Then
                Set result = iList
                GoTo Found
            End If
        Next iList
    Next iSheet

Found:
    Set GetList = result
    Exit Function
Catch:
    Set GetList = Nothing
End Function


Public Function ValidateHeaders(TheList As ListObject, HasHeaders As Collection) As ValidateHeadersResult
    Dim myResult As New ValidateHeadersResult
    
    Dim col As ListColumn
    Dim Header As Variant
    
    For Each col In TheList.ListColumns
        If ContainsValue(HasHeaders, col.Name) Then
            myResult.hasHeader col.Name
        Else
            myResult.HeaderExtra col.Name
        End If
    Next col
    
    For Each Header In HasHeaders
        If Not ContainsValue(myResult.ImplementedHeaders, Header) Then
            myResult.HeaderNotFound Header
        End If
    Next Header
    
    If myResult.MissingHeaders.Count > 0 Then
        myResult.HasRequiredHeaders = False
    Else
        myResult.HasRequiredHeaders = True
    End If
    
    Set ValidateHeaders = myResult
End Function

Private Function hasHeader(myList As ListObject, theHeader As String) As Boolean
    On Error GoTo Catch
    
    Dim myColumn As ListColumn
    Set myColumn = TheList.ListColumns(theHeader)
    hasHeader = True
    Exit Function
Catch:
    hasHeader = False
End Function

Public Function ListCell(TheList As ListObject, Header As String, Row As Long) As Range
    If Row > TheList.ListRows.Count Then _
        ThrowError ValueOutsideListDataRange, _
            "ListObjectHelper(" & TheList.Name & ", " & Header & ", " & Row & ")", _
            CStr(Row) & " is greater than the number of rows currently in List " & TheList.Name & ": " & TheList.ListRows.Count
    
    Dim returnColumn As ListColumn
    Set returnColumn = TheList.ListColumns(Header)
    
    Dim returnRow As Long
    returnRow = TheList.HeaderRowRange.Row + Row
    
    Set ListCell = TheList.Parent.Cells(returnRow, returnColumn.Range.Column)
End Function

Public Function AddRow(TheList As ListObject) As Long
    TheList.ListRows.Add
    AddRow = TheList.ListRows.Count
End Function
