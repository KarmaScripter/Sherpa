Attribute VB_Name = "SheetObjectHelper"
Public Function GetNewSheet(Optional Name As String = "") As Worksheet
    Dim newSheet As Worksheet
    Set newSheet = CreateNewBook().Sheets(1)
    If Name <> "" Then
        newSheet.Name = Name
    End If
    Set GetNewSheet = newSheet
End Function

Public Function CreateNewBook() As Workbook
    Set CreateNewBook = Application.Workbooks.Add
End Function

Public Sub DiscardWorkbook(theWorkbook As Workbook)
    theWorkbook.Close False
End Sub

Public Function HasSheet(TheBook As Workbook, SheetName As String) As Boolean
    On Error GoTo Catch
    Dim testSheet As Worksheet
    Set testSheet = TheBook.Sheets(SheetName)
    HasSheet = True
    Exit Function
Catch:
    HasSheet = False
End Function

Public Function GetSheet(TheBook As Workbook, SheetName As String) As Worksheet
    If Not HasSheet(TheBook, SheetName) Then
        Set GetSheet = Nothing
    Else
        Set GetSheet = TheBook.Sheets(SheetName)
    End If
End Function
