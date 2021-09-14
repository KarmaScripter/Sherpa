Attribute VB_Name = "FileSelector"
Public Function ChooseFile(DialogMessage As String) As String
    Dim targetPath As String
    targetPath = ""

    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Filters.Add "Excel Workbooks", "*.xlsx; *.xlsm; *.xls", 2
        .FilterIndex = 2
        .Title = DialogMessage
        .Show
        
        If .SelectedItems.Count > 0 Then
            targetPath = .SelectedItems(1)
        End If
    End With
    ChooseFile = targetPath
End Function
