VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BudgetFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_Base As String
Private m_Path As String
Private m_Name As String
Private m_Extension As String
Private m_ParentFolder As String
Private m_FileSys As Scripting.FileSystemObject
Private m_Folder As Scripting.folder
Private m_File As Scripting.File
Private m_Files As Scripting.Files
Private m_SubFolders As Scripting.folders
Private m_Drive As String
Private m_Type As String
Private m_DateCreated As Date
Private m_LastAccessed As Date
Private m_LastModified As Date
Private m_Stream As Scripting.TextStream
Private m_Error As String
Private m_Notification As String








'---------------------------------------------------------------------------------
'   Type:            Constructor
'   Name:
'   Parameters:      Void
'   Retval:          BudgetFile
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub



'---------------------------------------------------------------------------------
'   Type:            Destructor
'   Name:
'   Parameters:      Void
'   Retval:          BudgetFile
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Terminate()
End Sub



'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Base(pFilePath As String)
    On Error GoTo ErrorHandler:
    If Not pFilePath & "" = "" Then
        m_Base = pFilePath
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Base() As String
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" Then
        Base = m_Base
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FileCount() As Integer
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        FileCount = m_FileSys.GetFolder(m_Base).Files.count
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:            Extension
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get FolderCount() As Integer
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set m_SubFolders = m_FileSys.GetFolder(m_Base).SubFolders
        FolderCount = m_SubFolders.count
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:            Exists
'   Parameters:      String
'   Retval:          Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Exists() As Boolean
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        If Not Dir(m_Base) & "" = "" Then
            Exists = True
        Else
            Exists = False
        End If
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:            Exists
'   Parameters:      String
'   Retval:          Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get IsRooted() As Boolean
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        If m_FileSys.GetFolder(m_Base).IsRootFolder = True Then
            IsRooted = True
        Else
            IsRooted = False
        End If
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:            CurrentDirectory
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get CurrentDirectory() As String
    On Error GoTo ErrorHandler:
    Dim p_Current As String
    p_Current = CurDir()
    CurrentDirectory = p_Current
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:            Path
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Path() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not Dir(m_Base) & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            m_Path = m_FileSys.GetFolder(m_Base).ShortPath
            Path = m_Path
            Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Drive() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_Drive = m_FileSys.GetFolder(m_Base).Drive
        Drive = m_Drive
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ParentFolder() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_ParentFolder = m_FileSys.GetParentFolderName(m_Base)
        ParentFolder = m_ParentFolder
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:            DateCreated
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DateCreated() As Date
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_DateCreated = m_FileSys.GetFolder(m_Base).DateCreated
        DateCreated = m_DateCreated
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:            LastAccessed
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get LastAccessed() As Date
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_LastAccessed = m_FileSys.GetFolder(m_Base).DateLastAccessed
        LastAccessed = m_LastAccessed
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get LastModified() As Date
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_LastModified = m_FileSys.GetFolder(m_Base).DateLastModified
        LastModified = m_LastModified
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Name() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_Name = m_FileSys.GetFolder(m_Base).ShortName
        Name = m_Name
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:            Read-Only Property
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Size() As Variant
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set Size = m_FileSys.GetFolder(m_Base).Size
    End If
    Set m_FileSys = Nothing
ErrorHandler:
    ProcessError
    Exit Property
End Property



 
'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:            Extension
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get SubFolders() As Scripting.folders
    On Error GoTo ErrorHandler:
    Dim folders As Scripting.folders
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set m_SubFolders = m_FileSys.GetFolder(m_Base).SubFolders
        Set SubFolders = m_SubFolders
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:            FolderExists
'   Parameters:      String
'   Retval:          Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Function Verify(pPath As String) As Boolean
    On Error GoTo ErrorHandler:
    If Not pPath & "" = "" And _
        Not m_Base & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            If m_FileSys.FolderExists(pPath) = True Then
                Verify = True
            Else
                Verify = False
            End If
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetSize
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetName(pPath As String) As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not Dir(pPath) & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            GetName = m_FileSys.CreateFolder(pPath).Name
            Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetSize
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetSubFiles(pPath As String) As Scripting.Files
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not Dir(pPath) & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            Set GetSubFiles = m_FileSys.CreateFolder(pPath).Files
            Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetSubFolders
'   Parameters:      String
'   Retval:          Scripting.Folders
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetSubFolders(pPath As String) As Scripting.folders
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not Dir(pPath) & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            Set GetSubFolders = m_FileSys.CreateFolder(pPath).SubFolders
            Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetSize
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetSize(pPath As String) As Variant
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not Dir(pPath) & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            Set GetSize = m_FileSys.GetFolder(pPath).Size
    Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:
'   Parameters:      Void
'   Retval:          Scripting.Folder
'   Purpose:
'---------------------------------------------------------------------------------
Public Function Create() As Scripting.folder
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set m_Folder = m_FileSys.CreateFolder(m_Base)
        Set Create = m_Folder
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Function CreateDirectory(pPath As String) As Scripting.folder
    On Error GoTo ErrorHandler:
    Dim p_Folder As Scripting.folder
    If Not pPath & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set p_Folder = m_FileSys.CreateFolder(pPath)
        Set CreateDirectory = p_Folder
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            GetCurrentDirectory()
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Sub ChangeDirectory(pPath As String)
    On Error GoTo ErrorHandler:
    If Not pPath & "" = "" Then
        ChDir (m_Base)
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



 
'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:            Extension
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub AddSubFolder(pFolderName As String)
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Dim p_Folder As New Scripting.FileSystemObject
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set m_SubFolders = m_FileSys.GetFolder(m_Base).SubFolders
        Set p_Folder = m_SubFolders.add(pFolderName)
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get SubFiles() As Collection
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Dim p_File As Scripting.File
        Dim p_FileCollection As Collection
        Set p_FileCollection = New Collection
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        For Each p_File In m_FileSys.GetFolder(m_Base).Files
            p_FileCollection.add p_File, p_File.Name
        Next p_File
        Set SubFiles = p_FileCollection
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ContainsFile
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function ContainsFile(pFile As String) As Boolean
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Dim p_Files As Scripting.Files
        Dim p_File As Scripting.File
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.Folder")
        Set p_Files = m_FileSys.GetFolder(m_Base).Files
        If p_Files.count > 0 Then
            For Each p_File In p_Files
                If p_File.Name = pFile Then
                    ContainsFile = True
                    Exit Function
                End If
            Next p_File
        End If
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function






'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ContainsSubFolder
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function ContainsSubFolder(pFolder As String) As Boolean
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Dim p_Folders As Scripting.folders
        Dim p_Folder As Scripting.folder
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set p_Folders = m_FileSys.GetFolder(pFolder)
        If p_Folders.count > 0 Then
            For Each p_Folder In p_Folders
                If p_Folder.Name = pFolder Then
                    ContainsSubFolder = True
                    Exit Function
                End If
            Next p_Folder
        End If
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function






'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            Move
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Move(des As String)
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            FileCopy
'   Parameters:      String, String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Copy(dest As String)
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            Delete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Delete()
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_FileSys.DeleteFolder (m_Base)
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            CreateTextFile
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function CreateTextFile(pFilePath As String) As Scripting.TextStream
    On Error GoTo ErrorHandler:
    If Not Dir(pFilePath) & "" = "" And _
        Not m_Base & "" = "" Then
            Dim p_TextFile As Scripting.TextStream
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            Set p_TextFile = m_FileSys.CreateTextFile(filename:=pFilePath, Overwrite:=True)
            CreateTextFile = p_TextFile
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Browse()
    On Error GoTo ErrorHandler:
    Dim p_Dialog As FileDialog
    Dim p_Result As Integer
    Set m_FileSys = New FileSystemObject
    Set m_FileSys = CreateObject("Scripting.FileSystemObject")
    Set p_Dialog = Application.FileDialog(msoFileDialogFilePicker)
    p_Dialog.AllowMultiSelect = False
    p_Dialog.Title = "Select a file"
    p_Dialog.InitialFileName = "C:\"
    p_Dialog.Filters.Clear
    p_Dialog.Filters.add "Text files", "*.txt"
    p_Dialog.Filters.add "All files", "*.*"
    If p_Dialog.Show = -1 Then
       m_FileSys.OpenTextFile (p_Dialog.Item)
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

 


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            FileRename
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Rename(pName As String)
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not pName & "" = "" Then
         If Dir(m_Base) <> "" Then
            Name m_Base As pName
        End If
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub OpenFolderDialog()
    On Error GoTo ErrorHandler:
    Dim p_Dialog As FileDialog
    Dim p_Result As Integer
    Set p_Dialog = Application.FileDialog(msoFileDialogFolderPicker)
    p_Dialog.AllowMultiSelect = False
    p_Dialog.Title = "Select an Folder"
    p_Dialog.InitialFileName = "C:\"
    p_Dialog.Filters.Clear
    If p_Dialog.Show = -1 Then
       Debug.Print p_Dialog.SelectedItems(1)
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            Zip
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Zip(pFolderPath As Variant, pZippedPath As Variant)
    On Error GoTo ErrorHandler:
    Dim p_Shell As Object
    Open pZippedPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    Set p_Shell = CreateObject("Shell.Application")
    p_Shell.Namespace(pZippedPath).CopyHere p_Shell.Namespace(pFolderPath).Items
    Do Until p_Shell.Namespace(pZippedPath).Items.count = p_Shell.Namespace(pFolderPath).Items.count
        p_Shell.Wait (Now + TimeValue("0:00:01"))
    Loop
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



