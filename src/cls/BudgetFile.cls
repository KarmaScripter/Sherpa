VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BudgetFile"
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
Private m_File As Scripting.File
Private m_CurrentDirectory As String
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
Public Property Let Base(fp As String)
    On Error GoTo ErrorHandler:
    If Not fp & "" = "" Then
        m_Base = fp
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Get
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Base() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
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
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ParentFolder() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        m_ParentFolder = m_FileSys.GetFile(m_Base).ParentFolder
        ParentFolder = m_ParentFolder
        Set m_FileSys = Nothing
    End If
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
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        m_Path = m_FileSys.GetAbsolutePathName(m_Base)
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
Public Property Get DateCreated() As Date
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        m_DateCreated = m_FileSys.GetFile(m_Base).DateCreated
        DateCreated = m_DateCreated
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
Public Property Get LastAccessed() As Date
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        m_LastAccessed = m_FileSys.GetFile(m_Base).DateLastAccessed
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
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        m_LastModified = m_FileSys.GetFile(m_Base).DateLastModified
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
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        m_Name = m_FileSys.GetFile(m_Base).ShortName
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
Public Property Get FileType() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        m_Type = m_FileSys.GetFile(m_Base).Type
        FileType = m_Type
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
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        Size = m_FileSys.GetFile(m_Base).Size
        Set m_FileSys = Nothing
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
Public Property Get Extension() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        m_Extension = m_FileSys.GetExtensionName(m_Base)
        Extension = m_Extension
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
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Length() As Long
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Length = FileLen(m_Base)
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
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        m_Drive = m_FileSys.GetFile(m_Base).Drive
        Drive = m_Drive
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetPath
'   Parameters:      String
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetPath(pPath As String) As String
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" And _
        Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
            GetPath = m_FileSys.GetAbsolutePathName(m_Base)
            Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetName
'   Parameters:      String
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetName(pPath As String) As String
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" And _
        Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
            GetName = m_FileSys.GetFileName(pPath)
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
'   Retval:          Integer
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetSize(pPath As String) As Integer
    On Error GoTo ErrorHandler:
    If Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        GetSize = CInt(m_FileSys.GetFile(pPath).Size)
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetExtension
'   Parameters:      String
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetExtension(pPath As String) As String
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" And _
        Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        GetExtension = m_FileSys.GetExtensionName(pPath)
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetCreationDate
'   Parameters:      String
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetCreationDate(pPath As String) As Date
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" And _
        Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        GetCreationDate = m_FileSys.GetFile(pPath).DateCreated
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function





'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetCreationDate
'   Parameters:      String
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetParentFolder(pPath As String) As Scripting.folder
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" And _
        Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        Set GetParentFolder = m_FileSys.GetFile(pPath).ParentFolder
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetCreationDate
'   Parameters:      String
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetAccessDate(pPath As String) As Date
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" And _
        Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        GetAccessDate = m_FileSys.GetFile(pPath).DateLastAccessed
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetCreationDate
'   Parameters:      String
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetModifiedDate(pPath As String) As Date
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" And _
        Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        GetModifiedDate = m_FileSys.GetFile(pPath).DateLastModified
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetDrive
'   Parameters:      String
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetDrive(pPath As String) As Scripting.Drive
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" And _
        Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        Set GetDrive = m_FileSys.GetFile(pPath).Drive
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function





'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetDrive
'   Parameters:      String
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetType(pPath As String) As String
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" And _
        Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        GetType = m_FileSys.GetFile(pPath).Type
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function





'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:            FileExists
'   Parameters:      String
'   Retval:          Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Function Verify(pPath As String) As Boolean
    On Error GoTo ErrorHandler:
    If Not Dir(pPath) & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        If m_FileSys.FileExists(pPath) = True Then
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
'   Name:            GetStreamReader
'   Parameters:      String
'   Retval:          TextStream
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetStreamReader() As Scripting.TextStream
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        Set m_File = m_FileSys.GetFile(m_Base)
        Set GetStreamReader = m_File.OpenAsTextStream(IOMode:=ForReading, _
            Format:=TristateUseDefault)
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetStreamReader
'   Parameters:      String
'   Retval:          TextStream
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetStreamWriter() As Scripting.TextStream
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
            Set m_File = m_FileSys.GetFile(m_Base)
            Set GetStreamReader = m_File.OpenAsTextStream(IOMode:=ForWriting, _
                Format:=TristateUseDefault)
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            Create()
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function CreateText(pFileName As String) As Scripting.TextStream
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not pFileName & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
            Set CreateText = m_FileSys.CreateTextFile(filename:=pFileName, _
                Overwrite:=True)
            Set m_FileSys = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            ReadAll
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function ReadAll() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        Set m_File = m_FileSys.GetFile(m_Base)
        Dim m_Out As String
        Dim m_Stream As Scripting.TextStream
        Set m_Stream = m_File.OpenAsTextStream
        m_Out = m_Stream.ReadAll
        Set m_FileSys = Nothing
        Set m_File = Nothing
        Set m_Stream = Nothing
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
Public Function Readline() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        Set m_File = m_FileSys.GetFile(m_Base)
        If Not Dir(m_File) & "" = "" Then
            Dim p_Out As String
            Dim p_Stream As Scripting.TextStream
            Set p_Stream = m_File.OpenAsTextStream
            p_Out = p_Stream.Readline
            Set m_FileSys = Nothing
            Set m_File = Nothing
            Set p_Stream = Nothing
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
    If Not m_Path & "" = "" And _
        Not des & "" = "" Then
        Set m_FileSys = New Scripting.FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystem")
        Set m_File = m_FileSys.GetFile(m_Base)
        m_File.Move (des)
        Set m_FileSys = Nothing
        Set m_File = Nothing
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
Public Sub Copy(pSource As String, pDestination As String)
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not Dir(pSource) & "" = "" And _
        Not Dir(pDestination) & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            Set m_File = m_FileSys.GetFile(pSource)
            m_File.Copy Destination:=pDestination, OverWriteFiles:=True
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
Public Sub Delete(pPath As String)
    On Error GoTo ErrorHandler:
    Set m_FileSys = New FileSystemObject
    Set m_FileSys = CreateObject("Scripting.FileSystemObject")
    If Not Dir(pPath) & "" = "" Then
        m_FileSys.DeleteFile pPath
    End If
    Set m_FileSys = Nothing
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            FileRename
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Rename(pName As String)
    On Error GoTo ErrorHandler:
    If Not m_Path & "" = "" And _
        Not pName & "" = "" Then
            Name m_Path As pName
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RecursiveDelete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub RecursiveDelete(pFilePath As String, pFileName As String)
    On Error GoTo ErrorHandler:
    Dim p_CurrentDirectory As String
    Dim p_CurrentPath As String
    Dim p_File As Variant
    Dim p_Folders As Collection
    Set p_Folders = New Collection
    If Len(Dir(pFilePath & pFileName)) > 0 Then
       SetAttr pFilePath & pFileName, vbNormal
       Kill pFilePath & pFileName
    End If
    p_CurrentDirectory = Dir(pFilePath, vbDirectory)
    Do Until p_CurrentDirectory = vbNullString
        If Left(p_CurrentDirectory, 1) <> "." _
            And (GetAttr(pFilePath & p_CurrentDirectory) _
            And vbDirectory) = vbDirectory Then
                p_Folders.add pFilePath & p_CurrentDirectory & "\"
        End If
        p_CurrentDirectory = Dir()
    Loop
    For Each p_File In p_Folders
        RecursiveDelete CStr(p_File), pFileName
    Next p_File
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
Public Sub Browse()
    On Error GoTo ErrorHandler:
    Dim p_Dialog As FileDialog
    Dim result As Integer
    Set p_Dialog = Application.FileDialog(msoFileDialogFilePicker)
    p_Dialog.AllowMultiSelect = False
    p_Dialog.Title = "Select a file"
    p_Dialog.InitialFileName = "C:\"
    p_Dialog.Filters.Clear
    p_Dialog.Filters.add "All files", "*.*"
    If p_Dialog.Show = -1 Then
       Debug.Print p_Dialog.SelectedItems(1)
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
Public Function UnZip(pSource As String, pDestination As String) As String
    On Error GoTo ErrorHandler:
    Dim p_FailedReason As String
    Dim p_ZipToolPath As String
    Dim p_ShellCmd As String
    Dim p_Success As Boolean
    p_ZipToolPath = CurrentProject.Path & "\etc\7za"
    p_ShellCmd = p_ZipToolPath & " x " & pSource & " -o" & pDestination & " -ry"
    p_Success = Shell(p_ShellCmd, vbHide)
    If p_Success = True Then
        Kill pSource
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetCurrentDirectory()
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Function GetCurrentDirectory() As String
    On Error GoTo ErrorHandler:
    m_CurrentDirectory = CurDir()
    GetCurrentDirectory = m_CurrentDirectory
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
Public Sub OpenSaveDialog()
    On Error GoTo ErrorHandler:
    Dim p_Dialog As FileDialog
    Dim result As Integer
    Set p_Dialog = Application.FileDialog(msoFileDialogSaveAs)
    p_Dialog.AllowMultiSelect = False
    p_Dialog.Title = "Select location to save file."
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
'   Name:            GetAttribute
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetAttribute() As VbFileAttribute
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" Then
        Dim p_Attribute As VbFileAttribute
        p_Attribute = GetAttr(m_Base)
        GetAttribute = p_Attribute
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        KillFiles
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub KillFiles()
    If Not m_FileSys Is Nothing Then
        Set m_FileSys = Nothing
    End If
    If Not m_File Is Nothing Then
        Set m_File = Nothing
    End If
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
    End If
    MessageFactory.ShowError (m_Error)
    Err.Clear
End Sub



