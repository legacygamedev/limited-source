Attribute VB_Name = "modBatchProcess"
Option Explicit

Public RegFiles() As Boolean

Public Sub MakeRegBatch(ByVal FileName As String, ByVal FileExt As String, ByVal FilePath As String)
Dim Path As String
Dim f As Integer

Path = FilePath & FileName & ".bat"
f = FreeFile

Debug.Print Path
Open Path For Output As #f
    Print #f, "@ECHO OFF"
    Print #f, "ECHO :: Registering " & FileName & "." & FileExt & " for the RealFeel Engine!"
    Print #f, "COPY " & FilePath & FileName & "." & FileExt & " C:\Windows\System32 /Y"
    Print #f, "REGSVR32 " & FilePath & FileName & "." & FileExt & " / s"
    Print #f, "ECHO Registered Successfully!"
Close #f
End Sub

Public Sub RunRegBatch(ByVal FileName As String, ByVal FilePath As String, Optional KillIt As Boolean = False)
    Call Shell(FilePath & FileName & ".bat")
    If KillIt = True Then Call Kill(FilePath & FileName & ".bat")
End Sub

Public Function Exists(ByVal Path As String) As Boolean
Exists = False
If Dir(Path) <> "" Then Exists = True
End Function

Public Function CheckReg(ByVal Path As String)
Dim FSys As Object, Folder As Object, FolderFiles As Object, File As Object, FileName As String
Dim n As Byte
    ' Preset the return value
    CheckReg = True

    ' Create the file
    Set FSys = CreateObject("Scripting.FileSystemObject")

    'Set the folder objects
    Set Folder = FSys.GetFolder(Path)
    Set FolderFiles = Folder.Files

    n = 1
    For Each File In FolderFiles
        FileName = Mid(File, Len(Path) + 1, (Len(File) - Len(Path)))
        ReDim Preserve RegFiles(1 To n) As Boolean
        RegFiles(n) = False
        If Not Exists("C:\Windows\System32\" & FileName) Then
            If UCase$(Right$(FileName, 3)) = "OCX" Or UCase$(Right$(FileName, 3)) = "DLL" Then
                frmRegFiles.lstRegFiles.AddItem FileName
                CheckReg = False
                RegFiles(n) = False
            End If
        Else
            If UCase$(Right$(FileName, 3)) = "OCX" Or UCase$(Right$(FileName, 3)) = "DLL" Then
                frmRegFiles.lstRegFiles.AddItem FileName
                CheckReg = False
                RegFiles(n) = True
            End If
        End If
        n = n + 1
    Next File
    
    'Destroy the folder objects
    Set File = Nothing
    Set FolderFiles = Nothing
    Set Folder = Nothing
    Set FSys = Nothing
End Function
