Attribute VB_Name = "modRegister"
Option Explicit

'Files
Public Const MAX_REG As Byte = 13

Type FileRec
    Name As String
End Type

Public File(1 To MAX_REG) As FileRec

'Constants
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const MAX_MESSAGE_LENGTH = 512

'API declarations
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
ByVal dwFlags As Long, _
lpSource As Any, _
ByVal dwMessageId As Long, _
ByVal dwLanguageId As Long, _
ByVal lpBuffer As String, _
ByVal nSize As Long, _
Arguments As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Function IsDLLAvailable(ByVal DllFilename As String) As Boolean
    Dim hModule As Long

    hModule = LoadLibrary(DllFilename) 'attempt to load DLL
    If hModule > 32 Then
        FreeLibrary hModule 'decrement the DLL usage counter
        IsDLLAvailable = True 'Return true
    Else
        IsDLLAvailable = False 'Return False
    End If
End Function

Sub Register(ByVal Filename As String)
Dim File As String

File = App.Path & "\DLL\" & Filename

    If IsDLLAvailable(File) = False Then
        Call Shell("C:\WINDOWS\system32\regsvr32.exe /s " & """" & File & """" & "", vbHide)
    End If

End Sub
