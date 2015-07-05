Attribute VB_Name = "modSystem"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

' modINI
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

Public Sub WriteINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Public Function ReadINI(Section As String, KeyName As String, FileName As String, Default As String) As String
    Dim sRet As String

    sRet = String$(255, Chr$(0))

    ReadINI = Left$(sRet, GetPrivateProfileString(Section, ByVal KeyName, Default, sRet, Len(sRet), FileName))
End Function

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.text = Txt.text & vbNewLine & Msg
    Else
        Txt.text = Txt.text & Msg
    End If

    MAX_SERVLINES = MAX_SERVLINES + 1
    If MAX_SERVLINES >= MAX_LINES Then
        Txt.text = vbNullString
        MAX_SERVLINES = 0
    End If

    Txt.SelStart = Len(Txt.text)
End Sub
