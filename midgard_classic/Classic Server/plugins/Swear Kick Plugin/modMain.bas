Attribute VB_Name = "modMain"
Option Explicit

Public Codes As Object

Public Const DirUp = 0
Public Const DirDown = 1
Public Const DirLeft = 2
Public Const DirRight = 3

'Public Const LocalSocket = 1
Public Const Username = "SwearBot"
Public Const Password = "botpass"
Public Const CharIndex = 1

Public Const BotMap = 1
Public Const BotX = 4
Public Const BotY = 3
Public Const BotDir = DirDown

Public BotIndex As Integer
Public LocalSocket As Integer
Public PluginNum As Integer

Public Function AddKicked(Person As String)
    frmMain.lstKicked.AddItem Person
End Function

Public Function CheckMessage(Who As String, Msg As String)
Dim kArray() As String
Dim Swears As String
Dim i As Integer
    Swears = "ass,fuck,shit,bitch,damn"
    kArray = Split(Swears, ",")
    For i = 0 To UBound(kArray)
        If InStr(Msg, kArray(i)) Then
            Call KickPlayer(Who)
            Call AddKicked(Who)
        End If
    Next i
End Function
