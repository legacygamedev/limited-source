Attribute VB_Name = "modSound"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const END_CHAR_SYNC = &H0
Public Const END_CHAR_ASYNC = &H1
Public Const END_CHAR_NODEFAULT = &H2
Public Const END_CHAR_MEMORY = &H4
Public Const END_CHAR_LOOP = &H8
Public Const END_CHAR_NOSTOP = &H10
Public CurrentSong As String

Public Sub PlayMidi(Song As String)
Dim i As Long

If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
    If CurrentSong <> Song Then
        CurrentSong = Song
        i = mciSendString("close all", 0, 0, 0)
        i = mciSendString("open """ & App.Path & "\Music\" & Song & """ type sequencer alias background", 0, 0, 0)
        i = mciSendString("play background notify", 0, 0, frmEndieko.hWnd)
    End If
Else
    Call StopMidi
End If
End Sub

Public Sub StopMidi()
Dim i As Long

    CurrentSong = ""
    i = mciSendString("close all", 0, 0, 0)
End Sub

Public Sub MakeMidiLoop()
Dim SBuffer As String * 256

Call mciSendString("STATUS background MODE", SBuffer, 256, 0)

If Left$(SBuffer, 7) = "stopped" Then
    Call mciSendString("PLAY background FROM 0", vbNullString, 0, 0)
End If
End Sub

Public Sub PlaySound(Sound As String)
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
        If FileExist("Sounds\" & Sound) = False Then Exit Sub
        Call sndPlaySound(App.Path & "\Sounds\" & Sound, END_CHAR_ASYNC Or END_CHAR_NODEFAULT)
    End If
End Sub

Public Sub StopSound()
    Dim X As Long
    Dim wFlags As Long

    wFlags = END_CHAR_ASYNC Or END_CHAR_NODEFAULT
    X = sndPlaySound("", wFlags)
End Sub



