Attribute VB_Name = "modSound"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public CurrentSong As String
Option Explicit


Public Sub PlayMidi(Song As String)
Dim i As Long

If Val(GetVar(App.Path & "\config.ini", "CONFIG", "Music")) = 1 Then
            If CurrentSong <> Song Then
                Call StopMidi
                CurrentSong = Song
                Call MP3.OpenMP3(App.Path & "\music\" & Song)
                Call MP3.PlayMP3
            End If
Else
    Call StopMidi
End If

End Sub
Public Sub StopMidi()

        CurrentSong = vbNullString
        Call MP3.StopMP3
        Call MP3.CloseMP3
End Sub

Public Sub MakeMidiLoop()
Dim SBuffer As String * 256

If Right$(CurrentSong, 4) = ".mid" Then
Call mciSendString("STATUS background MODE", SBuffer, 256, 0)

If Left$(SBuffer, 7) = "stopped" Then
    Call mciSendString("PLAY background FROM 0", vbNullString, 0, 0)
End If
End If
End Sub

Public Sub PlaySound(Sound As String)
    If Val(GetVar(App.Path & "\config.ini", "CONFIG", "Sound")) = 1 Then
        If FileExist("SFX\" & Sound) = False Then Exit Sub
        Call sndPlaySound(App.Path & "\SFX\" & Sound, SND_ASYNC Or SND_NODEFAULT)
    End If
End Sub

Public Sub StopSound()
    Dim x As Long
    Dim wFlags As Long

    wFlags = SND_ASYNC Or SND_NODEFAULT
    x = sndPlaySound("", wFlags)
End Sub


