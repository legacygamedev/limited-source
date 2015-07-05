Attribute VB_Name = "modSound"
Option Explicit

Public Sub PlayMidi(Song As String)
Dim i As Long

    If CurrentSong <> Song Then
        CurrentSong = Song
        i = mciSendString("close all", 0, 0, 0)
        i = mciSendString("open """ & App.Path & "\Core files\Music\" & Song & """ type sequencer alias background", 0, 0, 0)
        i = mciSendString("play background notify", 0, 0, frmMainGame.hwnd)
    End If
End Sub

Public Sub StopMidi()
Dim i As Long

    CurrentSong = vbNullString
    i = mciSendString("close all", 0, 0, 0)
End Sub
Public Sub MakeMidiLoop()
Dim SBuffer As String * 256

    Call mciSendString("STATUS background MODE", SBuffer, 256, 0)
    
    If Left$(SBuffer, 7) = "stopped" Then
        Call mciSendString$("PLAY background FROM 0", vbNullString, 0, 0)
    End If
End Sub

Public Sub PlaySound(Sound As String)
    Call sndPlaySound(App.Path & "\" & Sound, SND_ASYNC Or SND_NOSTOP)
End Sub



