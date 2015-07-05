Attribute VB_Name = "modSound"
Public CurrentSong As String
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Sub PlayMidi(Song As String)
If CurrentSong <> Song Then
CurrentSong = Song

Dim i As Long

i = mciSendString("close all", 0, 0, 0)
i = mciSendString("open """ & App.Path & "\Core files\Music\" & Song & """ type sequencer alias background", 0, 0, 0)
i = mciSendString("play background notify", 0, 0, frmMainGame.hWnd)
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
    Call sndPlaySound(App.Path & "\" & Sound, SND_ASYNC Or SND_NOSTOP)
End Sub



