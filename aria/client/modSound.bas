Attribute VB_Name = "modSound"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public CurrentSong As String

'DirectSound Shit
Public DirectSoundLoaded As Boolean
Public Const cSoundBuffers = 15

Public Sub PlayMidi(Song As String)
Dim i As Long

If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
    If CurrentSong <> Song Then
        CurrentSong = Song
        i = mciSendString("close all", 0, 0, 0)
        i = mciSendString("open """ & App.Path & "\Music\" & Song & """ type sequencer alias background", 0, 0, 0)
        i = mciSendString("play background notify", 0, 0, frmMirage.hwnd)
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
    If DirectSoundLoaded = False Then Exit Sub
    
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
        If FileExist("SFX\" & Sound) = False Then Exit Sub
        Call PlaySoundAnyBuffer(Sound, 65, 50)
        'Call sndPlaySound(App.Path & "\SFX\" & Sound, SND_ASYNC Or SND_NODEFAULT)
    End If
End Sub

Public Sub StopSound()
    Dim X As Long
    Dim wFlags As Long

    wFlags = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound("", wFlags)
End Sub




