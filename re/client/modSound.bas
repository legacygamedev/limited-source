Attribute VB_Name = "modSound"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Sub PlayMidi(Song As String)
Dim I As Long

If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
            If CurrentSong <> Song Then
                Call StopMidi
                CurrentSong = Song
                If Not Right$(Song, 4) = ".mid" Then
                     Call PlayMP3(Song, FSOUND_LOOP_NORMAL)
                End If
            End If
    If Right$(Song, 4) = ".mid" Then
                I = mciSendString("close all", 0, 0, 0)
                I = mciSendString("open """ & App.Path & "\Music\" & Song & """ Type sequencer Alias background", 0, 0, 0)
                I = mciSendString("play background notify", 0, 0, frmMirage.hwnd)
    End If
           
Else
    Call StopMidi
End If

End Sub
Public Sub StopMidi()
Dim I As Long

    If Right$(CurrentSong, 4) = ".mid" Then
        I = mciSendString("close all", 0, 0, 0)
    End If
        CurrentSong = ""
        Call StopMP3
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
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
        If FileExist("SFX\" & Sound) = False Then Exit Sub
        Call sndPlaySound(App.Path & "\SFX\" & Sound, SND_ASYNC Or SND_NODEFAULT)
    End If
End Sub

Public Sub StopSound()
    Dim X As Long
    Dim wFlags As Long

    wFlags = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound("", wFlags)
End Sub

Public Sub PlayMP3(Sound As String, ModePlay As String)
Dim result As Boolean

    If FSOUND_IsPlaying(sampleChannel) Then
        Call StopMP3
    End If
    
    'Verify that file exists
    If FileExist("Music\" & Sound) = False Then Exit Sub
   
    'Load file In RAM
    sampleHandle = FSOUND_Sample_Load(FSOUND_FREE, App.Path & "\Music\" & Sound, ModePlay, 0, 0)
       
    'Play File
    sampleChannel = FSOUND_PlaySound(FSOUND_FREE, sampleHandle)
End Sub

Public Sub StopMP3()
FSOUND_StopSound sampleChannel
sampleChannel = 0
FSOUND_Sample_Free sampleHandle
sampleHandle = 0
End Sub
