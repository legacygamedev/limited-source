Attribute VB_Name = "modSound"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public MusicPlaying As String
Public MuteMusic As Boolean
Public MuteSound As Boolean

Dim SoundName$
Dim wFlags%
Dim x%

Dim Samplehandle As Long
Dim Streamhandle As Long
Dim SampleChannel As Long
Dim StreamChannel As Long

Public loadedSounds As Boolean




Public Sub PlayMidi(Song As String, Optional loopStart As Boolean = False, Optional fromMute As Boolean = False)

If MuteMusic = False Then
    If (MP3.MP3File <> App.Path & "\data\audio\music\" & Song & ".mid") Or loopStart = True Or fromMute = True Then
        MusicPlaying = Song
        MP3.MP3Stop
        DoEvents
        MP3.MP3File = App.Path & "\data\audio\music\" & Song & ".mid"
        DoEvents
        MP3.MP3Play
        DoEvents
        frmMirage.timMusic.Enabled = True
    End If
End If


End Sub

Public Sub PlayMP3(Song As String)
If MuteMusic = False Then
    MP3.MP3Stop
    MP3.MP3File = App.Path & "\data\audio\music\" & Song & ".mp3"
    MP3.MP3Play
End If
End Sub

Public Sub StopMidi()
MP3.MP3Stop
frmMirage.timMusic.Enabled = False
End Sub

Public Sub PlaySound(Sound As String)
If MuteSound = False Then
    'SoundName$ = App.Path & "\data\audio\sound\" & Sound & "_PCM.wav"
    'wFlags% = SND_ASYNC Or SND_NODEFAULT
    'x% = sndPlaySound(SoundName$, wFlags%)
    MP31.MP3Stop
    MP31.MP3File = App.Path & "\data\audio\sound\" & Sound & "_PCM.wav"
    MP31.MP3Play
End If
End Sub
