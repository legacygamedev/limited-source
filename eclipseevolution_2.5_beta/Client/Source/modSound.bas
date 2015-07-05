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
Public CurrentBGS As String
Public MapMusicStarted As Boolean
Public Sub MapMusic(ByVal Song As String)
If Map(GetPlayerMap(MyIndex)).Music = CurrentSong Then
    Exit Sub 'The music is the same so dont do anything
Else
    Call PlayBGM(Map(GetPlayerMap(MyIndex)).Music) 'It is different so change it
End If
End Sub
Public Sub PlayBGM(Song As String)
If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
    'Check if the song is blank
    If Not Song = vbNullString Or Not Song = vbNullString Then
        If Not Left(Song, 7) = "http://" Then
        Call frmMirage.MusicPlayer.PlayMedia(App.Path & "\music\" & Song, True)
        CurrentSong = Song
        Else
        Call frmMirage.MusicPlayer.PlayMedia(Song, True)
        CurrentSong = Song
        End If
    ElseIf ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 0 Then
        Call StopBGM
    End If
End If
End Sub

Public Sub StopBGM()
 Call frmMainMenu.MenuMusic.StopMedia
 Call frmMirage.MusicPlayer.StopMedia
 Call frmMirage.BGSPlayer.StopMedia
CurrentSong = "..."
CurrentBGS = "..."
End Sub

Public Sub PlaySound(Sound As String)
    If FileExist("\SFX\" & Sound) = True Then
        If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
            Call frmMirage.SoundPlayer.PlayMedia(App.Path & "\SFX\" & Sound, False)
        End If
    Else
        Call MsgBox(Sound & " does not exist!")
    End If
End Sub
Public Sub PlayBGS(Sound As String)
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
    Call frmMirage.BGSPlayer.PlayMedia(App.Path & "\BGS\" & Sound, True)
    End If
End Sub

Public Sub StopSound()
Call frmMirage.SoundPlayer.StopMedia
End Sub




