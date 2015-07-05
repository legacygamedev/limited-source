Attribute VB_Name = "modSound"
Option Explicit

Public Sub MapMusic(ByVal Song As String)
    'Call AddText("Map music stored is: " & Map(GetPlayerMap(MyIndex)).Music & " and map music playing is: " & CurrentSong, BRIGHTGREEN)
    If Not Map(GetPlayerMap(MyIndex)).music = CurrentSong Then
        Call PlayBGM(Map(GetPlayerMap(MyIndex)).music)
    End If
End Sub

Public Sub PlayBGM(Song As String)
    Call StopBGM
    If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
        If Not LenB(Song) = 0 Then
            If Not Left$(Song, 7) = "http://" Then
                MapSound = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\Music\" & Song), 0, 0, BASS_SAMPLE_LOOP)
                Call BASS_ChannelPlay(MapSound, BASSFALSE)
                CurrentSong = Song
            Else
                MapSound = BASS_StreamCreateURL(Song, 0, BASS_SAMPLE_LOOP, 0, 0)
                Call BASS_ChannelPlay(MapSound, BASSFALSE)
                CurrentSong = Song
            End If
        Else
            Call AddText(Song & " does not exist!", BRIGHTRED)
        End If
    End If
End Sub

Public Sub PlaySound(Sound As String)
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
        If FileExists("\SFX\" & Sound) Then
            Sounds = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\SFX\" & Sound), 0, 0, 0)
            Call BASS_ChannelPlay(Sounds, BASSFALSE)
        Else
            Call AddText(Sound & " does not exist!", BRIGHTRED)
        End If
    End If
End Sub

Public Sub PlayBGS(Sound As String)
    Call StopBGS
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
        If FileExists("\SFX\" & Sound) Then
            'load soundfile with loop options
            BGSound = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\SFX\" & Sound), 0, 0, BASS_SAMPLE_LOOP)
        
            'play opened file
            Call BASS_ChannelPlay(BGSound, BASSFALSE)
            CurrentSound = Sound
        Else
            Call AddText(App.Path & "\SFX\" & Sound & " does not exist!", BRIGHTRED)
        End If
    End If
End Sub

Public Sub StopBGM()
    If Not MapSound = "" Then
        Call BASS_ChannelStop(MapSound)
        CurrentSong = vbNullString
    End If
End Sub

Public Sub StopSound()
    If Not Sounds = "" Then
        Call BASS_ChannelStop(Sounds)
    End If
End Sub

Public Sub StopBGS()
    If Not BGSound = "" Then
        Call BASS_ChannelStop(BGSound)
        CurrentSound = vbNullString
    End If
End Sub

Public Sub FadeIn()
    Call BASS_ChannelSlideAttribute(MapSound, 0, BASS_ATTRIB_VOL, 1000)
End Sub

Public Sub Fadeout()
    Call BASS_ChannelSlideAttribute(MapSound, BASS_ATTRIB_VOL, 0, 1000)
End Sub
