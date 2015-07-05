Attribute VB_Name = "modSound"
Option Explicit

Public Sub MapMusic(ByVal Song As String)
    If Not Map(GetPlayerMap(MyIndex)).music = CurrentSong Then
        Call PlayBGM(Map(GetPlayerMap(MyIndex)).music)
    End If
End Sub

Public Sub PlayBGM(Song As String)
    If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
        If FileExists("\Music\" & Song) Then
            If Not LenB(Song) = 0 Then
                If Not Left$(Song, 7) = "http://" Then
                    Call frmMirage.MusicPlayer.PlayMedia(App.Path & "\Music\" & Song, True)
                    CurrentSong = Song
                Else
                    Call frmMirage.MusicPlayer.PlayMedia(Song, True)
                    CurrentSong = Song
                End If
            End If
        Else
            Call AddText(Song & " does not exist!", BRIGHTRED)
        End If
    End If
End Sub

Public Sub PlaySound(Sound As String)
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
        If FileExists("\SFX\" & Sound) Then
            Call frmMirage.SoundPlayer.PlayMedia(App.Path & "\SFX\" & Sound, False)
        Else
            Call AddText(Sound & " does not exist!", BRIGHTRED)
        End If
    End If
End Sub

Public Sub PlayBGS(Sound As String)
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
        If FileExists("\SFX\" & Sound) Then
            Call frmMirage.BGSPlayer.PlayMedia(App.Path & "\BGS\" & Sound, True)
        Else
            Call AddText(Sound & " does not exist!", BRIGHTRED)
        End If
    End If
End Sub

Public Sub StopBGM()
    Call frmMainMenu.MenuMusic.StopMedia
    Call frmMirage.MusicPlayer.StopMedia
    Call frmMirage.BGSPlayer.StopMedia
    CurrentSong = vbNullString
End Sub

Public Sub StopSound()
    Call frmMirage.SoundPlayer.StopMedia
End Sub
