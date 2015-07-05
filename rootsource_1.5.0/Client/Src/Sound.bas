Attribute VB_Name = "Sound"
Option Explicit

Public CurrentMusic As String
Private SoundManager As FMOD
Private SoundCollection As Collection

Public Sub InitSoundSys()
Set SoundManager = New FMOD
    Call SoundManager.InitFMOD
End Sub

Public Function PlayMusic(Filename) As Long
    Call SoundManager.PlayMusic(App.Path & MUSIC_PATH & Filename)
End Function

Public Function StopMusic()
    Call SoundManager.StopMusic
End Function

Public Sub KillSoundSys()
Set SoundManager = Nothing
End Sub

