Attribute VB_Name = "modSound"
Option Explicit

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module completely redone. Now using DirectSound and DirectMusic ;) - GIAKEN'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public CurrentSong As String

' DirectSound7
Public Ds As DirectSound
Public dsbuffer As DirectSoundBuffer
Public DsDesc As DSBUFFERDESC
Public DsWave As WAVEFORMATEX

' DirectMusic7
Public perf As DirectMusicPerformance
Public seg As DirectMusicSegment
Public segstate As DirectMusicSegmentState
Public loader As DirectMusicLoader

Public Sub PlayMidi(ByRef Song As String)
On Error GoTo ErrHandler

If Song = vbNullString Or Song = "None" Then
    Call StopMidi
    Exit Sub
End If

If MusicOn = YES Then

    If FileExist(App.Path & "\Music\" & Song) = False Then Exit Sub

    If CurrentSong <> Song Then
        CurrentSong = Song
        
        If Not (seg Is Nothing) Then Set seg = Nothing
        Set seg = loader.LoadSegment(App.Path & "\Music\" & Song)
        seg.SetStandardMidiFile
        Call perf.PlaySegment(seg, 0, 0)
        
    End If
Else
    Call StopMidi
End If

ErrHandler:
    Exit Sub
End Sub

Public Sub StopMidi()
Dim I As Long

    CurrentSong = vbNullString
    If perf Is Nothing Then Exit Sub
    Call perf.Stop(seg, segstate, 0, 0)
End Sub

Public Sub MakeMidiLoop()

    If seg Is Nothing Then Exit Sub
    If perf Is Nothing Then Exit Sub
    If perf.IsPlaying(seg, segstate) = False And CurrentSong <> vbNullString Then
        Set segstate = perf.PlaySegment(seg, 0, 0)
    End If

End Sub

Public Sub PlaySound(ByRef Sound As String)
On Error GoTo ErrHandler

    If SoundOn = YES Then
        If FileExist("SFX\" & Sound) = False Then Exit Sub
        If Not (dsbuffer Is Nothing) Then Set dsbuffer = Nothing
        Set dsbuffer = Ds.CreateSoundBufferFromFile(App.Path & "\SFX\" & Sound, DsDesc, DsWave)
        dsbuffer.Play DSBPLAY_DEFAULT
    End If
    Exit Sub

ErrHandler:
    Exit Sub
End Sub

Public Sub StopSound()
    
    dsbuffer.Stop
    dsbuffer.SetCurrentPosition 0
    
End Sub
