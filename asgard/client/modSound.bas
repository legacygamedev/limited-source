Attribute VB_Name = "modSound"
'   Copyright (c) 2006 Joshua Bendig
'   This file is part of Asgard.
'
'    Asgard is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Asgard is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Asgard; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public CurrentSong As String

Option Explicit

'Public declarations of variables holding songs, samples and streams
Dim songHandle As Long
Dim sampleHandle As Long
Dim sampleChannel As Long
Dim streamHandle As Long
Dim streamChannel As Long

Public Sub PlayMidi(Song As String)
Dim i As Long

If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
            If CurrentSong <> Song Then
                Call StopMidi
                CurrentSong = Song
                If Not Right$(Song, 4) = ".mid" Then
                     Call PlayMP3(Song, FSOUND_LOOP_NORMAL)
                End If
            End If
    If Right$(Song, 4) = ".mid" Then
                i = mciSendString("close all", 0, 0, 0)
                i = mciSendString("open """ & App.Path & "\Music\" & Song & """ Type sequencer Alias background", 0, 0, 0)
                i = mciSendString("play background notify", 0, 0, frmMirage.hwnd)
    End If
           
Else
    Call StopMidi
End If

End Sub

Public Sub StopMidi()
Dim i As Long

    If Right$(CurrentSong, 4) = ".mid" Then
        CurrentSong = ""
        i = mciSendString("close all", 0, 0, 0)

    Else
        CurrentSong = ""
        Call StopMP3
    End If
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
    Dim x As Long
    Dim wFlags As Long

    wFlags = SND_ASYNC Or SND_NODEFAULT
    x = sndPlaySound("", wFlags)
End Sub

Public Sub PlayMP3(Sound As String, ModePlay As String)
Dim result As Boolean

    'Verify that file exists
    If FileExist("Music\" & Sound) = False Then Exit Sub
    
    'Load file in RAM
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
