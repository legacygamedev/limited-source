Attribute VB_Name = "modSound"
Option Explicit

' Hardcoded sound effects
Public Const Sound_ButtonHover As String = "Cursor1.wav"
Public Const Sound_ButtonClick As String = "Decision1.wav"
Public Const Sound_Thunder As String = "Thunder.wav"

Public INIT_MUSIC As Boolean
Public INIT_SOUND As Boolean

Public lastButtonSound As Long
Public lastNpcChatsound As Long

Private Type SoundRec
    Handle As Long
    x As Long
    y As Long
    Loop As Long
    Channel As Long
End Type

Public CurrentSound As Long
Public Const Max_Sounds As Long = 30
Public Sounds(1 To Max_Sounds) As SoundRec

Public curSong As String
Private songHandle As Long
Private streamHandle As Long
Public ResetSounds As Boolean


Public Function InitFmod() As Boolean
Dim result As Boolean

    On Error GoTo errorhandler
    
    If App.LogMode = 0 Then Exit Function
    
    ' init music engine
    result = FSOUND_Init(44100, 32, FSOUND_INIT_USEDEFAULTMIDISYNTH)
    If Not result Then GoTo errorhandler
    
    ' return positive
    InitFmod = True
    INIT_MUSIC = True
    INIT_SOUND = True
    ReDim MapSounds(0)
    Exit Function
    
errorhandler:
    InitFmod = False
    INIT_MUSIC = False
    INIT_SOUND = False
End Function

Public Sub DestroyFmod()
    If INIT_MUSIC = False And INIT_SOUND = False Then Exit Sub
    ' destroy music engine
    StopMusic
    FSOUND_Close
    INIT_MUSIC = False
    curSong = vbNullString
End Sub

Public Sub PlayMusic(ByVal song As String)
    If Not INIT_MUSIC Then Exit Sub
    
    ' exit out early if we have the system turned off
    If Options.Music = 0 Then Exit Sub
    
    ' does it exist?
    If Not FileExist(App.Path & MUSIC_PATH & song, True) Then Exit Sub
    
    ' don't re-start currently playing songs
    If curSong = song Then Exit Sub
    
    ' stop the existing music
    StopMusic
    
    ' find the extension
    Select Case Right$(song, 4)
        Case ".mid", ".s3m", ".mod"
            ' open the song
            songHandle = FMUSIC_LoadSong(App.Path & MUSIC_PATH & song)
            ' play it
            FMUSIC_PlaySong songHandle
            ' set volume
            FMUSIC_SetMasterVolume songHandle, 150
            
        Case ".wav", ".mp3", ".ogg", ".wma"
            ' open the stream
            streamHandle = FSOUND_Stream_Open(App.Path & MUSIC_PATH & song, FSOUND_LOOP_NORMAL, 0, 0)
            ' play it
            FSOUND_Stream_Play 0, streamHandle
            ' set volume
            FSOUND_SetVolume streamHandle, 150
        Case Else
            Exit Sub
    End Select
    
    ' new current song
    curSong = song
End Sub

Public Sub StopMusic()
    If Not streamHandle = 0 Then
        ' stop stream
        FSOUND_Stream_Stop streamHandle
        ' destroy
        FSOUND_Stream_Close streamHandle
        streamHandle = 0
    End If
    
    If Not songHandle = 0 Then
        ' stop song
        FMUSIC_StopSong songHandle
        ' destroy
        FMUSIC_FreeSong songHandle
        songHandle = 0
    End If
    
    ' no music
    curSong = vbNullString
End Sub

Public Sub StopAllSounds()
    Dim i As Long
    For i = 1 To Max_Sounds
        If Sounds(i).Handle > 0 Then
            FSOUND_StopSound Sounds(i).Channel
            FSOUND_Sample_Free Sounds(i).Handle
        End If
    Next
    
    RemoveAllMapSounds
End Sub

Public Sub PlaySound(ByVal sound As String, x As Long, y As Long)
    Dim volume As Byte
    If Not INIT_SOUND Then Exit Sub
    If Options.sound = 0 Then Exit Sub
    
    If Not FileExist(App.Path & SOUND_PATH & sound, True) Then Exit Sub
    
    ' load the sound
    Sounds(CurrentSound).Handle = LoadSound(sound)
    If Sounds(CurrentSound).Handle > 0 Then
        Sounds(CurrentSound).x = x
        Sounds(CurrentSound).y = y
        volume = CalculateSoundVolume(Sounds(CurrentSound).x, Sounds(CurrentSound).y)
        Sounds(CurrentSound).Channel = FSOUND_PlaySound(FSOUND_FREE, Sounds(CurrentSound).Handle)
        FSOUND_SetVolume Sounds(CurrentSound).Channel, volume
    Else
        CurrentSound = CurrentSound - 1
    End If
End Sub

Public Function LoadSound(ByVal sound As String, Optional MapSound As Boolean = False) As Long
Dim i As Long
    If INIT_SOUND = False Then Exit Function
    If Options.sound = 0 Then Exit Function
    
    If Not MapSound Then
        CurrentSound = CurrentSound + 1
        If CurrentSound > Max_Sounds Then
            FSOUND_Sample_Free Sounds(1).Handle
            CurrentSound = 1
            ResetSounds = True
            For i = 2 To Max_Sounds
                FSOUND_Sample_Free Sounds(i).Handle
                Sounds(i).Handle = 0
            Next
        End If
    End If
    
    If MapSound Then
        LoadSound = FSOUND_Sample_Load(FSOUND_FREE, App.Path & SOUND_PATH & sound, FSOUND_LOOP_NORMAL, 0, 0)
    Else
        LoadSound = FSOUND_Sample_Load(FSOUND_FREE, App.Path & SOUND_PATH & sound, FSOUND_NORMAL, 0, 0)
    End If
End Function

Function CalculateSoundVolume(SoundX As Long, SoundY As Long) As Byte
Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long, distance As Double, volume As Long
    If InGame = False Then CalculateSoundVolume = 150: Exit Function
    If INIT_SOUND = False Then Exit Function
    If Options.sound = 0 Then Exit Function
    
    If SoundX > -1 Or SoundY > -1 Then
        x1 = (Player(MyIndex).x * 32) + Player(MyIndex).xOffset
        y1 = (Player(MyIndex).y * 32) + Player(MyIndex).yOffset
        x2 = (SoundX * 32) + 16
        y2 = (SoundY * 32) + 16
        If ((x1 - x2) ^ 2) + ((y1 - y2) ^ 2) < 0 Then
            distance = Sqr(((x1 - x2) ^ 2) + ((y1 - y2) ^ 2) * -1)
        Else
            distance = Sqr(((x1 - x2) ^ 2) + ((y1 - y2) ^ 2))
        End If
        If distance >= 256 Then
            CalculateSoundVolume = 0
        Else
            CalculateSoundVolume = 150 - ((distance / 256) * 150)
            If CalculateSoundVolume = 0 Then CalculateSoundVolume = 1
        End If
    Else
        CalculateSoundVolume = 150
    End If
    
End Function

Sub StopSound(soundindex As Long)
    FSOUND_StopSound Sounds(soundindex).Channel
End Sub

Sub RemoveAllMapSounds()
Dim i As Long
    If MapSoundCount > 0 Then
        For i = 1 To MapSoundCount
            MapSounds(i).InUse = False
            FSOUND_StopSound MapSounds(i).Channel
            FSOUND_Sample_Free MapSounds(i).SoundHandle
        Next
        MapSoundCount = 0
        ReDim MapSounds(0)
    End If
End Sub

Sub UpdateSounds()
If INIT_SOUND = False Then Exit Sub
If Options.sound = 0 Then Exit Sub
Dim i As Long, x As Long
    For i = 1 To Max_Sounds
        If Sounds(i).Handle <> 0 Then
            Call FSOUND_SetVolume(Sounds(i).Channel, CalculateSoundVolume(Sounds(i).x, Sounds(i).y))
        End If
    Next
    If MapSoundCount > 0 Then
        For i = 1 To MapSoundCount
            FSOUND_SetVolume MapSounds(i).Channel, CalculateSoundVolume(MapSounds(i).x, MapSounds(i).y)
        Next
    End If
End Sub

Sub CacheNewMapSounds()
If INIT_SOUND = False Then Exit Sub
If Options.sound = 0 Then Exit Sub
    Dim i As Long, x As Long, y As Long
    RemoveAllMapSounds
    If Trim$(Map.BGS) <> Null Then
        MapSoundCount = MapSoundCount + 1
        ReDim Preserve MapSounds(MapSoundCount)
        MapSounds(MapSoundCount).SoundHandle = LoadSound(Trim$(Map.BGS), True)
        FSOUND_SetLoopMode MapSounds(MapSoundCount).SoundHandle, FSOUND_LOOP_NORMAL
        MapSounds(MapSoundCount).x = -1
        MapSounds(MapSoundCount).y = -1
        MapSounds(MapSoundCount).Channel = FSOUND_PlaySound(FSOUND_FREE, MapSounds(MapSoundCount).SoundHandle)
        FSOUND_SetVolume MapSounds(MapSoundCount).Channel, CalculateSoundVolume(MapSounds(MapSoundCount).x, MapSounds(MapSoundCount).y)
    End If
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            If Map.Tile(x, y).Type = TILE_TYPE_SOUND Then
                MapSoundCount = MapSoundCount + 1
                ReDim Preserve MapSounds(MapSoundCount)
                MapSounds(MapSoundCount).SoundHandle = LoadSound(Trim$(Map.Tile(x, y).Data4), True)
                FSOUND_SetLoopMode MapSounds(MapSoundCount).SoundHandle, FSOUND_LOOP_NORMAL
                MapSounds(MapSoundCount).x = x
                MapSounds(MapSoundCount).y = y
                MapSounds(MapSoundCount).Channel = FSOUND_PlaySound(FSOUND_FREE, MapSounds(MapSoundCount).SoundHandle)
                FSOUND_SetVolume MapSounds(MapSoundCount).Channel, CalculateSoundVolume(MapSounds(MapSoundCount).x, MapSounds(MapSoundCount).y)
            End If
        Next
    Next

End Sub

