Attribute VB_Name = "modMusic"
Option Explicit

' Hardcoded sound effects
Public Const Sound_ButtonHover As String = "Cursor1.wav"
Public Const Sound_ButtonClick As String = "Decision1.wav"

Public lastButtonSound As Long
Public lastNpcChatsound As Long

Public bInit_Music As Boolean
Public bInit_Sound As Boolean
Public curSong As String

Private songHandle As Long
Private streamHandle As Long

Public Function Init_Music() As Boolean
Dim result As Boolean

    On Error GoTo errorhandler
    
    ' init music engine
    result = FSOUND_Init(44100, 32, FSOUND_INIT_USEDEFAULTMIDISYNTH)
    If Not result Then GoTo errorhandler
    
    ' return positive
    Init_Music = True
    bInit_Music = True
    Exit Function
    
errorhandler:
    Init_Music = False
    bInit_Music = False
End Function

Public Sub Destroy_Music()
    ' destroy music engine
    Stop_Music
    FSOUND_Close
    bInit_Music = False
    curSong = vbNullString
End Sub

Public Sub Play_Music(ByVal song As String)
    If Not bInit_Music Then Exit Sub
    
    ' exit out early if we have the system turned off
    If Options.Music = 0 Then Exit Sub
    
    ' does it exist?
    If Not FileExist(App.path & MUSIC_PATH & song) Then Exit Sub
    
    ' don't re-start currently playing songs
    If curSong = song Then Exit Sub
    
    ' stop the existing music
    Stop_Music
    
    ' find the extension
    Select Case Right$(song, 4)
        Case ".mid", ".s3m", ".mod"
            ' open the song
            songHandle = FMUSIC_LoadSong(App.path & MUSIC_PATH & song)
            ' play it
            FMUSIC_PlaySong songHandle
            ' set volume
            FMUSIC_SetMasterVolume songHandle, 150
            
        Case ".wav", ".mp3", ".ogg", ".wma"
            ' open the stream
            streamHandle = FSOUND_Stream_Open(App.path & MUSIC_PATH & song, FSOUND_LOOP_NORMAL, 0, 0)
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

Public Sub Stop_Music()
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

Public Sub Play_Sound(ByVal sound As String)
    If Not bInit_Sound Then Exit Sub
    
    ' exit out early if we have the system turned off
    If Options.sound = 0 Then Exit Sub
End Sub
