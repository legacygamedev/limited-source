Attribute VB_Name = "modDirectMusic7"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' -- DirectMusic7, plays .MID music files --
' ------------------------------------------

' DirectMusic variables
Public Performance As DirectMusicPerformance ' this controls the music
Public Segment As DirectMusicSegment ' this stores the music in memory
Private Loader As DirectMusicLoader 'parses the data from a file on the hard drive to the area in memory.

Public Sub InitDirectMusic()

    Set Loader = DX7.DirectMusicLoaderCreate
    Set Performance = DX7.DirectMusicPerformanceCreate
    
    Performance.Init Nothing, frmMainGame.hWnd
    Performance.SetPort -1, 80
    
    ' adjust volume 0-100
    Performance.SetMasterVolume 75 * 42 - 3000
    Performance.SetMasterAutoDownload True
    
End Sub

Public Sub DirectMusic_PlayMidi(ByVal FileName As String)
On Error GoTo ErrHandler

    If Not Music_On Then DirectMusic_StopMidi: Exit Sub
    
    If Music_On Then
        If ((Loader Is Nothing) Or (Performance Is Nothing)) Then
            DestroyDirectMusic
            InitDirectMusic
        End If
    End If
    
    Set Segment = Loader.LoadSegment(App.Path & MUSIC_PATH & FileName)
    
    ' repeat midi file
    Segment.SetLoopPoints 0, 0
    Segment.SetRepeats 100
    
    Performance.PlaySegment Segment, 0, 0
    
ErrHandler:
    
    If Err.Number > 0 And Err.Number <> DD_OK Then
        ErrorReport "DirectMusic_PlayMidi error " & Err.Number & " (" & Err.Description & ") with " & vbQuote & FileName & vbQuote
        Err.Clear
    End If
    
End Sub

Public Sub DirectMusic_StopMidi()

    If Segment Is Nothing Then Exit Sub
    Performance.Stop Segment, Nothing, 0, 0
    
End Sub

Public Sub DestroyDirectMusic()

    Set Performance = Nothing
    Set Loader = Nothing
    
End Sub

