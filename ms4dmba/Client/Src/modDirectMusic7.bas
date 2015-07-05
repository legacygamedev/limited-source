Attribute VB_Name = "modDirectMusic7"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ** DirectMusic7, plays .MID music files **
' ******************************************

' DirectMusic variables
Public Performance As DirectMusicPerformance ' this controls the music
Public Segment As DirectMusicSegment ' this stores the music in memory
Private Loader As DirectMusicLoader 'parses the data from a file on the hard drive to the area in memory.

' Is the music engine initiated?
Private MEngineIsLoaded As Boolean

Public MusicVolume As Integer  ' ranges from 0-100
Public CurrentMusic As String

Public Sub InitDirectMusic()
On Error GoTo ErrorHandle:

Dim i As Long
Dim portcaps As DMUS_PORTCAPS

    ' Initialize Loader
    Set Loader = DX7.DirectMusicLoaderCreate
    Call Loader.SetSearchDirectory(App.Path & MUSIC_PATH) ' set music directory
        
    ' Initialize Performance
    Set Performance = DX7.DirectMusicPerformanceCreate
    Call Performance.Init(Nothing, frmMirage.hWnd)
    
    ' Find usable port
    For i = 1 To Performance.GetPortCount
        Call Performance.GetPortCaps(i, portcaps)
        If portcaps.lFlags And DMUS_PC_SOFTWARESYNTH Then
            Call Performance.SetPort(i, 1)
            Exit For
        End If
    Next i

    ' The volume is specified in hundredths of decibels (dB).
    MusicVolume = 60
    Call Performance.SetMasterVolume(MusicVolume * 42 - 3000)
    
    ' ensure that the DLS data for the instruments is downloaded to the port
    Call Performance.SetMasterAutoDownload(True)
    
    MEngineIsLoaded = True
    
    Exit Sub
    
ErrorHandle:

    Select Case Err.Number
    
        Case 91
            Call MsgBox("DirectX7 master object not created.")

        Case Else
            Call MsgBox("Unknown error has occured. Music is disabled.")

    End Select
    
    MEngineIsLoaded = False
   
    
End Sub

Public Sub DirectMusic_PlayMidi(FileName As String)
On Error GoTo ErrorHandle

    If MEngineIsLoaded = False Then Exit Sub

    ' load segment from file
    Set Segment = Loader.LoadSegment(FileName)
    
    ' If not set, certain details of playback might not be handled properly.
    If StrConv(Right(FileName, 4), vbLowerCase) = ".mid" Then
        Call Segment.SetStandardMidiFile
    End If

    ' init repeat data
    Segment.SetLoopPoints 0, 0 ' range to repeat, entire segment
    Segment.SetRepeats 100 ' number of times to repeat
    
    ' play music from beginning of segment
    Performance.PlaySegment Segment, 0, 0
    
    Exit Sub
    
ErrorHandle:

    Select Case Err.Number
    
         ' File open failed because the file does not exist or is locked.
        Case DMUS_E_LOADER_FAILEDOPEN
            Call DevMsg("Error: Could not load MIDI file!", BrightRed)
            
    End Select
    
End Sub

Public Sub DirectMusic_StopMidi()
    Performance.Stop Segment, Nothing, 0, 0
    CurrentMusic = 0
End Sub

Public Sub DestroyDirectMusic()
    Set Performance = Nothing
    Set Segment = Nothing
    Set Loader = Nothing
End Sub

