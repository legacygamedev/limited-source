Attribute VB_Name = "modDirectSound7"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ** DirectSound7, plays .WAV sfx files   **
' ******************************************

' DirectSound Variables.
Private DS As DirectSound

' The maximum amount of sounds.
Private Const SOUND_BUFFERS = 20

' Type that defines the buffers capabilities.
Private Type BufferCaps
    Volume As Boolean
    Frequency As Boolean
    Pan As Boolean
End Type

' Type that holds the buffer itself.
Private Type SoundArray
    DSBuffer As DirectSoundBuffer
    DSCaps As BufferCaps
    DSSourceName As String
End Type

' Contains all the data needed for sound manipulation.
Private Sound(1 To SOUND_BUFFERS) As SoundArray

' Contains the current sound index.
Private SoundIndex As Long

' Is the sound engine initiated?
Private SEngineIsLoaded As Boolean

' Has the array relooped yet?
Private SEngineRestart As Boolean

Public Sub InitDirectSound()
On Error GoTo ErrorHandle:

    'Create the DirectSound object
    Set DS = DX7.DirectSoundCreate(vbNullString)

    'Set the DirectSound object's cooperative level (Priority gives us sole control)
    DS.SetCooperativeLevel frmMirage.hWnd, DSSCL_PRIORITY

    ' Successfully initiated the sound engine.
    SEngineIsLoaded = True
    
    Exit Sub

ErrorHandle:

    Select Case Err.Number
    
        Case 91
            Call MsgBox("DirectX7 master object not created.")
    
        Case DSERR_NODRIVER
            Call MsgBox("No sound driver is available for use. Sound is disabled.")
            
        Case Else
            Call MsgBox("Unknown error has occured. Sound is disabled.")

    End Select


End Sub

Public Sub SoundLoad(ByVal File As String)
    Dim DSBufferDescription As DSBUFFERDESC
    Dim DSFormat As WAVEFORMATEX

    ' Set the sound index one higher for each sound.
    SoundIndex = SoundIndex + 1
    
    ' Reset the sound array if the array height is reached.
    If SoundIndex > UBound(Sound) Then
        SEngineRestart = True
        SoundIndex = 1
    End If

    ' Remove the sound if it exists (needed for re-loop).
    If SEngineRestart Then
        If GetState(SoundIndex) = DSBSTATUS_PLAYING Then
            Call SoundStop(SoundIndex)
            Call SoundRemove(SoundIndex)
        End If
    End If

    ' Load the sound array with the data given.
    With Sound(SoundIndex)
        .DSSourceName = File            'What is the name of the source?
        .DSCaps.Frequency = True        'Is this sound to have frequency altering capabilities?
        .DSCaps.Pan = True              'Is this sound to have Left and Right panning capabilities?
        .DSCaps.Volume = True           'Is this sound capable of altered volume settings?
    End With
    
    'Set the buffer description according to the data provided
    With DSBufferDescription
        If Sound(SoundIndex).DSCaps.Frequency Then
            .lFlags = .lFlags Or DSBCAPS_CTRLFREQUENCY
        End If
        If Sound(SoundIndex).DSCaps.Pan Then
            .lFlags = .lFlags Or DSBCAPS_CTRLPAN
        End If
        If Sound(SoundIndex).DSCaps.Volume Then
            .lFlags = .lFlags Or DSBCAPS_CTRLVOLUME
        End If
    End With

    'Set the Wave Format
    With DSFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 2
        .lSamplesPerSec = 22050
        .nBitsPerSample = 16
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With

    Set Sound(SoundIndex).DSBuffer = DS.CreateSoundBufferFromFile(App.Path & SOUND_PATH & Sound(SoundIndex).DSSourceName, DSBufferDescription, DSFormat)
End Sub

Public Sub SoundRemove(ByVal Index As Integer)
    'Reset all the variables in the sound array
    With Sound(Index)
        Set .DSBuffer = Nothing

        .DSCaps.Frequency = False
        .DSCaps.Pan = False
        .DSCaps.Volume = False
        .DSSourceName = vbNullString
    End With
End Sub

Public Sub SoundPlay(ByVal File As String, Optional ByVal Volume As Long = 100, Optional ByVal Pan As Long = 50)
On Error GoTo ErrorHandle:
    ' Check to see if DirectSound was successfully initalized.
    If Not SEngineIsLoaded Then Exit Sub

    ' Loads our sound into memory.
    Call SoundLoad(File)

    ' Sets the volume for the sound.
    Call SetVolume(SoundIndex, Volume)
    
    ' Sets the pan for the sound.
    Call SetPan(SoundIndex, Pan)
    
    ' Sets the frequency for the sound.
    ' Call SetFrequency(SoundIndex, Freq)

    ' Play the sound.
    Sound(SoundIndex).DSBuffer.Play DSBPLAY_DEFAULT
    
ErrorHandle:

    Select Case Err.Number
    
        ' File not found
        Case 432
            Call DevMsg("Error: Couldn't find 'SFX\" & File & "'", BrightRed)
            
    End Select
    
End Sub

Public Sub SoundStop(ByVal Index As Integer)
    'Stop the buffer and reset to the beginning
    Sound(Index).DSBuffer.Stop
    Sound(Index).DSBuffer.SetCurrentPosition 0
End Sub

Public Sub SoundPause(ByVal Index As Integer)
    'Stop the buffer
    Sound(Index).DSBuffer.Stop
End Sub

Public Sub SetFrequency(ByVal Index As Integer, ByVal Freq As Long)
    'Check to make sure that the buffer has the capability of altering its frequency
    If Not Sound(Index).DSCaps.Frequency Then Exit Sub

    'Alter the frequency according to the Freq provided
    Sound(Index).DSBuffer.SetFrequency Freq
End Sub

Public Sub SetVolume(ByVal Index As Integer, ByVal Vol As Long)
    'Check to make sure that the buffer has the capability of altering its volume
    If Not Sound(Index).DSCaps.Volume Then Exit Sub

    'Alter the volume according to the Vol provided
    If Vol > 0 Then
        Sound(Index).DSBuffer.SetVolume (60 * Vol) - 6000
    Else
        Sound(Index).DSBuffer.SetVolume -6000
    End If
End Sub

Public Sub SetPan(ByVal Index As Integer, ByVal Pan As Long)
    'Check to make sure that the buffer has the capability of altering its pan
    If Not Sound(Index).DSCaps.Pan Then Exit Sub

    'Alter the pan according to the pan provided
    Select Case Pan
        Case 0
            Sound(Index).DSBuffer.SetPan -10000
        Case 100
            Sound(Index).DSBuffer.SetPan 10000
        Case Else
            Sound(Index).DSBuffer.SetPan (100 * Pan) - 5000
    End Select
End Sub

Public Function GetFrequency(ByVal Index As Integer) As Long
    'Check to make sure that the buffer has the capability of altering its frequency
    If Not Sound(Index).DSCaps.Frequency Then Exit Function
    
    'Return the frequency value
    GetFrequency = Sound(Index).DSBuffer.GetFrequency()
End Function

Public Function GetVolume(ByVal Index As Integer) As Long
    'Check to make sure that the buffer has the capability of altering its volume
    If Not Sound(Index).DSCaps.Volume Then Exit Function
    
    'Return the volume value
    GetVolume = Sound(Index).DSBuffer.GetVolume()
End Function

Public Function GetPan(ByVal Index As Integer) As Long
    'Check to make sure that the buffer has the capability of altering its pan
    If Not Sound(Index).DSCaps.Pan Then Exit Function
    
    'Return the pan value
    GetPan = Sound(Index).DSBuffer.GetPan()
End Function

Public Function GetState(ByVal Index As Integer) As String
    'Returns the current state of the given sound
    GetState = Sound(Index).DSBuffer.GetStatus
End Function

Public Sub DestroyDirectSound()
Dim i As Long

    ' Delete all of the sounds created.
    If SEngineRestart Then
        For i = 1 To UBound(Sound)
            Call SoundStop(i)
            Call SoundRemove(i)
        Next
    Else
        For i = 1 To SoundIndex
            Call SoundStop(i)
            Call SoundRemove(i)
        Next
    End If
End Sub

