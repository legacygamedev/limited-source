VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSWave8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Visual Basic Game Programming With DirectX
' Chapter 11 : Playing Sound and Music With DirectX Audio
' DSWave8 Class Source Code File
'----------------------------------------------------------------------
Option Explicit
Option Base 0

'DirectSound objects
Dim dsBuffer As DirectSoundBuffer
'Dim dsBuffer As DirectSoundSecondaryBuffer8

'DirectSound structures
Dim format As WAVEFORMATEX
Dim dsBuf As DSBUFFERDESC
Dim DSFormat As WAVEFORMATEX

'program variables
Dim bInterrupt As Boolean
Dim bLooping As Boolean

Private Sub Class_Initialize()
    bInterrupt = False
    bLooping = False
End Sub

Private Sub Class_Terminate()
    StopSound
    Set ds3dBuffer = Nothing
    Set dsBuffer = Nothing
End Sub

Public Property Get DSBObj() As DirectSoundBuffer
    Set DSBObj = dsBuffer
End Property

Public Property Set DSBObj(ByRef dssbNewValue As DirectSoundBuffer)
    Set dsBuffer = dssbNewValue
End Property

Public Property Get Interrupt() As Boolean
    Interrupt = bInterrupt
End Property

Public Property Let Interrupt(ByVal bNewValue As Boolean)
    bInterrupt = bNewValue
End Property

Public Property Get Looping() As Boolean
    Looping = bLooping
End Property

Public Property Let Looping(ByVal bNewValue As Boolean)
    bLooping = bNewValue
End Property

Public Property Get Frequency() As Long
    Frequency = dsBuffer.GetFrequency
End Property

Public Property Get Channels() As Long
    dsBuffer.GetFormat format
    Channels = format.nChannels
End Property

Public Property Get Samples() As Long
    dsBuffer.GetFormat format
    Samples = format.lSamplesPerSec
End Property

Public Property Get BitsPerSample() As Long
    dsBuffer.GetFormat format
    BitsPerSample = format.nBitsPerSample
End Property

Public Property Get Size() As Long
    dsBuffer.GetFormat format
    Size = format.nSize
End Property

Public Sub LoadSound(ByRef objDS As clsDirectSound8, ByVal sFilename As String)
    'set buffer parameter for normal playback
    dsBuf.lFlags = DSBCAPS_STATIC
    
    'load wave file into DirectSound buffer
    Set dsBuffer = Nothing
    Set dsBuffer = objDS.DSObj.CreateSoundBufferFromFile(sFilename, dsBuf, DSFormat)
    'set positional sound environment
End Sub

Public Sub PlaySound()
    'stop currently playing waves?
    If bInterrupt Then
        StopSound
    End If
    
    'loop the sound?
    If bLooping Then
        dsBuffer.Play DSBPLAY_LOOPING
    Else
        dsBuffer.Play DSBPLAY_DEFAULT
    End If
End Sub

Public Sub StopSound()
    dsBuffer.Stop
    dsBuffer.SetCurrentPosition 0
End Sub


