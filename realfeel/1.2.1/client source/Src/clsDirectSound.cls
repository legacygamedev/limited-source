VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DS As DirectSound8
Public DSEnum As DirectSoundEnum8
Private Sound(1 To MAX_SOUND) As DirectSoundSecondaryBuffer8
Private SoundName(1 To MAX_SOUND) As String
Private SoundTaken(1 To MAX_SOUND) As Byte

Public Sub InitDirectSound()
'On Local Error Resume Next
Dim n As Byte, Target As Byte, Message As String

Set DS = Nothing
For n = 1 To MAX_SOUND
    Set Sound(n) = Nothing
Next n

Set DSEnum = DX8.GetDSEnum

If DSEnum.GetCount = 0 Then
    Set DS = DX8.DirectSoundCreate(vbNullString)
Else
    Message = "The following sound devices have been found on your PC:" & vbCrLf & vbCrLf
    For n = 1 To DSEnum.GetCount
        Message = Message & "(" & n & ") " & DSEnum.GetDescription(n) & vbCrLf
    Next n
    Message = Message & vbCrLf & "Please enter the number of the sound device you would like to use!"
    Target = InputBox(Message, "Multiple Sound Devices found!")
    Set DS = DX8.DirectSoundCreate(DSEnum.GetGuid(Target))
End If
'If Err.Number <> 0 Then
'    MsgBox "Error creating DirectSound object"
'    Call GameDestroy
'End If
DS.SetCooperativeLevel frmDualSolace.hwnd, DSSCL_NORMAL

For n = 1 To MAX_SOUND
    Set Sound(n) = Nothing
Next n

For n = 1 To MAX_SOUND
    SoundName(n) = ""
Next n

For n = 1 To MAX_SOUND
    SoundTaken(n) = False
Next n
End Sub

Public Function LoadSound(ByVal sFileName As String) As DirectSoundSecondaryBuffer8
Dim dsBuf As DSBUFFERDESC

'set up sound buffer for normal playback
dsBuf.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME

'load wave file into DirectSound buffer
Set LoadSound = DS.CreateSoundBufferFromFile(sFileName, dsBuf)
frmGameSettings.scrTempo.Value = CInt(dsBuf.fxFormat.lSamplesPerSec / 10)
'If Err.Number <> 0 Then
'    MsgBox "Error loading sound file: " & sFileName & vbCrLf & Err.Description
'    Call GameDestroy
'End If
End Function

Public Sub PlaySound(ByVal sFileName As String)
Dim n As Byte, soundlock As Byte
'Changed to DirectSound
'-smchronos
Dim Path As String
Path = App.Path & "\sound\"

soundlock = 0

'Find an open sound slot
For n = 1 To UBound(Sound)
    If SoundTaken(n) = False Then
        soundlock = n
        SoundTaken(n) = True
        Exit For
    End If
Next n

'If soundlock still equals 0, no sound slot open
'Clear first sound slot
If Not (Sound(1) Is Nothing) Then
    Sound(1).Stop
    Set Sound(1) = Nothing
    soundlock = 1
    SoundTaken(1) = True
End If

Set Sound(soundlock) = LoadSound(Path & sFileName & ".wav")
'Reset the sound data
Sound(soundlock).SetVolume frmGameSettings.scrAudio.Value
Sound(soundlock).SetFrequency CLng(frmGameSettings.scrTempo.Value) * 10
Sound(soundlock).SetPan 0
SoundName(soundlock) = sFileName

'now play sound once
Sound(soundlock).Play (DSBPLAY_DEFAULT)
SoundTaken(soundlock) = False
SoundName(soundlock) = ""
Set Sound(soundlock) = Nothing
End Sub

Public Sub DestroyDirectSound()
Dim n As Long

For n = 1 To MAX_SOUND
    If Not (Sound(n) Is Nothing) Then
        Sound(n).Stop
    End If
Next n

'clear wave files
For n = 1 To UBound(Sound)
    Set Sound(n) = Nothing
Next n

'clear directsound object
Set DS = Nothing
End Sub
