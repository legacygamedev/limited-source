VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dmLoader As DirectMusicLoader8
Private dmPerf As DirectMusicPerformance8
Private dmSeg As DirectMusicSegment8
Private dmSegState As DirectMusicSegmentState8
Private dmPath As DirectMusicAudioPath8
Private dmAudio As DMUS_AUDIOPARAMS

Public Sub InitDirectMusic()
On Local Error Resume Next

'Set the DirectX8 music loader
Set dmLoader = DX8.DirectMusicLoaderCreate
If Err.Number <> 0 Then
    MsgBox "Error creating Direct Music loader object!"
    Call GameDestroy
End If

Set dmPerf = DX8.DirectMusicPerformanceCreate
If Err.Number <> 0 Then
    MsgBox "Error creating Direct Music performance object!"
    Call GameDestroy
End If

dmPerf.InitAudio frmDualSolace.hwnd, DMUS_AUDIOF_ALL, dmAudio
If Err.Number <> 0 Then
    MsgBox "Error initializing Direct Music audio system!"
    Call GameDestroy
End If

dmPerf.SetMasterAutoDownload False
If Err.Number <> 0 Then
    MsgBox "Error setting Direct Music Auto-Download option!"
    Call GameDestroy
End If

'dmPerf.SetMasterVolume 0
'If Err.Number <> 0 Then
'    MsgBox "Error setting Direct Music Master Volume!"
'    Call GameDestroy
'End If

'dmPerf.SetMasterTempo 1#
'If Err.Number <> 0 Then
'    MsgBox "Error setting Direct Music Master Tempo!"
'    Call GameDestroy
'End If

Set dmPath = dmPerf.CreateStandardAudioPath(DMUS_APATH_DYNAMIC_3D, 64, True)
If Err.Number <> 0 Then
    MsgBox "Error creating Direct Music audio path!" & vbCrLf & "Error number: " & Err.Number & vbCrLf & Err.Description
    Call GameDestroy
End If
End Sub

Public Sub DestroyDirectMusic()
'stop music
If Not (dmPerf Is Nothing) And Not (dmSeg Is Nothing) Then
    dmPerf.StopEx dmSeg, 0, 0
    dmPerf.CloseDown
End If

'delete the objects
Set dmLoader = Nothing
Set dmSeg = Nothing
Set dmSegState = Nothing
Set dmPath = Nothing
Set dmPerf = Nothing
End Sub

Public Sub PlayMusic(Song As String)
'Read Microsoft Visual Basic Game Programming With DirectX
'-smchronos
Dim Path As String
Path = App.Path & "\music\"

'load music and check for an error
If Not LoadMusic(Path & Song) Then
    MsgBox "Error loading music file:" & vbCrLf & Path & Song
    Call GameDestroy
End If

If dmSeg Is Nothing Then Exit Sub

Set dmSegState = dmPerf.PlaySegmentEx(dmSeg, 0, 0, Nothing, dmPath)
End Sub

Public Sub StopMusic()
'Read Microsoft Visual Basic Game Programming With DirectX
'-smchronos
If dmSeg Is Nothing Then Exit Sub
dmPerf.StopEx dmSeg, 0, 0
End Sub

Public Function LoadMusic(sFile As String) As Boolean
On Local Error Resume Next
LoadMusic = False

If Len(sFile) = 0 Then Exit Function

'Unload existing segment
If Not (dmSeg Is Nothing) Then
    dmPerf.StopEx dmSeg, 0, 0
    dmSeg.Unload dmPath
    Set dmSeg = Nothing
End If

'Load the music file
dmLoader.SetSearchDirectory (App.Path & "\music\")
Set dmSeg = dmLoader.LoadSegment(sFile)
If Err.Number <> 0 Then Exit Function
dmSeg.SetStandardMidiFile

'download the music segment
dmSeg.Download dmPath
If Err.Number <> 0 Then Exit Function

'Set the repeats (-1 makes it continuous)
dmSeg.SetRepeats -1
If Err.Number <> 0 Then Exit Function

LoadMusic = True
End Function
