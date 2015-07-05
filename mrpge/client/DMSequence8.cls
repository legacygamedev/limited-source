VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDMSequence8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Visual Basic Game Programming With DirectX
' Chapter 11 : Playing Sound and Music With DirectX Audio
' DMSequence8 Class Source Code File
'----------------------------------------------------------------------
Option Explicit
Option Base 0

'local reference to clsDirectMusic8
Private dmusic As clsDirectMusic8

'DirectMusic loader object
Private dmLoader As DirectMusicLoader8

'DirectMusic segment object
Private dmSeg As DirectMusicSegment8

'DirectMusic segment state object
Private dmSegState As DirectMusicSegmentState8

Private Sub Class_Terminate()
    Set dmLoader = Nothing
    Set dmSeg = Nothing
    Set dmSegState = Nothing
End Sub

Public Property Get Length() As Long
    Length = dmSeg.GetLength
End Property
    
Public Property Get Name() As String
    Name = dmSeg.GetName
End Property

Public Property Get Looping() As Boolean
    Looping = CBool(dmSeg.GetRepeats)
End Property

Public Property Let Looping(ByVal bNewValue As Boolean)
    dmSeg.SetRepeats (CLng(bNewValue))
End Property

Public Property Get Playing() As Boolean
    Playing = dmusic.DMPerfObj.IsPlaying(dmSeg, dmSegState)
End Property

Public Function LoadMusic(ByRef dm As clsDirectMusic8, ByVal sFile As String) As Boolean
    On Local Error Resume Next
    LoadMusic = False
    If Len(sFile) = 0 Then Exit Function

    'set reference to clsDirectMusic8 object
    Set dmusic = dm
    If Err.number <> 0 Then Exit Function
    
    'create the DirectMusic loader object
    Set dmLoader = Nothing
    Set dmLoader = dmusic.DXObj.DirectMusicLoaderCreate
    If Err.number <> 0 Then Exit Function
    
    'remove existing segment
    If Not (dmSeg Is Nothing) Then
        dmSeg.Unload dmusic.DMPathObj
        Set dmSeg = Nothing
    End If

    'load the MIDI file
    Set dmSeg = dmLoader.LoadSegment(sFile)
    If Err.number <> 0 Then Exit Function
    dmSeg.SetStandardMidiFile

    'download the music segment
    dmSeg.Download dmusic.DMPathObj
    If Err.number <> 0 Then Exit Function

    'success
    LoadMusic = True
End Function

Public Sub PlayMusic()
    If dmSeg Is Nothing Then Exit Sub
    Set dmSegState = dmusic.DMPerfObj.PlaySegmentEx(dmSeg, 0, 0, Nothing, dmusic.DMPathObj)
End Sub

Public Sub StopMusic()
    If dmSeg Is Nothing Then Exit Sub
    dmusic.DMPerfObj.StopEx dmSeg, 0, 0
End Sub


