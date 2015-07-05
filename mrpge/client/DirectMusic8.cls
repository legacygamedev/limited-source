VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectMusic8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Visual Basic Game Programming With DirectX
' Chapter 11 : Playing Sound and Music With DirectX Audio
' DirectMusic8 Class Source Code File
'----------------------------------------------------------------------
Option Explicit
Option Base 0

'main DirectX object
Dim objDX As New DirectX8

'DirectMusic performance object
Private dmPerf As DirectMusicPerformance8

'DirectMusic audio path object
Private dmPath As DirectMusicAudioPath8

'DirectMusic audio parameters
Dim dmA As DMUS_AUDIOPARAMS

Private Sub Class_Initialize()
    'create the DirectX object
    Set objDX = New DirectX8
End Sub

Private Sub Class_Terminate()
    'stop music playback
    If Not (dmPerf Is Nothing) Then
        dmPerf.CloseDown
    End If
    
    'delete DirectMusic objects
    Set dmPath = Nothing
    Set dmPerf = Nothing
    Set objDX = Nothing
End Sub

Public Property Let Volume(ByVal lNewValue As Long)
    If lNewValue <= 0 Then
        dmPath.SetVolume lNewValue, 0
    End If
End Property

Public Property Get ClockTime() As Long
    ClockTime = dmPerf.GetClockTime
End Property

Public Property Get MusicTime() As Long
    MusicTime = dmPerf.GetMusicTime
End Property

Public Property Get LatencyTime() As Long
    LatencyTime = dmPerf.GetLatencyTime
End Property

Public Property Get DXObj() As DirectX8
    Set DXObj = objDX
End Property

Public Property Get DMPerfObj() As DirectMusicPerformance8
    Set DMPerfObj = dmPerf
End Property

Public Property Get DMPathObj() As DirectMusicAudioPath8
    Set DMPathObj = dmPath
End Property

Public Function Startup(ByVal lWindowHandle As Long) As Boolean
    On Local Error Resume Next
    Startup = False
    
    'create the DirectMusic performance object
    Set dmPerf = objDX.DirectMusicPerformanceCreate
    If Err.number <> 0 Then Exit Function
    
    'initialize DirectMusic
    dmPerf.InitAudio lWindowHandle, DMUS_AUDIOF_ALL, dmA
    If Err.number <> 0 Then Exit Function
    
    'create the DirectMusic audio path object
    Set dmPath = dmPerf.CreateStandardAudioPath( _
        DMUS_APATH_DYNAMIC_3D, 64, True)
        
    If Err.number <> 0 Then Exit Function
    
    Startup = True
End Function

