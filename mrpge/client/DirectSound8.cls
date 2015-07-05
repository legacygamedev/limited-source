VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectSound8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Visual Basic Game Programming With DirectX
' Chapter 11 : Playing Sound and Music With DirectX Audio
' DirectSound8 Class Source Code File
'----------------------------------------------------------------------
Option Explicit
Option Base 0

Dim objDS As DirectSound

Private Sub Class_Terminate()
    Set objDS = Nothing
    Set objDX = Nothing
End Sub

Public Property Get DSObj() As DirectSound
    Set DSObj = objDS
End Property

Public Sub Startup(ByRef lWindowHandle As Long)
    'create the DirectX8 object
    'Set objDX = New DirectX8
    
    'create the DirectSound8 object
    Set objDS = DX.DirectSoundCreate("")

    'set the priority level for DirectSound
    objDS.SetCooperativeLevel lWindowHandle, DSSCL_PRIORITY
End Sub


