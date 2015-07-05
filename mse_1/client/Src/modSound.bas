Attribute VB_Name = "modSound"
Public Sub PlayMidi(Song As String)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added music constant.
'****************************************************************

Dim i As Long
    
    i = mciSendString("close all", 0, 0, 0)
    i = mciSendString("open " & App.Path & MUSIC_PATH & Song & " type sequencer alias background", 0, 0, 0)
    i = mciSendString("play background notify", 0, 0, frmMirage.hWnd)
End Sub

Public Sub StopMidi()
Dim i As Long
  
    i = mciSendString("close all", 0, 0, 0)
End Sub

Public Sub PlaySound(Sound As String)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added sound constant.
'****************************************************************

    Call sndPlaySound(App.Path & SOUND_PATH & Sound, SND_ASYNC Or SND_NOSTOP)
End Sub



