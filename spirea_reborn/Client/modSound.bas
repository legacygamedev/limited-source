Attribute VB_Name = "modSound"
Public Sub PlayMidi(Song As String)
Dim i As Long
    
    i = mciSendString("close all", 0, 0, 0)
    i = mciSendString("open " & Song & " type sequencer alias background", 0, 0, 0)
    i = mciSendString("play background notify", 0, 0, frmMirage.hWnd)
End Sub

Public Sub StopMidi()
Dim i As Long
  
    i = mciSendString("close all", 0, 0, 0)
End Sub

Public Sub PlaySound(Sound As String)
    Call sndPlaySound(App.Path & "\" & Sound, SND_ASYNC Or SND_NOSTOP)
End Sub



