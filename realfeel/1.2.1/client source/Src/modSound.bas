Attribute VB_Name = "modSound"
Const SND_ASYNC = &H1

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal _
    lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

' Play a WAV file.
'
' FileName is a string containing the full path of the file.
' If SyncExec is True, the sound is played synchronously
' Returns True if no errors occurred

Function PlayWAV(FileName As String, Optional SyncExec As Boolean) As Boolean
    If SyncExec Then
        ' play the file synchronously
        PlayWAV = PlaySound(App.Path & "\sound\" & FileName, 0, 0)
    Else
        ' play the file asynchronously
        PlayWAV = PlaySound(App.Path & "\sound\" & FileName, 0, SND_ASYNC)
    End If
End Function

'OLD MIDI PLAYING SUBS! DON'T USE! NOW USING DIRECT MUSIC!
'---------------------------------------------
'Public Sub PlayMidi(Song As String)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added music constant.
'****************************************************************
'
'Dim i As Long
'
'    i = mciSendString("close all", 0, 0, 0)
'    i = mciSendString("open " & App.Path & MUSIC_PATH & Song & " type sequencer alias background", 0, 0, 0)
'    i = mciSendString("play background notify", 0, 0, frmDualSolace.hWnd)
'End Sub

'---------------------------------------------

'Public Sub StopMidi()
'Dim i As Long
'
'    i = mciSendString("close all", 0, 0, 0)
'End Sub
'---------------------------------------------
