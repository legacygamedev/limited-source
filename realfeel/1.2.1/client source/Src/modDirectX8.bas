Attribute VB_Name = "modDirectX8"
Option Explicit
'Added DirectX8 module (04/23/07)
'Moved DirectMusic and DirectSound to this module (04/23/07)
'Created Direct3D processes here
'-smchronos

Public Function InitDirectX8() As Boolean
Dim FSys As Object, Folder As Object, FolderFiles As Object, File As Object, FileName As String
Dim MusicList() As String, n As Byte, MusicSelected As String
InitDirectX8 = False
'Let's start up that DirectX8 instance
Set DX8 = New DirectX8

Set Direct3D = New clsDirect3D
Set DirectMusic = New clsDirectMusic
Set DirectSound = New clsDirectSound
Set DirectShow = New clsDirectShow

'Load Direct3D
Call SetStatus("Initializing Direct3D...")
DoEvents
Call Direct3D.InitDirect3D(frmDualSolace.picScreen.hwnd, frmDualSolace.picScreen.ScaleWidth, frmDualSolace.picScreen.ScaleHeight, False, True)

'load DirectMusic
Call SetStatus("Initializing DirectMusic...")
DoEvents
Call DirectMusic.InitDirectMusic

'Load DirectShow
Call SetStatus("Initializing DirectShow...")
DoEvents
If UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_STYLE")) <> "OFF" Then
    Call DirectShow.SetPlayBackBalance(0)
    Call DirectShow.SetPlayBackVolume(0)
    
    If UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_STYLE")) = "FIXED" Then
        If Not DirectShow.LoadMP3(App.Path & "\" & GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_POINTER")) Then
            MsgBox ("Error loading mp3!")
            Call GameDestroy
        End If
    ElseIf UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_STYLE")) = "RANDOM" Then
        ' Create the file
        Set FSys = CreateObject("Scripting.FileSystemObject")

        'Set the folder objects
        Set Folder = FSys.GetFolder(App.Path & MUSIC_PATH)
        Set FolderFiles = Folder.Files
        
        n = 1
        
        For Each File In FolderFiles
            ' Make sure it is a music file
            If UCase$(Right$(Mid(File, Len(App.Path & MUSIC_PATH) + 1, (Len(File) - Len(App.Path & MUSIC_PATH))), 3)) = "MP3" Or UCase$(Right$(Mid(File, Len(App.Path & MUSIC_PATH) + 1, (Len(File) - Len(App.Path & MUSIC_PATH))), 3)) = "MID" Or UCase$(Right$(Mid(File, Len(App.Path & MUSIC_PATH) + 1, (Len(File) - Len(App.Path & MUSIC_PATH))), 3)) = "MIDI" Then
                ReDim Preserve MusicList(1 To n)
                MusicList(n) = Mid(File, Len(App.Path & MUSIC_PATH) + 1, (Len(File) - Len(App.Path & MUSIC_PATH)))
                Debug.Print MusicList(n)
                n = n + 1
            End If
        Next File
        
        'Destroy the folder objects
        Set File = Nothing
        Set FolderFiles = Nothing
        Set Folder = Nothing
        Set FSys = Nothing
        
        ' Set up the random seed generator
        Randomize
        
        ' Choose a random number
        ' Find the music
        MusicSelected = MusicList(Int(Rnd * (n - 1)) + 1)
        If Not DirectShow.LoadMP3(App.Path & "\" & GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_POINTER") & MusicSelected) Then
            MsgBox ("Error loading mp3!")
            Call GameDestroy
        End If
    Else
        If Not DirectShow.LoadMP3(App.Path & "\" & GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_POINTER")) Then
            MsgBox ("Error loading mp3!")
            Call GameDestroy
        End If
    End If
    
    If UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "NORMAL" Then
        Call DirectShow.SetPlayBackSpeed(1)
    ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "MODERATELY_SLOW" Then
        Call DirectShow.SetPlayBackSpeed(0.75)
    ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "SLOW" Then
        Call DirectShow.SetPlayBackSpeed(0.5)
    ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "EXTREMELY_SLOW" Then
        Call DirectShow.SetPlayBackSpeed(0.25)
    ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "MAX_SLOW" Then
        Call DirectShow.SetPlayBackSpeed(0.05)
    ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "MODERATELY_FAST" Then
        Call DirectShow.SetPlayBackSpeed(1.25)
    ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "FAST" Then
        Call DirectShow.SetPlayBackSpeed(1.5)
    ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "EXTREMELY_FAST" Then
        Call DirectShow.SetPlayBackSpeed(1.75)
    ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "MAX_FAST" Then
        Call DirectShow.SetPlayBackSpeed(2)
    Else
        Call DirectShow.SetPlayBackSpeed(1)
    End If
    
    If UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYLOOP")) = "TRUE" Then
        LoopIntro = True
    ElseIf UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYLOOP")) = "False" Then
        LoopIntro = False
    Else
        LoopIntro = True
    End If
    
    'Play the mp3
    Call DirectShow.PlayMP3
End If

'load DirectSound (not tested)
'Call DirectSound.InitDirectSound

'load pictureboxes for the editor, they are actually used for blitting right now
Call SetStatus("Loading graphics for equipment and items...")
DoEvents
Call SetPicSize(App.Path + GFX_PATH + "items" + GFX_EXT, frmEditor.picItems)
frmEditor.picItems.Picture = LoadPicture(App.Path & GFX_PATH & "items" & GFX_EXT)

'load these with the standard setting
AllowMovement = False
AttributeDisplay = True
DepictAttributeTiles = True

InitDirectX8 = True
End Function

Sub DestroyDirectX8()
    Call Direct3D.UnloadDirect3D(True)
    Call DirectMusic.DestroyDirectMusic
    'Call DirectSound.DestroyDirectSound
    Call DirectShow.StopMP3
    Call DirectShow.TerminateEngine
    Set DX8 = Nothing
End Sub

