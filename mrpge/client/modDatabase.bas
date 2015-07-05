Attribute VB_Name = "modDatabase"
Option Explicit

Function FileExist(ByVal filename As String) As Boolean
    If Dir(App.Path & "\" & filename) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Sub AddLog(ByVal text As String)
Dim filename As String
Dim f As Long

    If Trim(command) = "-debug" Then
        If frmDebug.Visible = False Then
            frmDebug.Visible = True
        End If
        
        filename = App.Path & "\debug.txt"
    
        If Not FileExist("debug.txt") Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open filename For Append As #f
            Print #f, Time & ": " & text
        Close #f
    End If
End Sub

Sub SaveLocalMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\maps\map" & MapNum & ".dat"
            
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , SaveMap
    Close #f
End Sub

Sub LoadMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open filename For Binary As #f
        Get #f, , SaveMap
    Close #f
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
Dim filename As String
Dim f As Long
Dim TmpMap As MapRec

    filename = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open filename For Binary As #f
        Get #f, , TmpMap
    Close #f
        
    GetMapRevision = TmpMap.Revision
End Function

Function LoadSoundOption() As Boolean
Dim Sound As String
Dim music As String
Dim filenum As Long

filenum = FreeFile
    If PathExists(App.Path & "\Mconfig.dat") Then
        Open App.Path & "\Mconfig.dat" For Input As #filenum
            Input #filenum, Sound, music
        Close #filenum
        LoadSoundOption = CBool(Sound)
    Else
        LoadSoundOption = False
    End If
End Function

Function LoadMusicOption() As Boolean
Dim Sound As String
Dim music As String
Dim filenum As Long

filenum = FreeFile
    If PathExists(App.Path & "\Mconfig.dat") Then
        Open App.Path & "\Mconfig.dat" For Input As #filenum
            Input #filenum, Sound, music
        Close #filenum
        LoadMusicOption = CBool(music)
    Else
        LoadMusicOption = False
    End If
End Function

Sub setMconfig(ByVal music As Boolean, ByVal sounds As Boolean)
Dim filenum As Long
    filenum = FreeFile
    
        Open App.Path & "\Mconfig.dat" For Output As #filenum
            Print #filenum, sounds
            Print #filenum, music
        Close #filenum
End Sub



Public Function PathExists(ByVal Pathname As String, Optional ByVal IsFolder As Boolean = False) As Boolean
  PathExists = (Dir$(Pathname, vbArchive + vbHidden + vbReadOnly + vbSystem + IIf(IsFolder, vbDirectory, 0)) <> "")
End Function






