Attribute VB_Name = "modMapServer"
Option Explicit
Public MapsAvailable() As Boolean

Sub IncomingMapData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim Start As Long

    frmMirage.MapSocket.GetData Buffer, vbString, DataLength
    MapBuffer = MapBuffer & Buffer
        
    Start = InStr(MapBuffer, END_CHAR)
    Do While Start > 0
        Packet = Mid(MapBuffer, 1, Start - 1)
        MapBuffer = Mid(MapBuffer, Start + 1, Len(MapBuffer))
        Start = InStr(MapBuffer, END_CHAR)
        If Len(Packet) > 0 Then
            Call HandleMapData(Packet)
        End If
    Loop
End Sub

Sub HandleMapData(ByVal Data As String)
Dim Parse() As String
Dim i As Long, n As Long, X As Long, Y As Long

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "mapdata" Then
        n = 1
        
        SaveMap.name = Parse(n + 1)
        SaveMap.Revision = Val(Parse(n + 2))
        SaveMap.Moral = Val(Parse(n + 3))
        SaveMap.Up = Val(Parse(n + 4))
        SaveMap.Down = Val(Parse(n + 5))
        SaveMap.Left = Val(Parse(n + 6))
        SaveMap.Right = Val(Parse(n + 7))
        SaveMap.Music = Parse(n + 8)
        SaveMap.BootMap = Val(Parse(n + 9))
        SaveMap.BootX = Val(Parse(n + 10))
        SaveMap.BootY = Val(Parse(n + 11))
        SaveMap.Indoor = Val(Parse(n + 12))
        SaveMap.Random = 0
        
        n = n + 13
        
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                SaveMap.Tile(X, Y).Ground = Val(Parse(n))
                SaveMap.Tile(X, Y).Mask = Val(Parse(n + 1))
                SaveMap.Tile(X, Y).Anim = Val(Parse(n + 2))
                SaveMap.Tile(X, Y).Mask2 = Val(Parse(n + 3))
                SaveMap.Tile(X, Y).M2Anim = Val(Parse(n + 4))
                SaveMap.Tile(X, Y).Fringe = Val(Parse(n + 5))
                SaveMap.Tile(X, Y).FAnim = Val(Parse(n + 6))
                SaveMap.Tile(X, Y).Fringe2 = Val(Parse(n + 7))
                SaveMap.Tile(X, Y).F2Anim = Val(Parse(n + 8))
                SaveMap.Tile(X, Y).Type = Val(Parse(n + 9))
                SaveMap.Tile(X, Y).Data1 = Val(Parse(n + 10))
                SaveMap.Tile(X, Y).Data2 = Val(Parse(n + 11))
                SaveMap.Tile(X, Y).Data3 = Val(Parse(n + 12))
                SaveMap.Tile(X, Y).String1 = Parse(n + 13)
                SaveMap.Tile(X, Y).String2 = Parse(n + 14)
                SaveMap.Tile(X, Y).String3 = Parse(n + 15)
                
                n = n + 16
            Next X
        Next Y
        
        For X = 1 To MAX_MAP_NPCS
            SaveMap.Npc(X) = Val(Parse(n))
            n = n + 1
        Next X
                
        ' Save the map
        Call SaveLocalMap(Val(Parse(1)))
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMapEditor.Visible = False
            'frmmirage.show
            'frmMirage.picMapEditor.Visible = False
            
            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If
            
            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdone" Then
        CheckMap(GetPlayerMap(MyIndex)) = SaveMap
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
        
        GettingMap = False
        
        ' Play music
        If Trim(CheckMap(GetPlayerMap(MyIndex)).Music) <> "None" Then
            If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
                Call PlayMidi(Trim(CheckMap(GetPlayerMap(MyIndex)).Music))
            Else
                Call StopMidi
            End If
        Else
            Call StopMidi
        End If
        
        Exit Sub
    End If
End Sub

Function ConnectToMapServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsMapConnected Then
        ConnectToMapServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMirage.MapSocket.Close
    frmMirage.MapSocket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsMapConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    
    If IsMapConnected Then
        ConnectToMapServer = True
    Else
        ConnectToMapServer = False
    End If
End Function

Function IsMapConnected() As Boolean
    If frmMirage.MapSocket.State = sckConnected Then
        IsMapConnected = True
    Else
        IsMapConnected = False
    End If
End Function

Sub SendMapData(ByVal Data As String)
    If IsMapConnected Then
        frmMirage.MapSocket.SendData Data
        DoEvents
    End If
End Sub

Sub LoadMaps(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    If FileExist("maps\map" & MapNum & ".dat") Then
        filename = App.Path & "\maps\map" & MapNum & ".dat"
            
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , CheckMap(MapNum)
        Close #f
        MapsAvailable(MapNum) = True
    End If
End Sub
