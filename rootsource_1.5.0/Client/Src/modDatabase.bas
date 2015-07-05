Attribute VB_Name = "Database"
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************

Public Function FileExist(ByVal Filename As String, Optional RAW As Boolean = False) As Boolean
    If Not RAW Then
        If LenB(Dir(App.Path & "\" & Filename)) > 0 Then
            FileExist = True
         End If
    Else
        If LenB(Dir(Filename)) > 0 Then
            FileExist = True
        End If
    End If
End Function

Public Sub AddLog(ByVal Text As String)
Dim Filename As String
Dim F As Long

    If DebugMode Then
        If Not frmDebug.Visible Then
            frmDebug.Visible = True
        End If
        
        Filename = App.Path & LOG_PATH & LOG_DEBUG
    
        If Not FileExist(LOG_DEBUG, True) Then
            F = FreeFile
            Open Filename For Output As #F
            Close #F
        End If
    
        F = FreeFile
        Open Filename For Append As #F
            Print #F, Time & ": " & Text
        Close #F
    End If
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim Filename As String
Dim F As Long

    Filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
            
    F = FreeFile
    Open Filename For Binary As #F
        Put #F, , map(5)
    Close #F
End Sub

Public Sub LoadMaps(ByVal i As Long)
    Dim BlankMap As MapRec
    Call ClearTempTile
    
    If FileExist(MAP_PATH & "map" & i & MAP_EXT) Then
        Call LoadMap(i, 5)
        If map(5).Up <> 0 Then
            Call LoadMap(map(5).Up, 2)
        Else
            map(2) = BlankMap
        End If
        
        If map(5).Down <> 0 Then
            Call LoadMap(map(5).Down, 8)
        Else
            map(8) = BlankMap
        End If
        
        If map(5).Left <> 0 Then
            Call LoadMap(map(5).Left, 4)
            If map(4).Up <> 0 Then
                Call LoadMap(map(4).Up, 1)
            Else
                map(1) = BlankMap
            End If
            If map(4).Down <> 0 Then
                Call LoadMap(map(4).Down, 7)
            Else
                map(7) = BlankMap
            End If
        Else
            map(4) = BlankMap
            map(1) = BlankMap
            map(7) = BlankMap
        End If
        
        If map(5).Right <> 0 Then
            Call LoadMap(map(5).Right, 6)
            If map(6).Up <> 0 Then
                Call LoadMap(map(6).Up, 3)
            Else
                map(3) = BlankMap
            End If
            If map(6).Down <> 0 Then
                Call LoadMap(map(6).Down, 9)
            Else
                map(9) = BlankMap
            End If
        Else
            map(6) = BlankMap
            map(3) = BlankMap
            map(9) = BlankMap
        End If
        
    If map(5).Music > 0 Then
        If Trim$(CStr(map(5).Music)) & ".mid" <> CurrentMusic Then
            StopMusic
            PlayMusic (Trim$(CStr(map(5).Music)) & ".mid")
            CurrentMusic = Trim$(CStr(map(5).Music)) & ".mid"
        End If
    Else
        StopMusic
        CurrentMusic = 0
    End If

    tMap(5) = i
    tMap(2) = map(5).Up
    tMap(8) = map(5).Down
    tMap(4) = map(5).Left
    tMap(6) = map(5).Right
    
    If tMap(4) <> 0 Then
        tMap(7) = map(4).Down
        tMap(1) = map(4).Up
    End If
    
    If tMap(6) <> 0 Then
        tMap(9) = map(6).Down
        tMap(3) = map(6).Up
    End If
    Else
        For i = 1 To 9
            map(i) = BlankMap
        Next
    End If
End Sub

Public Sub LoadMap(ByVal MapNum As Long, ByVal MapLoc As Long)
Dim Filename As String
Dim F As Long

    Filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
        
    F = FreeFile
    Open Filename For Binary As #F
        Get #F, , map(MapLoc)
    Close #F
End Sub

Public Sub LoadDataFile()
Dim Filename As String
Dim F As Long

    ' Check if the logs directory is there, if its not make it
    If LCase$(Dir(App.Path & "\data\logs", vbDirectory)) <> "logs" Then
        Call MkDir(App.Path & "\data\logs")
    End If
    
    ' Check if the music directory is there, if its not make it
    If LCase$(Dir(App.Path & "\data\music", vbDirectory)) <> "music" Then
        Call MkDir(App.Path & "\data\music")
    End If
    
    ' Check if the music directory is there, if its not make it
    If LCase$(Dir(App.Path & "\data\screenshots", vbDirectory)) <> "screenshots" Then
        Call MkDir(App.Path & "\data\screenshots")
    End If

    Filename = App.Path & DATA_PATH & "config.dat"

    If Not FileExist("data\config.dat") Then
            GameData.IP = "127.0.0.1"
            GameData.Port = 7234
            GameData.MusicExt = ".mid"
            GameData.ScreenNum = 0
            GameData.VerProcess = -1
            F = FreeFile
            Open Filename For Binary As #F
            Put #F, , GameData
            Close #F
        Else
            F = FreeFile
            Open Filename For Binary As #F
            Get #F, , GameData
            Close #F
        End If
End Sub

Public Sub CheckTiles()
Dim i As Long

    i = 1

    While FileExist(GFX_PATH & "tiles" & i & GFX_EXT)
        NumTilesets = NumTilesets + 1
        i = i + 1
    Wend

    frmMainGame.scrlTileSet.Max = NumTilesets

End Sub

Public Sub CheckSprites()
Dim i As Long

    i = 1

    While FileExist(GFX_PATH & "sprites\" & i & GFX_EXT)
        NumSprites = NumSprites + 1
        i = i + 1
    Wend

    ReDim Tr_Sprite(1 To NumSprites)
    
    ReDim SpriteTimer(1 To NumSprites)

End Sub

Public Sub CheckSpells()
Dim i As Long

    i = 1

    While FileExist(GFX_PATH & "Spells\" & i & GFX_EXT)
        NumSpells = NumSpells + 1
        i = i + 1
    Wend

    ReDim Tr_Spell(1 To NumSpells)
    
    ReDim SpellTimer(1 To NumSpells)

End Sub

Public Sub CheckItems()
Dim i As Long

    i = 1

    While FileExist(GFX_PATH & "Items\" & i & GFX_EXT)
        NumItems = NumItems + 1
        i = i + 1
    Wend

    ReDim DDS_Item(1 To NumItems)
    ReDim DDSD_Item(1 To NumItems)
    
    ReDim ItemTimer(1 To NumItems)

End Sub

Public Sub ClearPlayer(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString
End Sub

Public Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
End Sub

Public Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

Public Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(Index, MapNum)), LenB(MapItem(Index, MapNum)))
End Sub

Public Sub ClearMap()
    'Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    map(5).Name = vbNullString
    map(5).TileSet = 1
End Sub

Public Sub ClearMapItems()
Dim i As Long
Dim n As Long

    For i = 1 To MAX_MAP_ITEMS
        For n = 1 To MAX_MAPS
            Call ClearMapItem(i, n)
        Next n
    Next
End Sub

Public Sub ClearMapNpc(ByVal Index As Long, ByVal map As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(Index, map)), LenB(MapNpc(Index, map)))
End Sub

Public Sub ClearMapNpcs()
Dim i As Long
Dim n As Long

    For i = 1 To MAX_MAP_NPCS
        For n = 1 To MAX_MAPS
            Call ClearMapNpc(i, n)
        Next
    Next
End Sub

' *****************************
' ** Player Public Functions **
' *****************************

Public Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Name)
End Function
Public Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Public Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function
Public Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Public Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Sprite
    If GetPlayerSprite = 0 Then GetPlayerSprite = 1
End Function
Public Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Public Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function
Public Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Public Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Exp
End Function
Public Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Exp = Exp
End Sub

Public Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function
Public Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Public Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function
Public Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Public Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Public Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Vital(Vital) = Value
    
    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If
End Sub

Public Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            GetPlayerMaxVital = Player(Index).MaxHP
        Case MP
            GetPlayerMaxVital = Player(Index).MaxMP
        Case SP
            GetPlayerMaxVital = Player(Index).MaxSP
    End Select
End Function

Public Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    GetPlayerStat = Player(Index).Stat(Stat)
End Function
Public Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    Player(Index).Stat(Stat) = Value
End Sub

Public Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function
Public Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

Public Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).map
End Function
Public Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).map = MapNum
End Sub

Public Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).x
End Function
Public Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).x = x
End Sub

Public Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).y
End Function
Public Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).y = y
End Sub

Public Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function
Public Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Public Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = PlayerInv(InvSlot).Num
End Function
Public Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    PlayerInv(InvSlot).Num = ItemNum
End Sub

Public Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = PlayerInv(InvSlot).Value
End Function
Public Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    PlayerInv(InvSlot).Value = ItemValue
End Sub

Public Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = PlayerInv(InvSlot).Dur
End Function
Public Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    PlayerInv(InvSlot).Dur = ItemDur
End Sub

Public Function GetPlayerEquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
    GetPlayerEquipmentSlot = Player(Index).Equipment(EquipmentSlot)
End Function
Public Sub SetPlayerEquipmentSlot(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot) = InvNum
End Sub

Public Function CheckMapRevision(ByVal MapNum As Long, ByVal rev As Long) As Boolean
    Call LoadMap(MapNum, 5)
    If map(5).Revision = rev Then
        CheckMapRevision = True
    End If
End Function
