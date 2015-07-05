Attribute VB_Name = "modDatabase"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    If Not RAW Then
        If LenB(Dir(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
         End If
    Else
        If LenB(Dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If
End Function

Public Sub AddLog(ByVal Text As String)
Dim FileName As String
Dim f As Long

    If DEBUG_MODE Then
        If Not frmDebug.Visible Then
            frmDebug.Visible = True
        End If
        
        FileName = App.Path & LOG_PATH & LOG_DEBUG
    
        If Not FileExist(LOG_DEBUG, True) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open FileName For Append As #f
            Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long
Dim X As Long
Dim Y As Long

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
            
    If FileExist(FileName, True) Then Kill FileName
    
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map.Name
        Put #f, , Map.Revision
        Put #f, , Map.Moral
        Put #f, , Map.Up
        Put #f, , Map.Down
        Put #f, , Map.Left
        Put #f, , Map.Right
        Put #f, , Map.Music
        Put #f, , Map.BootMap
        Put #f, , Map.BootX
        Put #f, , Map.BootY
        Put #f, , Map.Shop
        Put #f, , Map.MaxX
        Put #f, , Map.MaxY

        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Put #f, , Map.Tile(X, Y)
            Next
            DoEvents
        Next

        For X = 1 To MAX_MAP_NPCS
            Put #f, , Map.Npc(X)
        Next
    Close #f
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long
Dim X As Long
Dim Y As Long

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
        
    ClearMap
        
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Map.Name
        Get #f, , Map.Revision
        Get #f, , Map.Moral
        Get #f, , Map.Up
        Get #f, , Map.Down
        Get #f, , Map.Left
        Get #f, , Map.Right
        Get #f, , Map.Music
        Get #f, , Map.BootMap
        Get #f, , Map.BootX
        Get #f, , Map.BootY
        Get #f, , Map.Shop
        Get #f, , Map.MaxX
        Get #f, , Map.MaxY
        
        ' have to set the tile()
        ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
        
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Get #f, , Map.Tile(X, Y)
            Next
        Next
        
        For X = 1 To MAX_MAP_NPCS
            Get #f, , Map.Npc(X)
        Next
    Close #f
    
    ClearTempTile
End Sub

Public Sub CheckTiles()
Dim i As Long

    i = 1

    While FileExist(GFX_PATH & "tiles" & i & GFX_EXT)
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend

    frmMirage.scrlTileSet.Max = NumTileSets

End Sub

Public Sub CheckSprites()
Dim i As Long

    i = 1

    While FileExist(GFX_PATH & "sprites\" & i & GFX_EXT)
        NumSprites = NumSprites + 1
        i = i + 1
    Wend

    ReDim DDS_Sprite(1 To NumSprites)
    ReDim DDSD_Sprite(1 To NumSprites)
    
    ReDim SpriteTimer(1 To NumSprites)

End Sub

Public Sub CheckSpells()
Dim i As Long

    i = 1

    While FileExist(GFX_PATH & "Spells\" & i & GFX_EXT)
        NumSpells = NumSpells + 1
        i = i + 1
    Wend

    ReDim DDS_Spell(1 To NumSpells)
    ReDim DDSD_Spell(1 To NumSpells)
    
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

Sub ClearPlayer(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString
End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

Sub ClearMapItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
End Sub

Sub ClearMap()
    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.TileSet = 1
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    
End Sub

Sub ClearMapItems()
Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
End Sub

Sub ClearMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next
End Sub

' **********************
' ** Player functions **
' **********************

Function GetPlayerName(ByVal Index As Long) As String
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value
    
    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            GetPlayerMaxVital = Player(Index).MaxHP
        Case MP
            GetPlayerMaxVital = Player(Index).MaxMP
        Case SP
            GetPlayerMaxVital = Player(Index).MaxSP
    End Select
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemNum = PlayerInv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemDur = PlayerInv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerEquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipmentSlot = Player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipmentSlot(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = InvNum
End Sub
