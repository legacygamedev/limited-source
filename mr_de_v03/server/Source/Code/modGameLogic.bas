Attribute VB_Name = "modGameLogic"
Option Explicit

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function Clamp(ByVal Value As Long, ByVal Min As Long, ByVal Max As Long) As Long
    Clamp = Value
    If Value < Min Then Clamp = Min
    If Value > Max Then Clamp = Max
End Function

Public Function IsAlphaNumeric(s As String) As Boolean
    If Not s Like "*[!0-9A-Za-z]*" Then IsAlphaNumeric = True
End Function

Public Function IsAlpha(s As String) As Boolean
    If Not s Like "*[!A-Za-z]*" Then IsAlpha = True
End Function

Public Sub SetExpMod(ByVal Value As Long)

    ExpMod = Clamp(Value, 0, MAX_INTEGER)
    PutVar App.Path & "\Data\Core Files\Configuration.ini", "server", "ExpMod", CStr(ExpMod)
    SendGlobalMsg "[Breaking Realm News!] The experience modifier has changed.", White
    AddLog "The experience modifier has changed to " & ExpMod, ADMIN_LOG
End Sub

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long
    
    For i = 1 To OnlinePlayersCount
        ' Make sure we dont try to check a name thats to small
        If Len(Current_Name(OnlinePlayers(i))) >= Len(Trim$(Name)) Then
            If UCase$(Mid$(Current_Name(OnlinePlayers(i)), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                FindPlayer = OnlinePlayers(i)
                Exit Function
            End If
        End If
    Next
End Function

'Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
'Dim i As Long
'
'    For i = 1 To OnlinePlayersCount
'        If Current_Map(OnlinePlayers(i)) = MapNum Then
'            GetTotalMapPlayers = GetTotalMapPlayers + 1
'        End If
'    Next
'End Function

Function FindOpenPlayerSlot() As Long
Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next
End Function

'***********************************************
' Checks for a player at a certain map, x and y
'***********************************************
Public Function PlayerAtPosition(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
Dim i As Long
    
'    For i = 1 To OnlinePlayersCount
'        If Current_Map(OnlinePlayers(i)) = MapNum Then
'            If Current_X(OnlinePlayers(i)) = X Then
'                If Current_Y(OnlinePlayers(i)) = Y Then
'                    PlayerAtPosition = True
'                    Exit Function
'                End If
'            End If
'        End If
'    Next
    For i = 1 To MapData(MapNum).MapPlayersCount
        If Current_X(MapData(MapNum).MapPlayers(i)) = X Then
            If Current_Y(MapData(MapNum).MapPlayers(i)) = Y Then
                PlayerAtPosition = True
                Exit Function
            End If
        End If
    Next
End Function

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim NpcNum As Long, MobGroup As Long, Dir As Long
Dim i As Long, n As Long, X As Long, Y As Long
Dim Spawned As Boolean
Dim SpawnList() As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Then Exit Sub
    If MapNpcNum > MapData(MapNum).NpcCount Then Exit Sub
    If MapNum <= 0 Then Exit Sub
    If MapNum > MAX_MAPS Then Exit Sub
    
    Spawned = False
    NpcNum = MapData(MapNum).Npc(MapNpcNum).Npc
    MobGroup = MapData(MapNum).Npc(MapNpcNum).MobGroup
    ReDim SpawnList(CLng(Map(MapNum).MaxY) * CLng((Map(MapNum).MaxX + 1)) + Map(MapNum).MaxX)
    
    If NpcNum <= 0 Then Exit Sub
                   
    'Check if theres a spawn tile for the specific npc
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_MOBSPAWN Then
                If Map(MapNum).Tile(X, Y).Data1 = MobGroup Then
                    SpawnList(i) = (Y * (Map(MapNum).MaxX + 1)) + X
                    i = i + 1
                End If
            End If
        Next
    Next
    
    ' Check if there's any spawn points
    If i <= 0 Then Exit Sub
    
    ' Cut down the spawnlist
    ReDim Preserve SpawnList(i - 1)

    ' NPC doesent have specific spawn point
    ' Well try 100 times to randomly place the sprite
    For i = 0 To 100
        n = Rand(0, UBound(SpawnList))
        X = SpawnList(n) Mod (Map(MapNum).MaxX + 1)
        Y = SpawnList(n) \ (Map(MapNum).MaxX + 1)
        
        ' Check if the tile is walkable
        If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Or _
            Map(MapNum).Tile(X, Y).Type = TILE_TYPE_MOBSPAWN Then
                MapData(MapNum).MapNpc(MapNpcNum).X = X
                MapData(MapNum).MapNpc(MapNpcNum).Y = Y
                ' Random direction
                Dir = Int(Rnd * 4)
                ' Checks if you specified a spawn direction
                If Map(MapNum).Tile(X, Y).Data2 > -1 Then Dir = Map(MapNum).Tile(X, Y).Data2
                Spawned = True
                Exit For
        End If
    Next
    
    ' Didn't spawn, so now we'll just try to find a free tile
    If Not Spawned Then
        For i = 0 To UBound(SpawnList)
            X = SpawnList(i) Mod (Map(MapNum).MaxX + 1)
            Y = SpawnList(i) \ (Map(MapNum).MaxX + 1)
            
            ' Check if the tile is walkable
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Or _
                Map(MapNum).Tile(X, Y).Type = TILE_TYPE_MOBSPAWN Then
                    MapData(MapNum).MapNpc(MapNpcNum).X = X
                    MapData(MapNum).MapNpc(MapNpcNum).Y = Y
                    ' Random direction
                    Dir = Int(Rnd * 4)
                    ' Checks if you specified a spawn direction
                    If Map(MapNum).Tile(X, Y).Data2 > -1 Then Dir = Map(MapNum).Tile(X, Y).Data2
                    Spawned = True
                    Exit For
            End If
        Next
    End If
         
    ' If we suceeded in spawning then send it to everyone
    If Spawned Then
        MapData(MapNum).MapNpc(MapNpcNum).Num = NpcNum
        MapData(MapNum).MapNpc(MapNpcNum).Target = 0
        MapData(MapNum).MapNpc(MapNpcNum).Dir = Dir
        
        For i = 1 To Vitals.Vital_Count
            MapData(MapNum).MapNpc(MapNpcNum).Vital(i) = Npc_MaxVital(NpcNum, i)
        Next
    
        SendSpawnNpc MapNum, MapNpcNum
    End If
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To MapData(MapNum).NpcCount
        SpawnNpc i, MapNum
    Next
End Sub

Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAPS
        SpawnMapNpcs (i)
    Next
End Sub

Function CanNpcAttackPlayer(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim NpcNum As Long
Dim X As Long, Y As Long

    CanNpcAttackPlayer = False
    
    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Function
    If MapNpcNum <= 0 Then Exit Function
    If MapNpcNum > MapData(MapNum).NpcCount Then Exit Function
    
    NpcNum = MapData(MapNum).MapNpc(MapNpcNum).Num
    
    If NpcNum <= 0 Then Exit Function
    
    ' Make sure the npc isn't already dead
    If MapData(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Function
    
    ' Doesn't matter if they are dead
    If Current_IsDead(Index) Then Exit Function
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapData(MapNum).MapNpc(MapNpcNum).AttackTimer + 1000 Then Exit Function
    
    ' Make sure we dont attack the player if they are switching maps
    If Player(Index).GettingMap = 1 Then Exit Function
    
    ' Make sure they are on the same map

    ' Check if at same coordinates
    Select Case MapData(MapNum).MapNpc(MapNpcNum).Dir
        Case DIR_UP
            Y = Current_Y(Index) + 1
            X = Current_X(Index)
        Case DIR_DOWN
            Y = Current_Y(Index) - 1
            X = Current_X(Index)
        Case DIR_LEFT
            Y = Current_Y(Index)
            X = Current_X(Index) + 1
        Case DIR_RIGHT
            Y = Current_Y(Index)
            X = Current_X(Index) - 1
    End Select
    
    If X = MapData(MapNum).MapNpc(MapNpcNum).X Then
        If Y = MapData(MapNum).MapNpc(MapNpcNum).Y Then
            CanNpcAttackPlayer = True
            MapData(MapNum).MapNpc(MapNpcNum).AttackTimer = GetTickCount
        End If
    End If

End Function

Function NpcCheckDirection(ByVal MapNum As Long, ByVal MapNpcNum As Long, Direction As Byte) As Boolean
Dim i As Long, n As Long
Dim X As Long, Y As Long
    
    NpcCheckDirection = False
    
    Select Case Direction
        Case DIR_UP
            X = MapData(MapNum).MapNpc(MapNpcNum).X
            Y = MapData(MapNum).MapNpc(MapNpcNum).Y - 1
        Case DIR_DOWN
            X = MapData(MapNum).MapNpc(MapNpcNum).X
            Y = MapData(MapNum).MapNpc(MapNpcNum).Y + 1
        Case DIR_LEFT
            X = MapData(MapNum).MapNpc(MapNpcNum).X - 1
            Y = MapData(MapNum).MapNpc(MapNpcNum).Y
        Case DIR_RIGHT
            X = MapData(MapNum).MapNpc(MapNpcNum).X + 1
            Y = MapData(MapNum).MapNpc(MapNpcNum).Y
    End Select
        
    If X < 0 Then
        NpcCheckDirection = True
        Exit Function
    End If
    
    If X > Map(MapNum).MaxX Then
        NpcCheckDirection = True
        Exit Function
    End If
    
    If Y < 0 Then
        NpcCheckDirection = True
        Exit Function
    End If
    
    If Y > Map(MapNum).MaxY Then
        NpcCheckDirection = True
        Exit Function
    End If
    
    n = Map(MapNum).Tile(X, Y).Type
    
    ' Check to make sure that there is not a player in the way
    If PlayerAtPosition(MapNum, X, Y) Then
        NpcCheckDirection = True
        Exit Function
    End If
'    For i = 1 To MAX_PLAYERS
'        If IsPlaying(i) Then
'            If Current_Map(i) = MapNum Then
'                If Current_X(i) = X Then
'                    If Current_Y(i) = Y Then
'                        NpcCheckDirection = True
'                        Exit Function
'                    End If
'                End If
'            End If
'        End If
'    Next
    
    ' Check to make sure that there is not another npc in the way
    For i = 1 To MapData(MapNum).NpcCount
        If i <> MapNpcNum Then
            If MapData(MapNum).MapNpc(i).Num > 0 Then
                If MapData(MapNum).MapNpc(i).X = X Then
                    If MapData(MapNum).MapNpc(i).Y = Y Then
                        NpcCheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    ' Check to make sure that the tile is walkable
    If n <> TILE_TYPE_WALKABLE Then
        If n <> TILE_TYPE_ITEM Then
            If n <> TILE_TYPE_MOBSPAWN Then
                NpcCheckDirection = True
                Exit Function
            End If
        End If
    End If
    
    ' Check for item block
    If n = TILE_TYPE_ITEM Then
        If Map(MapNum).Tile(X, Y).Data3 Then
            NpcCheckDirection = True
            Exit Function
        End If
    End If
End Function

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean

    CanNpcMove = False

    ' Check for subscript out of range
    If MapNum <= 0 Then Exit Function
    If MapNum > MAX_MAPS Then Exit Function
    If MapNpcNum <= 0 Then Exit Function
    If MapNpcNum > MapData(MapNum).NpcCount Then Exit Function
    If Dir < DIR_UP Then Exit Function
    If Dir > DIR_RIGHT Then Exit Function
    
    ' Check to make sure not outside of boundries
    If NpcCheckDirection(MapNum, MapNpcNum, Dir) Then Exit Function
    
    CanNpcMove = True
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
Dim Movement As Byte

    ' Check for subscript out of range
    If MapNum <= 0 Then Exit Sub
    If MapNum > MAX_MAPS Then Exit Sub
    If MapNpcNum <= 0 Then Exit Sub
    If MapNpcNum > MapData(MapNum).NpcCount Then Exit Sub
    If Dir < DIR_UP Then Exit Sub
    If Dir > DIR_RIGHT Then Exit Sub
    
    MapData(MapNum).MapNpc(MapNpcNum).Dir = Dir
    Movement = Npc(MapData(MapNum).MapNpc(MapNpcNum).Num).MovementSpeed
    
    Select Case Dir
        Case DIR_UP
            MapData(MapNum).MapNpc(MapNpcNum).Y = MapData(MapNum).MapNpc(MapNpcNum).Y - 1
               
        Case DIR_DOWN
            MapData(MapNum).MapNpc(MapNpcNum).Y = MapData(MapNum).MapNpc(MapNpcNum).Y + 1
                
        Case DIR_LEFT
            MapData(MapNum).MapNpc(MapNpcNum).X = MapData(MapNum).MapNpc(MapNpcNum).X - 1
                
        Case DIR_RIGHT
            MapData(MapNum).MapNpc(MapNpcNum).X = MapData(MapNum).MapNpc(MapNpcNum).X + 1
    End Select
    
    SendNpcMove MapNum, MapNpcNum, Movement
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)

    ' Check for subscript out of range
    If MapNum <= 0 Then Exit Sub
    If MapNum > MAX_MAPS Then Exit Sub
    If MapNpcNum <= 0 Then Exit Sub
    If MapNpcNum > MapData(MapNum).NpcCount Then Exit Sub
    If Dir < DIR_UP Then Exit Sub
    If Dir > DIR_RIGHT Then Exit Sub
    
    MapData(MapNum).MapNpc(MapNpcNum).Dir = Dir
    SendNpcDir MapNum, MapNpcNum
End Sub

Function NpcInRange(ByVal Index As Long, ByVal MapNpcNum As Byte, ByVal Distance As Byte) As Boolean
Dim DistanceX As Long, DistanceY As Long

    NpcInRange = False
    
    If MapData(Current_Map(Index)).MapNpc(MapNpcNum).Num <= 0 Then Exit Function
    
    DistanceX = MapData(Current_Map(Index)).MapNpc(MapNpcNum).X - Current_X(Index)
    DistanceY = MapData(Current_Map(Index)).MapNpc(MapNpcNum).Y - Current_Y(Index)
    
    ' Make sure we get a positive value
    If DistanceX < 0 Then DistanceX = -DistanceX
    If DistanceY < 0 Then DistanceY = -DistanceY
    
    ' Are they in range?
    If DistanceX <= Distance Then
        If DistanceY <= Distance Then
            NpcInRange = True
        End If
    End If
End Function

' //////////////////////
' // CLASS FUNCTIONS  //
' //////////////////////

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + (GetClassStat(ClassNum, Stats.Strength) \ 2) + GetClassStat(ClassNum, Stats.Strength)) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + (GetClassStat(ClassNum, Stats.Wisdom) \ 2) + GetClassStat(ClassNum, Stats.Wisdom)) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + (GetClassStat(ClassNum, Stats.Dexterity) \ 2) + GetClassStat(ClassNum, Stats.Dexterity)) * 2
End Function

Public Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Function GetGuildName(ByVal Index As Long) As String
    GetGuildName = Trim$(Guild(Index).GuildName)
End Function

Function GetGuildAbbreviation(ByVal Index As Long) As String
    If Index > 0 Then
        GetGuildAbbreviation = Trim$(Guild(Index).GuildAbbreviation)
    End If
End Function

Function GetGuildGMOTD(ByVal Index As Long) As String
    GetGuildGMOTD = Trim$(Guild(Index).GMOTD)
End Function

Function GetGuildOwner(ByVal Index As Long) As String
    GetGuildOwner = Trim$(Guild(Index).Owner)
End Function

Function GetGuildRank(ByVal Index As Long, ByVal GuildRank As Long) As String
    GetGuildRank = Trim$(Guild(Index).Rank(GuildRank))
End Function

Sub SetGuildName(ByVal Index As Long, ByVal GuildName As Long)
    Guild(Index).GuildName = GuildName
End Sub

Sub SetGuildAbbreviation(ByVal Index As Long, ByVal GuildAbbreviation As String)
    Guild(Index).GuildAbbreviation = GuildAbbreviation
End Sub

Sub SetGuildGMOTD(ByVal Index As Long, ByVal Text As Long)
    Guild(Index).GMOTD = Text
End Sub

Sub SetGuildOwner(ByVal Index As Long, ByVal GuildOwner As Long)
    Guild(Index).Owner = GuildOwner
End Sub

Sub SetGuildRank(ByVal Index As Long, ByVal GuildRankNum As Long, ByVal GuildRank As Long)
    Guild(Index).Rank(GuildRankNum) = GuildRank
End Sub

Public Function ItemCount() As Long
Dim i As Long
Dim n As Long

    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) <> vbNullString Then
            n = n + 1
        End If
    Next
    ItemCount = n
End Function

Public Function NpcCount() As Long
Dim i As Long
Dim n As Long

    For i = 1 To MAX_NPCS
        If Trim$(Npc(i).Name) <> vbNullString Then
            n = n + 1
        End If
    Next
    NpcCount = n
End Function

Public Function EmoticonCount() As Long
Dim i As Long
Dim n As Long

    For i = 1 To MAX_EMOTICONS
        If Len(Emoticons(i).Command) > 0 Then
            n = n + 1
        End If
    Next
    EmoticonCount = n
End Function

Public Function ShopCount() As Long
Dim i As Long
Dim n As Long

    For i = 1 To MAX_SHOPS
        If Trim$(Shop(i).Name) <> vbNullString Then
            n = n + 1
        End If
    Next
    ShopCount = n
End Function

Public Function SpellCount() As Long
Dim i As Long
Dim n As Long

    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).Name) <> vbNullString Then
            n = n + 1
        End If
    Next
    SpellCount = n
End Function

Public Function AnimationCount() As Long
Dim i As Long
Dim n As Long

    For i = 1 To MAX_ANIMATIONS
        If Trim$(Animation(i).Name) <> vbNullString Then
            n = n + 1
        End If
    Next
    AnimationCount = n
End Function

Public Sub UpdateMapNpcs()
Dim MapNum As Long
    For MapNum = 1 To MAX_MAPS
        UpdateMapNpc MapNum
    Next
End Sub

Public Sub UpdateMapNpc(ByVal MapNum As Long)
Dim i As Long
Dim n As Long
Dim X As Long

    ' Saves the map npc count for the map
    With MapData(MapNum)
        .NpcCount = 0
        For i = 1 To MAX_MOBS
            .NpcCount = .NpcCount + Map(MapNum).Mobs(i).NpcCount
        Next
        
        ReDim .Npc(.NpcCount)
        ReDim .MapNpc(.NpcCount)
        
        For i = 1 To MAX_MOBS
            For n = 1 To Map(MapNum).Mobs(i).NpcCount
                X = X + 1
                .Npc(X).Npc = Map(MapNum).Mobs(i).Npc(n)
                .Npc(X).MobGroup = i
            Next
        Next
    End With
End Sub

