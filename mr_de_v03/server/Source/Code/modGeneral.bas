Attribute VB_Name = "modGeneral"
Option Explicit

Sub Main()
    InitServer
End Sub

Sub InitServer()
Dim startTick As Long

    startTick = GetTickCount

    Randomize Timer                                                             ' Randomizes the system timer

    frmServer.Show

    SetupTray
    
    ' Setup names for enums
    SetupEnumNames
    
    CheckDirectories                                                            ' Checks for paths

    ServerLog = True                                                            ' Flags default logs

    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP                ' Sets up the server ip
    frmServer.Socket(0).LocalPort = GAME_PORT                                   ' Sets up the default port
    frmServer.Socket(0).Listen                                                  ' Start listening

    UpdateCaption                                                               ' Sets the title bar
    
    InitMessages                                                                ' Need to init messages for packets
    EncryptPackets = 1                                                          ' Encrypt packets? - 1 = Yes 0 = No
    If EncryptPackets Then GenerateEncryptionKeys PacketKeys                    ' Set up encryption keys

    SetStatus ("Initialization complete. Server loaded in " & GetTickCount - startTick & "ms.")

    ServerLogic                                                                 ' Starts the server logic loop
End Sub

Private Sub SetupEnumNames()
    ReDim StatName(1 To Stats.Stat_Count)
    StatName(Stats.Strength) = "Strength"
    StatName(Stats.Dexterity) = "Dexterity"
    StatName(Stats.Vitality) = "Vitality"
    StatName(Stats.Intelligence) = "Intelligence"
    StatName(Stats.Wisdom) = "Wisdom"

    ReDim StatAbbreviation(1 To Stats.Stat_Count)
    StatAbbreviation(Stats.Strength) = "STR"
    StatAbbreviation(Stats.Dexterity) = "DEX"
    StatAbbreviation(Stats.Vitality) = "VIT"
    StatAbbreviation(Stats.Intelligence) = "INT"
    StatAbbreviation(Stats.Wisdom) = "WIS"
    
    ReDim VitalName(1 To Vitals.Vital_Count)
    VitalName(Vitals.HP) = "HP"
    VitalName(Vitals.MP) = "MP"
    VitalName(Vitals.SP) = "SP"

    ReDim EquipmentName(1 To Slots.Slot_Count)
    EquipmentName(Slots.Armor) = "Armor"
    EquipmentName(Slots.Weapon) = "Weapon"
    EquipmentName(Slots.Helmet) = "Helmet"
    EquipmentName(Slots.Shield) = "Shield"
End Sub

' Used to make sure we have all the directories
Public Sub CheckDirectories()
Dim f As Long, i As Long
    
    ' First let's set up our data paths
    AccountPath = App.Path & "\Data\accounts"                ' Account Path
    EmoticonPath = App.Path & "\Data\emoticons"              ' Emoticon Path
    GuildPath = App.Path & "\Data\guilds"                    ' Guild Path
    ItemPath = App.Path & "\Data\items"                      ' Item Path
    LogPath = App.Path & "\Data\Logs"                        ' Log Path
    MapPath = App.Path & "\Data\maps"                        ' Map Path
    NpcPath = App.Path & "\Data\npcs"                        ' Npc Path
    ShopPath = App.Path & "\Data\shops"                      ' Shop Path
    SpellPath = App.Path & "\Data\spells"                    ' Spell Path
    AnimationPath = App.Path & "\Data\animations"       ' Animations Path
    QuestPath = App.Path & "\Data\quests"               ' Quests Path
    
    CheckDirectory AccountPath, "accounts"              ' Accounts directory
    CheckDirectory EmoticonPath, "emoticons"            ' Emoticons directory
    CheckDirectory GuildPath, "guilds"                  ' Guilds directory
    CheckDirectory ItemPath, "items"                    ' Items directory
    CheckDirectory LogPath, "logs"                      ' Logs directory
    CheckDirectory MapPath, "maps"                      ' Maps directory
    CheckDirectory NpcPath, "npcs"                      ' Npcs directory
    CheckDirectory ShopPath, "shops"                    ' Shops directory
    CheckDirectory SpellPath, "spells"                  ' Spells directory
    CheckDirectory AnimationPath, "animations"          ' Animations directory
    CheckDirectory QuestPath, "quests"                  ' Quests directory
    
    If Not FileExist(AccountPath & "\charlist.txt", True) Then                  ' Char List file
        f = FreeFile
        Open AccountPath & "\charlist.txt" For Output As #f
        Close #f
    End If
    
    If Not FileExist(GuildPath & "\GuildName.txt", True) Then                   ' Guild Name File
        f = FreeFile
        Open GuildPath & "\GuildName.txt" For Output As #f
        Close #f
    End If
    
    If Not FileExist(GuildPath & "\GuildAbbreviation.txt", True) Then           ' Guild Abbreviation File
        f = FreeFile
        Open GuildPath & "\GuildAbbreviation.txt" For Output As #f
        Close #f
    End If
                                                                                ' Load the MOTD into the variable
    GameMOTD = GetVar(App.Path & "\Data\Core Files\Configuration.ini", "message of the day", "MOTD")
                                                                                ' Load the expmod
    ExpMod = GetVar(App.Path & "\Data\Core Files\Configuration.ini", "server", "ExpMod")
    frmServer.txtExpMod.Text = ExpMod
    
    For i = 1 To MAX_PLAYERS                                                    ' Init all the player sockets
        ClearPlayer i                                                           ' Clear player array
        
        Load frmServer.Socket(i)                                                ' load sockets
    Next
    
    Party_Clear_All
    
    SetStatus ("Loading classes...")
    ClearClasses
    LoadClasses
    
    SetStatus ("Loading maps...")
    ClearMaps
    ClearTempTiles
    ClearMapItems
    ClearMapNpcs
    LoadMaps
    
    SetStatus ("Loading items...")
    ClearItems
    LoadItems
    
    SetStatus ("Loading npcs...")
    ClearNpcs
    LoadNpcs
    
    SetStatus ("Loading shops...")
    ClearShops
    LoadShops
    
    SetStatus ("Loading spells...")
    ClearSpells
    LoadSpells
    
    SetStatus ("Loading emoticons...")
    ClearEmos
    LoadEmos
    
    SetStatus ("Loading guilds...")
    ClearGuilds
    LoadGuilds
    
    SetStatus ("Loading animations...")
    ClearAnimations
    LoadAnimations
    
    SetStatus ("Loading quests...")
    ClearQuests
    LoadQuests
    
    SetStatus "Updating Npc Quest Lists..."
    Update_Npc_Quests
    
    SetStatus ("Spawning map items...")
    SpawnAllMapsItems
    
    SetStatus ("Spawning map npcs...")
    SpawnAllMapNpcs
    
    ' Now we set all our UDT sizes
    AnimationSize = LenB(Animation(1))
    EmoticonSize = LenB(Emoticons(i))
    ItemSize = LenB(Item(1))
    NpcSize = LenB(Npc(1))
    ShopSize = LenB(Shop(1))
    SpellSize = LenB(Spell(1))
    QuestSize = LenB(Quest(1))
    
    SetStatus "Caching maps..."
    CacheMaps
    
    SetStatus "Caching Items..."
    CacheItems
    
    SetStatus "Caching Npcs..."
    CacheNpcs
    
    SetStatus "Caching Emoticons..."
    CacheEmoticons
    
    SetStatus "Caching Shops..."
    CacheShops
    
    SetStatus "Caching Spells..."
    CacheSpells
    
    SetStatus "Caching Animations..."
    CacheAnimations
    
    SetStatus "Caching Quests..."
    CacheQuests
End Sub

Sub CheckDirectory(ByVal Path As String, ByVal PathName As String)
    If LCase$(Dir$(Path, vbDirectory)) <> PathName Then
        MkDir Path
    End If
End Sub

Sub DestroyServer()
Dim i As Long

    ServerOnline = 0
    
    DestroyTray
 
    SetStatus ("Saving players online...")
    SaveAllPlayersOnline
        
    SetStatus ("Clearing maps...")
    ClearMaps
    SetStatus ("Clearing map items...")
    ClearMapItems
    SetStatus ("Clearing map npcs...")
    ClearMapNpcs
    SetStatus ("Clearing npcs...")
    ClearNpcs
    SetStatus ("Clearing items...")
    ClearItems
    SetStatus ("Clearing shops...")
    ClearShops
    SetStatus ("Clearing emoticons...")
    ClearEmos
    SetStatus ("Clearing animations...")
    ClearAnimations
    SetStatus ("Clearing guilds...")
    ClearGuilds
    SetStatus ("Unloading sockets and timers...")
    
    For i = 1 To MAX_PLAYERS
        Set Player(i).Buffer = Nothing
        Unload frmServer.Socket(i)
    Next

    End
End Sub

Sub ServerLogic()
Dim i As Long
Dim StartTickCount As Long
Dim Elapsed As Long
Dim LastUpdateNpc As Long
Dim LastUpdateMapDeleteItems As Long
Dim LastUpdateMapSpawnItems As Long
Dim LastDisconnectCheck As Long
Dim LastUpdatePlayer As Long

    ServerOnline = 1
    
    Do While ServerOnline
        
        StartTickCount = GetTickCount
        
        '/////////////////////////////////////////////
        '// Checks if it's time to update something //
        '/////////////////////////////////////////////
        
        ' Check for disconnections every second
        If StartTickCount > LastDisconnectCheck Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    CloseSocket i
                End If
            Next
            LastDisconnectCheck = GetTickCount + 1000
        End If
        
        ' Updates players every half second
        If StartTickCount > LastUpdatePlayer Then
            For i = 1 To OnlinePlayersCount
                OnUpdate OnlinePlayers(i)
            Next
'            For i = 1 To MAX_PLAYERS
'                If IsPlaying(i) Then
'                    OnUpdate i
'                End If
'            Next
            LastUpdatePlayer = GetTickCount + 500
        End If
        
        ' Update npcs every second
        If StartTickCount > LastUpdateNpc Then
            UpdateNpcAI
            LastUpdateNpc = GetTickCount + 500
        End If
        
        ' Checks to delete map items every 10 seconds - this isn't a time critical event
        If StartTickCount > LastUpdateMapDeleteItems Then
            UpdateMapDeleteItems
            LastUpdateMapDeleteItems = GetTickCount + 60000
        End If
        
        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If StartTickCount > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If
        
        ' So windows doesn't freak out
        DoEvents
        
        'Check if we have enough time to sleep (stop the process from running completely, freeing up the CPU)
        Elapsed = GetTickCount - StartTickCount
        If Elapsed < 5 Then
            If Elapsed >= 0 Then    'Make sure nothing weird happens, causing for a huge sleep time
                Sleep Int(5 - Elapsed)
            End If
        End If
    Loop
End Sub

Sub UpdateMapDeleteItems()
Dim MapNum As Long, MapItemNum As Long, TickCount As Long

    TickCount = GetTickCount
    
    For MapNum = 1 To MAX_MAPS
        For MapItemNum = 1 To MAX_MAP_ITEMS
            If MapData(MapNum).MapItem(MapItemNum).DropTime > 0 Then
                ' Should be 5 minutes...
                If TickCount > MapData(MapNum).MapItem(MapItemNum).DropTime + 300000 Then
                    SpawnItemSlot MapItemNum, 0, 0, MapNum, MapData(MapNum).MapItem(MapItemNum).X, MapData(MapNum).MapItem(MapItemNum).Y, 0

                    ' Erase item from the map
                    MapData(MapNum).MapItem(MapItemNum).Num = 0
                    MapData(MapNum).MapItem(MapItemNum).Value = 0
                    MapData(MapNum).MapItem(MapItemNum).X = 0
                    MapData(MapNum).MapItem(MapItemNum).Y = 0
                    MapData(MapNum).MapItem(MapItemNum).DropTime = 0
                End If
            End If
        Next
    Next
End Sub

Sub UpdateMapSpawnItems()
Dim MapNum As Long, MapItemNum As Long
    
    For MapNum = 1 To MAX_MAPS
         ' Check to spawn map items
        If MapData(MapNum).MapPlayersCount = 0 Then
            ' Clear out unnecessary junk
            For MapItemNum = 1 To MAX_MAP_ITEMS
                ClearMapItem MapNum, MapItemNum
            Next

            ' Spawn the items
            SpawnMapItems MapNum
            'SendMapItemsToAll MapNum
        End If
    Next
End Sub
   
' Used for updating NPCS. Walking, Attacking, HP regain. Also resets doors.
Sub UpdateNpcAI()
Dim MapNum As Long, MapNpcNum As Long
Dim TickCount As Long

    For MapNum = 1 To MAX_MAPS
    
        TickCount = GetTickCount
        
        ' /////////////////////////////////////////
        ' // This is used for updating map data  //
        ' /////////////////////////////////////////
        UpdateMapData TickCount, MapNum
        
        ' /////////////////////////////////////////
        ' // This is used for updating map npcs  //
        ' /////////////////////////////////////////
        For MapNpcNum = 1 To MapData(MapNum).NpcCount
            If MapData(MapNum).MapNpc(MapNpcNum).Num Then
                ' Update status effects
                MapNpc_Update MapNum, MapNpcNum
                
                ' Check if the map has players
                If MapData(MapNum).MapPlayersCount Then
                    ' Update movement and such
                    UpdateNpc TickCount, MapNum, MapNpcNum
                End If
            ElseIf MapData(MapNum).MapNpc(MapNpcNum).Num = 0 Then
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapData(MapNum).Npc(MapNpcNum).Npc Then
                    If TickCount > MapData(MapNum).MapNpc(MapNpcNum).SpawnWait Then
                        SpawnNpc MapNpcNum, MapNum
                    End If
                End If
            End If
        Next
    Next
End Sub

Sub UpdateMapData(ByRef TickCount As Long, ByRef MapNum As Long)
Dim X As Long, X2 As Long
Dim Y As Long, Y2 As Long

    ' ////////////////////////////////////
    ' // This is used for closing doors //
    ' ////////////////////////////////////
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
        
            Select Case Map(MapNum).Tile(X, Y).Type
                Case TILE_TYPE_KEY
                    If TickCount > MapData(MapNum).TempTile.DoorTimer(X, Y) Then
                        If MapData(MapNum).TempTile.DoorOpen(X, Y) Then
                            If Not PlayerAtPosition(MapNum, X, Y) Then
                                MapData(MapNum).TempTile.DoorOpen(X, Y) = False
                                SendMapKey MapNum, X, Y, 0
                            End If
                        End If
                    End If
                    
                Case TILE_TYPE_KEYOPEN
                    X2 = Map(MapNum).Tile(X, Y).Data1
                    Y2 = Map(MapNum).Tile(X, Y).Data2
                    If Map(MapNum).Tile(X, Y).Data3 > 0 Then
                        If PlayerAtPosition(MapNum, X, Y) Then
                            If Not MapData(MapNum).TempTile.DoorOpen(X2, Y2) Then
                                MapData(MapNum).TempTile.DoorOpen(X2, Y2) = True
                                SendMapKey MapNum, X2, Y2, 1
                            End If
                        Else
                            If MapData(MapNum).TempTile.DoorOpen(X2, Y2) Then
                                If Not PlayerAtPosition(MapNum, X2, Y2) Then
                                    MapData(MapNum).TempTile.DoorOpen(X2, Y2) = False
                                    SendMapKey MapNum, X2, Y2, 0
                                End If
                            End If
                        End If
                    End If
            End Select
            
        Next
    Next
End Sub

Sub UpdateNpc(ByRef TickCount As Long, ByRef MapNum As Long, ByVal MapNpcNum As Long)
Dim i As Long, n As Long
Dim NpcNum As Long
Dim Target As Long
Dim X As Long, Y As Long

    NpcNum = MapData(MapNum).MapNpc(MapNpcNum).Num
    
    ' Make sure theres a npc with the map
    If NpcNum > 0 Then
        
        Target = MapData(MapNum).MapNpc(MapNpcNum).Target
        
        ' /////////////////////////////////////////
        ' // This is used for ATTACKING ON SIGHT //
        ' /////////////////////////////////////////
        ' If the npc is a attack on sight, search for a player on the map
        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Then
            If Target = 0 Then
                For i = 1 To MapData(MapNum).MapPlayersCount
                    If Current_Access(MapData(MapNum).MapPlayers(i)) <= ADMIN_MONITER Then
                        If Not Current_IsDead(MapData(MapNum).MapPlayers(i)) Then
                            n = Npc(NpcNum).Range

                            X = MapData(MapNum).MapNpc(MapNpcNum).X - Current_X(MapData(MapNum).MapPlayers(i))
                            Y = MapData(MapNum).MapNpc(MapNpcNum).Y - Current_Y(MapData(MapNum).MapPlayers(i))

                            ' Make sure we get a positive value
                            If X < 0 Then X = -X
                            If Y < 0 Then Y = -Y

                            ' Are they in range?  if so GET'M!
                            If X <= n Then
                                If Y <= n Then
                                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Current_PK(MapData(MapNum).MapPlayers(i)) Then
                                        If Trim$(Npc(NpcNum).AttackSay) <> vbNullString Then
                                            SendPlayerMsg i, Trim$(Npc(NpcNum).Name) & ": " & Trim$(Npc(NpcNum).AttackSay), SayColor
                                        End If
    
                                        MapData(MapNum).MapNpc(MapNpcNum).Target = MapData(MapNum).MapPlayers(i)
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        
        ' /////////////////////////////////////////////
        ' // This is used for NPC walking/targetting //
        ' /////////////////////////////////////////////
        MapNpc_Move MapNum, MapNpcNum, NpcNum, Target
    End If
End Sub

' *****************************************************************
' *** Below get / set UDT data with byte arrays and copy memory ***
' *****************************************************************
'
' Animations
'
Public Function Get_AnimationData(ByRef AnimationNum As Long) As Byte()
Dim AnimationData() As Byte
    ReDim AnimationData(0 To AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Get_AnimationData = AnimationData
End Function

Public Sub Set_AnimationData(ByRef AnimationNum As Long, ByRef AnimationData() As Byte)
    CopyMemory ByVal VarPtr(Animation(AnimationNum)), ByVal VarPtr(AnimationData(0)), AnimationSize
End Sub

'
' Emoticons
'
Public Function Get_EmoticonData(ByRef EmoticonNum As Long) As Byte()
Dim EmoticonData() As Byte
    ReDim EmoticonData(0 To EmoticonSize - 1)
    CopyMemory EmoticonData(0), ByVal VarPtr(Emoticons(EmoticonNum)), EmoticonSize
    Get_EmoticonData = EmoticonData
End Function

Public Sub Set_EmoticonData(ByRef EmoticonNum As Long, ByRef EmoticonData() As Byte)
    CopyMemory ByVal VarPtr(Emoticons(EmoticonNum)), ByVal VarPtr(EmoticonData(0)), EmoticonSize
End Sub

'
' Items
'
Public Function Get_ItemData(ByRef ItemNum As Long) As Byte()
Dim ItemData() As Byte
    ReDim ItemData(0 To ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Get_ItemData = ItemData
End Function

Public Sub Set_ItemData(ByRef ItemNum As Long, ByRef ItemData() As Byte)
    CopyMemory ByVal VarPtr(Item(ItemNum)), ByVal VarPtr(ItemData(0)), ItemSize
End Sub

'
' Npcs
'
Public Function Get_NpcData(ByRef NpcNum As Long) As Byte()
Dim NpcData() As Byte
    ReDim NpcData(0 To NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    Get_NpcData = NpcData
End Function

Public Sub Set_NpcData(ByRef NpcNum As Long, ByRef NpcData() As Byte)
    CopyMemory ByVal VarPtr(Npc(NpcNum)), ByVal VarPtr(NpcData(0)), NpcSize
End Sub

'
' Shops
'
Public Function Get_ShopData(ByRef ShopNum As Long) As Byte()
Dim ShopData() As Byte
    ReDim ShopData(0 To ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    Get_ShopData = ShopData
End Function

Public Sub Set_ShopData(ByRef ShopNum As Long, ByRef ShopData() As Byte)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize
End Sub

'
' Spells
'
Public Function Get_SpellData(ByRef SpellNum As Long) As Byte()
Dim SpellData() As Byte
    ReDim SpellData(0 To SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    Get_SpellData = SpellData
End Function

Public Sub Set_SpellData(ByRef SpellNum As Long, ByRef SpellData() As Byte)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
End Sub

' ************************************************************
' *** Below will cache all data into one packet for logins ***
' ************************************************************

Public Sub CacheMaps()
Dim i As Long
    For i = 1 To MAX_MAPS
        CacheMap i
    Next
End Sub

Public Sub CacheMap(ByVal MapNum As Long)
Dim Buffer As clsBuffer
Dim Data As clsBuffer
Dim X As Long
Dim Y As Long
Dim TileSize As Long
Dim TileData() As Byte

    Set Buffer = New clsBuffer
    Set Data = New clsBuffer
        
    Data.WriteLong MapNum
    With Map(MapNum)
        Data.WriteString .Name
        Data.WriteLong .Revision
        Data.WriteByte .Moral
        Data.WriteInteger .Up
        Data.WriteInteger .Down
        Data.WriteInteger .Left
        Data.WriteInteger .Right
        Data.WriteByte .Music
        Data.WriteInteger .BootMap
        Data.WriteByte .BootX
        Data.WriteByte .BootY
        Data.WriteByte .TileSet
        Data.WriteByte .MaxX
        Data.WriteByte .MaxY
        
        For X = 1 To MAX_MOBS
            Data.WriteLong .Mobs(X).NpcCount
            If .Mobs(X).NpcCount > 0 Then
                For Y = 1 To .Mobs(X).NpcCount
                    Data.WriteLong .Mobs(X).Npc(Y)
                Next
            End If
        Next
        
        TileSize = LenB(.Tile(0, 0)) * ((UBound(.Tile, 1) + 1) * (UBound(.Tile, 2) + 1))
        ReDim TileData(0 To TileSize - 1)
        CopyMemory TileData(0), ByVal VarPtr(.Tile(0, 0)), TileSize
        Data.WriteBytes TileData
    End With
    
    Data.CompressBuffer
        
    Buffer.PreAllocate Data.Length + 4
    Buffer.WriteLong CMsgMapData
    Buffer.WriteBytes Data.ToArray()
        
    MapCache(MapNum).Data() = Buffer.ToArray()
End Sub

Public Sub CacheItems()
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer

    Buffer.PreAllocate ((ItemSize + 4) * ItemCount) + 4
    Buffer.WriteLong ItemCount      ' Item Count
    
    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) <> vbNullString Then
            Buffer.WriteLong i
            Buffer.WriteBytes Get_ItemData(i)
        End If
    Next
        
    Buffer.CompressBuffer
    
    ItemsCache() = Buffer.ToArray()
End Sub

Public Sub CacheNpcs()
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong NpcCount       ' Npc Count
    
    For i = 1 To MAX_NPCS
        If Trim$(Npc(i).Name) <> vbNullString Then
            Buffer.WriteLong i
            Buffer.WriteString Trim$(Npc(i).Name)
            Buffer.WriteInteger Npc(i).Sprite
            Buffer.WriteByte Npc(i).Behavior
            Buffer.WriteByte Npc(i).MovementSpeed
            Buffer.WriteByte Npc(i).MovementFrequency
        End If
    Next
        
    Buffer.CompressBuffer
    
    NpcsCache() = Buffer.ToArray()
End Sub

Public Sub CacheEmoticons()
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer

    Buffer.PreAllocate ((EmoticonSize + 4) * EmoticonCount) + 4
    Buffer.WriteLong EmoticonCount
    
    For i = 1 To MAX_EMOTICONS
        If Len(Emoticons(i).Command) > 0 Then
            Buffer.WriteLong i
            Buffer.WriteBytes Get_EmoticonData(i)
        End If
    Next
        
    Buffer.CompressBuffer
    
    EmoticonsCache() = Buffer.ToArray()
End Sub

Public Sub CacheShops()
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer

    Buffer.PreAllocate ((ShopSize + 4) * ShopCount) + 4
    Buffer.WriteLong ShopCount
    
    For i = 1 To MAX_SHOPS
        If Trim$(Shop(i).Name) <> vbNullString Then
            Buffer.WriteLong i
            Buffer.WriteBytes Get_ShopData(i)
        End If
    Next
        
    Buffer.CompressBuffer
    
    ShopsCache() = Buffer.ToArray()
End Sub

Public Sub CacheSpells()
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer

    Buffer.PreAllocate ((SpellSize + 4) * SpellCount) + 4
    Buffer.WriteLong SpellCount
    
    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).Name) <> vbNullString Then
            Buffer.WriteLong i
            Buffer.WriteBytes Get_SpellData(i)
        End If
    Next
        
    Buffer.CompressBuffer
    
    SpellsCache() = Buffer.ToArray()
End Sub

Public Sub CacheAnimations()
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer

    Buffer.PreAllocate ((AnimationSize + 4) * AnimationCount) + 4
    Buffer.WriteLong AnimationCount
    
    For i = 1 To MAX_ANIMATIONS
        If Trim$(Animation(i).Name) <> vbNullString Then
            Buffer.WriteLong i
            Buffer.WriteBytes Get_AnimationData(i)
        End If
    Next
    
    Buffer.CompressBuffer
    
    AnimationsCache() = Buffer.ToArray()
End Sub
