Attribute VB_Name = "modGeneral"
Option Explicit

Sub InitServer()
    Dim Index As Integer

    Call SetStatus("Checking Folders...")

    ' Check folders
    If Not FolderExists(App.Path & "\Maps") Then
        Call MkDir(App.Path & "\Maps")
    End If

    If Not FolderExists(App.Path & "\Logs") Then
        Call MkDir(App.Path & "\Logs")
    End If

    If Not FolderExists(App.Path & "\Accounts") Then
        Call MkDir(App.Path & "\Accounts")
    End If

    If Not FolderExists(App.Path & "\NPCs") Then
        Call MkDir(App.Path & "\NPCs")
    End If

    If Not FolderExists(App.Path & "\Items") Then
        Call MkDir(App.Path & "\Items")
    End If

    If Not FolderExists(App.Path & "\Spells") Then
        Call MkDir(App.Path & "\Spells")
    End If

    If Not FolderExists(App.Path & "\Shops") Then
        Call MkDir(App.Path & "\Shops")
    End If

    If Not FolderExists(App.Path & "\Banks") Then
        Call MkDir(App.Path & "\Banks")
    End If
    
    If Not FolderExists(App.Path & "\Classes") Then
        Call MkDir(App.Path & "\Classes")
    End If
    
    If Not FolderExists(App.Path & "\Mail") Then
        Call MkDir(App.Path & "\Mail")
    End If
    
    If Not FolderExists(App.Path & "\Mail\Outbox") Then
        Call MkDir(App.Path & "\Mail\Outbox")
    End If
    
    Call SetStatus("Checking Files...")

    If Not FileExists("Data.ini") Then
        PutVar App.Path & "\Data.ini", "CONFIG", "GameName", "Eclipse Stable"
        PutVar App.Path & "\Data.ini", "CONFIG", "GameVersion", "1.0"
        PutVar App.Path & "\Data.ini", "CONFIG", "WebSite", vbNullString
        PutVar App.Path & "\Data.ini", "CONFIG", "Port", 4000
        PutVar App.Path & "\Data.ini", "CONFIG", "HPRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "HPTimer", 1000
        PutVar App.Path & "\Data.ini", "CONFIG", "MPRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "MPTimer", 1000
        PutVar App.Path & "\Data.ini", "CONFIG", "SPRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "SPTimer", 1000
        PutVar App.Path & "\Data.ini", "CONFIG", "NPCRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "Stat1", "Strength"
        PutVar App.Path & "\Data.ini", "CONFIG", "Stat2", "Defense"
        PutVar App.Path & "\Data.ini", "CONFIG", "Stat3", "Speed"
        PutVar App.Path & "\Data.ini", "CONFIG", "Stat4", "Magic"
        PutVar App.Path & "\Data.ini", "CONFIG", "PlayerCard", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "Scrolling", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "ScrollX", 30
        PutVar App.Path & "\Data.ini", "CONFIG", "ScrollY", 30
        PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "ScriptErrors", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "PaperDoll", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "SaveTime", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "SpriteSize", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "Custom", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "Level", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "PKMinLvl", 10
        PutVar App.Path & "\Data.ini", "CONFIG", "Email", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "VerifyAcc", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "Classes", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "SPAttack", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "SPRunning", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "WalkFix", 0
        
        
        PutVar App.Path & "\Data.ini", "MAX", "MAX_PLAYERS", 50
        PutVar App.Path & "\Data.ini", "MAX", "MAX_ITEMS", 100
        PutVar App.Path & "\Data.ini", "MAX", "MAX_NPCS", 100
        PutVar App.Path & "\Data.ini", "MAX", "MAX_SHOPS", 100
        PutVar App.Path & "\Data.ini", "MAX", "MAX_SPELLS", 100
        PutVar App.Path & "\Data.ini", "MAX", "MAX_MAPS", 50
        PutVar App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS", 20
        PutVar App.Path & "\Data.ini", "MAX", "MAX_GUILDS", 20
        PutVar App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS", 10
        PutVar App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS", 10
        PutVar App.Path & "\Data.ini", "MAX", "MAX_LEVEL", 500
        PutVar App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS", 4
        PutVar App.Path & "\Data.ini", "MAX", "MAX_ELEMENTS", 20
        PutVar App.Path & "\Data.ini", "MAX", "MAX_SCRIPTSPELLS", 500
    End If

    If Not FileExists("Stats.ini") Then
        PutVar App.Path & "\Stats.ini", "HP", "AddPerLevel", 10
        PutVar App.Path & "\Stats.ini", "HP", "AddPerStr", 10
        PutVar App.Path & "\Stats.ini", "HP", "AddPerDef", 0
        PutVar App.Path & "\Stats.ini", "HP", "AddPerMagi", 0
        PutVar App.Path & "\Stats.ini", "HP", "AddPerSpeed", 0
        PutVar App.Path & "\Stats.ini", "MP", "AddPerLevel", 10
        PutVar App.Path & "\Stats.ini", "MP", "AddPerStr", 0
        PutVar App.Path & "\Stats.ini", "MP", "AddPerDef", 0
        PutVar App.Path & "\Stats.ini", "MP", "AddPerMagi", 10
        PutVar App.Path & "\Stats.ini", "MP", "AddPerSpeed", 0
        PutVar App.Path & "\Stats.ini", "SP", "AddPerLevel", 10
        PutVar App.Path & "\Stats.ini", "SP", "AddPerStr", 0
        PutVar App.Path & "\Stats.ini", "SP", "AddPerDef", 0
        PutVar App.Path & "\Stats.ini", "SP", "AddPerMagi", 0
        PutVar App.Path & "\Stats.ini", "SP", "AddPerSpeed", 20
    End If

    If Not FileExists("News.ini") Then
        PutVar App.Path & "\News.ini", "DATA", "NewsTitle", "Change this message in News.ini."
        PutVar App.Path & "\News.ini", "DATA", "NewsBody", "Change this message in News.ini."
        PutVar App.Path & "\News.ini", "COLOR", "Red", 255
        PutVar App.Path & "\News.ini", "COLOR", "Green", 255
        PutVar App.Path & "\News.ini", "COLOR", "Blue", 255
    End If
    
    If Not FileExists("MOTD.ini") Then
        PutVar App.Path & "\MOTD.ini", "MOTD", "Msg", "Change this message in MOTD.ini."
    End If

    If Not FileExists("Tiles.ini") Then
        For Index = 0 To 100
            PutVar App.Path & "\Tiles.ini", "Names", "Tile" & Index, CStr(Index)
        Next Index
    End If

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExists("Accounts\Charlist.txt") Then
        Index = FreeFile
        Open App.Path & "\Accounts\CharList.txt" For Output As #Index
        Close #Index
    End If

    Call SetStatus("Loading Settings...")

    On Error GoTo LoadErr
    addHP.LEVEL = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerLevel"))
    addHP.STR = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerStr"))
    addHP.DEF = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerDef"))
    addHP.Magi = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerMagi"))
    addHP.Speed = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerSpeed"))
    addMP.LEVEL = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerLevel"))
    addMP.STR = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerStr"))
    addMP.DEF = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerDef"))
    addMP.Magi = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerMagi"))
    addMP.Speed = Val(GetVar(App.Path & "\Stats.ini", "MP", vbNullString))
    addSP.LEVEL = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerLevel"))
    addSP.STR = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerStr"))
    addSP.DEF = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerDef"))
    addSP.Magi = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerMagi"))
    addSP.Speed = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerSpeed"))

    GAME_NAME = GetVar(App.Path & "\Data.ini", "CONFIG", "GameName")
    GAME_PORT = GetVar(App.Path & "\Data.ini", "CONFIG", "Port")
    MAX_PLAYERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_PLAYERS")
    MAX_ITEMS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_ITEMS")
    MAX_NPCS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_NPCS")
    MAX_SHOPS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SHOPS")
    MAX_SPELLS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPELLS")
    MAX_MAPS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAPS")
    MAX_MAP_ITEMS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS")
    MAX_GUILDS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILDS")
    MAX_GUILD_MEMBERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS")
    MAX_EMOTICONS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS")
    MAX_LEVEL = GetVar(App.Path & "\Data.ini", "MAX", "MAX_LEVEL")
    MAX_PARTY_MEMBERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS")
    MAX_ELEMENTS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_ELEMENTS")
    MAX_SCRIPTSPELLS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SCRIPTSPELLS")
    SCRIPTING = GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting")
    PAPERDOLL = GetVar(App.Path & "\Data.ini", "CONFIG", "PaperDoll")
    SPRITESIZE = GetVar(App.Path & "\Data.ini", "CONFIG", "SpriteSize")
    HP_REGEN = GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen")
    HP_TIMER = GetVar(App.Path & "\Data.ini", "CONFIG", "HPTimer")
    MP_REGEN = GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen")
    MP_TIMER = GetVar(App.Path & "\Data.ini", "CONFIG", "MPTimer")
    SP_REGEN = GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen")
    SP_TIMER = GetVar(App.Path & "\Data.ini", "CONFIG", "SPTimer")
    NPC_REGEN = GetVar(App.Path & "\Data.ini", "CONFIG", "NPCRegen")
    STAT1 = GetVar(App.Path & "\Data.ini", "CONFIG", "Stat1")
    STAT2 = GetVar(App.Path & "\Data.ini", "CONFIG", "Stat2")
    STAT3 = GetVar(App.Path & "\Data.ini", "CONFIG", "Stat3")
    STAT4 = GetVar(App.Path & "\Data.ini", "CONFIG", "Stat4")
    SP_ATTACK = GetVar(App.Path & "\Data.ini", "CONFIG", "SPAttack")
    SP_RUNNING = GetVar(App.Path & "\Data.ini", "CONFIG", "SPRunning")
    CUSTOM_SPRITE = GetVar(App.Path & "\Data.ini", "CONFIG", "Custom")
    EMAIL_AUTH = GetVar(App.Path & "\Data.ini", "CONFIG", "Email")
    SAVETIME = GetVar(App.Path & "\Data.ini", "CONFIG", "SaveTime")
    LEVEL = GetVar(App.Path & "\Data.ini", "CONFIG", "Level")
    PKMINLVL = GetVar(App.Path & "\Data.ini", "CONFIG", "PKMinLvl")
    ACC_VERIFY = GetVar(App.Path & "\Data.ini", "CONFIG", "VerifyAcc")
    CLASSES = GetVar(App.Path & "\Data.ini", "CONFIG", "Classes")

    If GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling") = 0 Then
        IS_SCROLLING = 0
        MAX_MAPX = 19
        MAX_MAPY = 14
    Else
        IS_SCROLLING = 1
        MAX_MAPX = GetVar(App.Path & "\Data.ini", "CONFIG", "ScrollX")
        MAX_MAPY = GetVar(App.Path & "\Data.ini", "CONFIG", "ScrollY")
    End If

    ' Weather variables.
    WeatherType = WEATHER_NONE
    WeatherLevel = 25

    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)

    ServerLog = True
    
    GoTo LoadSuccess

LoadErr:
    Call MsgBox("Error reading from Data.ini or Stats.ini.", vbOKOnly)
    End

LoadSuccess:
    ' Restore error handling
    On Error GoTo 0

    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim MapCache(1 To MAX_MAPS) As String
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim Item(0 To MAX_ITEMS) As ItemRec
    ReDim NPC(0 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNPC(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim Element(0 To MAX_ELEMENTS) As ElementRec

    For Index = 1 To MAX_GUILDS
        ReDim Guild(Index).Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
    Next Index
    
    For Index = 1 To MAX_MAPS
        ReDim Map(Index).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim TempTile(Index).DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
    Next Index

    ReDim Experience(1 To MAX_LEVEL) As Long

    START_MAP = 1
    START_X = MAX_MAPX / 2
    START_Y = MAX_MAPY / 2

    Set CTimers = New Collection

    Call IncrementBar

    On Error GoTo ScriptErr

    ' Scripting
    frmServer.lblScriptOn.caption = "Scripts: OFF"
    
    ' Check for main.ess
    If Not FileExists("\Scripts\main.ess") Then
        Call MsgBox("main.ess not found. Scripts disabled.", vbExclamation)
        SCRIPTING = 0
    End If
    
    ' Continue
    If SCRIPTING = 1 Then
        Call SetStatus("Loading scripts...")
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\main.ess", "Scripts\main.ess", MyScript.SControl
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        frmServer.lblScriptOn.caption = "Scripts: ON"
    End If

    Call IncrementBar

    GoTo ScriptsGood

ScriptErr:
    Call MsgBox("Error loading the scripting engine.", vbOKOnly)
    End

ScriptsGood:

    On Error GoTo 0

    ' Get the listening socket ready to go
    Set GameServer = New clsServer

    ' Init all the player sockets
    Call SetStatus("Initializing player array...")
    For Index = 1 To MAX_PLAYERS
        Call ClearPlayer(Index)

        Call GameServer.Sockets.Add(CStr(Index))
    Next Index

    For Index = 1 To MAX_PLAYERS
        Call ShowPLR(Index)
    Next Index

    Call IncrementBar

    frmServer.lstTopics.Clear
    frmServer.lstTopics.AddItem "Starting Up"
    frmServer.lstTopics.AddItem "Editing Your Server"
    frmServer.lstTopics.AddItem "Player Access"
    frmServer.lstTopics.AddItem "Running A Server"
    frmServer.lstTopics.AddItem "Player Controls"
    frmServer.lstTopics.AddItem "Player Commands"
    frmServer.lstTopics.AddItem "Chatting"
    frmServer.lstTopics.AddItem "Bugs/Errors"
    frmServer.lstTopics.AddItem "Map Convertor"
    frmServer.lstTopics.AddItem "Map Editing"
    frmServer.lstTopics.AddItem "Scripting Commands"
    frmServer.lstTopics.AddItem "Questions?"
    frmServer.lstTopics.AddItem "New Features"
    frmServer.lstTopics.Selected(0) = True

    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTile
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing exp...")
    Call ClearExperience
    Call SetStatus("Clearing emoticons...")
    Call ClearEmoticon
    Call IncrementBar
    Call SetStatus("Loading emoticons...")
    Call IncrementBar
    Call LoadEmoticon
    Call SetStatus("Loading elements...")
    Call IncrementBar
    Call LoadElements
    Call SetStatus("Clearing arrows...")
    Call ClearArrows
    Call SetStatus("Loading arrows...")
    Call IncrementBar
    Call LoadArrows
    Call SetStatus("Loading exp...")
    Call IncrementBar
    Call LoadExperience
    Call SetStatus("Loading classes...")
    Call IncrementBar
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call IncrementBar
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call IncrementBar
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call IncrementBar
    Call LoadNpcs
    Call SetStatus("Loading shops...")
    Call IncrementBar
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call IncrementBar
    Call LoadSpells
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call IncrementBar

    frmServer.MapList.Clear

    For Index = 1 To MAX_MAPS
        frmServer.MapList.AddItem Index & ": " & Map(Index).Name
    Next Index

    frmServer.MapList.Selected(0) = True
    frmServer.tmrPlayerSave.Enabled = True
    frmServer.tmrSpawnMapItems.Enabled = True
    frmServer.Timer1.Enabled = True

    ' Error handling for 'Address in use' error
    Err.Clear
    On Error Resume Next

    ' Start listening
    GameServer.StartListening

    ' RTE 10048 occured
    If Err.Number = 10048 Then
        Call MsgBox("The port on this address is already busy.", vbOKOnly)
        End
    End If
    
    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\main.ess", "OnServerLoad"
    End If

    ' Restore error handling
    On Error GoTo 0

    Call UpdateTitle
    Call UpdateTOP

    frmLoad.Visible = False
    frmServer.Show

    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
    frmServer.tmrScriptedTimer.Enabled = True
End Sub

Sub DestroyServer()
    Dim I As Long
    
    Call SaveAllPlayersOnline

    frmServer.Visible = False
    frmLoad.Visible = True

    For I = 1 To MAX_PLAYERS
        temp = I / MAX_PLAYERS * 100
        Call SetStatus("Unloading Sockets... " & temp & "%")
        Call GameServer.Sockets.Remove(CStr(I))
    Next I

    Set GameServer = Nothing

    End
End Sub

Sub SetStatus(ByVal Status As String)
    frmLoad.lblStatus.caption = Status
    DoEvents
End Sub

Sub IncrementBar()
    On Error Resume Next
    ' Increment prog bar
    frmLoad.loadProgressBar.Value = frmLoad.loadProgressBar.Value + 1
End Sub

Sub ServerLogic()
    Call CheckGiveVitals
    Call GameAI
    Call ScriptedTimer
End Sub

Sub CheckSpawnMapItems()
    Dim X As Long
    Dim Y As Long

    ' Used for map item respawning
    SpawnSeconds = SpawnSeconds + 1

    ' Respawns the map items.
    If SpawnSeconds >= 120 Then
        ' 2 minutes have passed
        For Y = 1 To MAX_MAPS
            ' Make sure no one is on the map when it respawns
            If Not PlayersOnMap(Y) Then
                ' Clear out unnecessary junk
                For X = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(X, Y)
                Next X

                ' Spawn the items
                Call SpawnMapItems(Y)
                Call SendMapItemsToAll(Y)
            End If
        Next Y

        SpawnSeconds = 0
    End If
End Sub

Public Sub GameAI()
    Dim I As Long, X As Long, Y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
    Dim Damage As Long, DistanceX As Long, DistanceY As Long, NPCnum As Long, Target As Long
    Dim DidWalk As Boolean

    On Error Resume Next

    For Y = 1 To MAX_MAPS
        If PlayersOnMap(Y) = YES Then
            TickCount = GetTickCount
            
            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            If TickCount > TempTile(Y).DoorTimer + 5000 Then
                For y1 = 0 To MAX_MAPY
                    For x1 = 0 To MAX_MAPX
                        If Map(Y).Tile(x1, y1).Type = TILE_TYPE_KEY Then
                            If TempTile(Y).DoorOpen(x1, y1) = YES Then
                                TempTile(Y).DoorOpen(x1, y1) = NO
                                Call SendDataToMap(Y, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                            End If
                        End If
                        
                        If Map(Y).Tile(x1, y1).Type = TILE_TYPE_DOOR Then
                            If TempTile(Y).DoorOpen(x1, y1) = YES Then
                                TempTile(Y).DoorOpen(x1, y1) = NO
                                Call SendDataToMap(Y, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                            End If
                        End If
                    Next x1
                Next y1
            End If
            
            For X = 1 To MAX_MAP_NPCS
                NPCnum = MapNPC(Y, X).num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        ' If the npc is a attack on sight, search for a player on the map
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(NPCnum).Behavior = NPC_BEHAVIOR_GUARD Then
                            For I = 1 To MAX_PLAYERS
                                If IsPlaying(I) Then
                                    If GetPlayerMap(I) = Y Then
                                        If MapNPC(Y, X).Target = 0 Then
                                            If GetPlayerAccess(I) <= ADMIN_MONITER Then
                                                n = NPC(NPCnum).Range
                                                
                                                DistanceX = MapNPC(Y, X).X - GetPlayerX(I)
                                                DistanceY = MapNPC(Y, X).Y - GetPlayerY(I)
                                                
                                                ' Make sure we get a positive value
                                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                If DistanceY < 0 Then DistanceY = DistanceY * -1
                                                
                                                ' Are they in range?  if so GET'M!
                                                If DistanceX <= n Then
                                                    If DistanceY <= n Then
                                                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(I) = YES Then
                                                            If Trim$(NPC(NPCnum).AttackSay) <> vbNullString Then
                                                                Call PlayerMsg(I, "A " & Trim$(NPC(NPCnum).Name) & " : " & Trim$(NPC(NPCnum).AttackSay), SayColor)
                                                            End If
                                                            
                                                            MapNPC(Y, X).Target = I
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next I
                        End If
                    End If
                End If
                                                                        
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        Target = MapNPC(Y, X).Target
                        
                        ' Check to see if its time for the npc to walk
                        If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            ' Check to see if we are following a player or not
                            If Target > 0 Then
                                ' Check if the player is even playing, if so follow'm
                                If IsPlaying(Target) Then
                                    If GetPlayerMap(Target) = Y Then
                                        DidWalk = False
                                        
                                        I = Int(Rnd * 4)
                                        
                                        ' Lets move the npc
                                        Select Case I
                                            Case 0
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                            
                                            Case 1
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                
                                            Case 2
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                            
                                            Case 3
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                        End Select
                                    
                                        ' Check if we can't move and if player is behind something and if we can just switch dirs
                                        If Not DidWalk Then
                                            If MapNPC(Y, X).X - 1 = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_LEFT Then
                                                        Call NpcDir(Y, X, DIR_LEFT)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            If MapNPC(Y, X).X + 1 = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_RIGHT Then
                                                        Call NpcDir(Y, X, DIR_RIGHT)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            If MapNPC(Y, X).X = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y - 1 = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_UP Then
                                                        Call NpcDir(Y, X, DIR_UP)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            If MapNPC(Y, X).X = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y + 1 = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_DOWN Then
                                                        Call NpcDir(Y, X, DIR_DOWN)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            
                                            ' We could not move so player must be behind something, walk randomly.
                                            If Not DidWalk Then
                                                I = Int(Rnd * 2)
                                                If I = 1 Then
                                                    I = Int(Rnd * 4)
                                                    If CanNpcMove(Y, X, I) Then
                                                        Call NpcMove(Y, X, I, MOVING_WALKING)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        MapNPC(Y, X).Target = 0
                                    End If
                                End If
                                
                            Else
                                I = Int(Rnd * 4)
                                If I = 1 Then
                                    I = Int(Rnd * 4)
                                    If CanNpcMove(Y, X, I) Then
                                        Call NpcMove(Y, X, I, MOVING_WALKING)
                                    End If
                                End If
                            End If
                            
                        End If
                    End If
                    
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        Target = MapNPC(Y, X).Target
                        
                        ' Check if the npc can attack the targeted player player
                        If Target > 0 Then
                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                ' Can the npc attack the player?
                                If CanNpcAttackPlayer(X, Target) Then
                                    If Not CanPlayerBlockHit(Target) Then
                                    
                                        Damage = NPC(NPCnum).STR - GetPlayerProtection(Target)
                                        
                                        If Damage > 0 Then
                                            If SCRIPTING = 1 Then
                                                MyScript.ExecuteStatement "Scripts\main.ess", "PlayerHit " & Target & "," & X & "," & Damage
                                            Else
                                                Call NpcAttackPlayer(X, Target, Damage)
                                            End If
                                        Else
                                            Call BattleMsg(Target, "The " & Trim$(NPC(NPCnum).Name) & " couldn't hurt you!", BRIGHTBLUE, 1)
                                            
                                            'Call PlayerMsg(Target, "The " & Trim(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                        End If
                                    Else
                                        Call BattleMsg(Target, "You blocked the " & Trim$(NPC(NPCnum).Name) & "'s hit!", BRIGHTCYAN, 1)
                                        
                                        'Call PlayerMsg(Target, "Your " & Trim(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                    End If
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                MapNPC(Y, X).Target = 0
                            End If
                        End If

                    End If
                End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNPC(Y, X).num > 0 Then
                    If TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNPC(Y, X).HP > 0 Then
                            MapNPC(Y, X).HP = MapNPC(Y, X).HP + GetNpcHPRegen(NPCnum)
                        
                            ' Check if they have more then they should and if so just set it to max
                            If MapNPC(Y, X).HP > GetNpcMaxHP(NPCnum) Then
                                MapNPC(Y, X).HP = GetNpcMaxHP(NPCnum)
                            End If
                        End If
                    End If
                End If

                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNPC(Y, X).num = 0 Then
                    If Map(Y).NPC(X) > 0 Then
                        If TickCount > MapNPC(Y, X).SpawnWait + (NPC(Map(Y).NPC(X)).SpawnSecs * 1000) Then
                            Call SpawnNpc(X, Y)
                        End If
                    End If
                End If
            Next X
        End If
    Next Y
    
    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

End Sub


Sub ScriptedTimer()
    Dim X As Long, n As Long
    Dim CustomTimer As clsCTimers

    n = 0
    X = CTimers.Count
    For Each CustomTimer In CTimers
        n = n + 1
        If GetTickCount > CustomTimer.tmrWait Then
            MyScript.ExecuteStatement "Scripts\main.ess", CustomTimer.Name ' & " " & Index & "," & PointType
            If CTimers.Count < X Then
                n = n - X - CTimers.Count
                X = CTimers.Count
            End If
            If n > 0 Then
                CTimers.Item(n).tmrWait = GetTickCount + CustomTimer.Interval
            Else
                Exit For
            End If
        End If
    Next CustomTimer
End Sub

Sub CheckGiveVitals()
    Dim I As Long

    If HP_REGEN = 1 Then
        If GetTickCount >= GiveHPTimer + HP_TIMER Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerHP(I) < GetPlayerMaxHP(I) Then
                        Call SetPlayerHP(I, GetPlayerHP(I) + GetPlayerHPRegen(I))
                        Call SendHP(I)
                    End If
                End If
            Next I

            GiveHPTimer = GetTickCount
        End If
    End If

    If MP_REGEN = 1 Then
        If GetTickCount >= GiveMPTimer + MP_TIMER Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerMP(I) < GetPlayerMaxMP(I) Then
                        Call SetPlayerMP(I, GetPlayerMP(I) + GetPlayerMPRegen(I))
                        Call SendMP(I)
                    End If
                End If
            Next I

            GiveMPTimer = GetTickCount
        End If
    End If

    If SP_REGEN = 1 Then
        If GetTickCount >= GiveSPTimer + SP_TIMER Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerSP(I) < GetPlayerMaxSP(I) Then
                        Call SetPlayerSP(I, GetPlayerSP(I) + GetPlayerSPRegen(I))
                        Call SendSP(I)
                    End If
                End If
            Next I

            GiveSPTimer = GetTickCount
        End If
    End If
End Sub

Sub PlayerSaveTimer()
    Dim I As Long

    PLYRSAVE_TIMER = PLYRSAVE_TIMER + 1

    If SAVETIME <> 0 Then
        If PLYRSAVE_TIMER >= SAVETIME Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    Call SavePlayer(I)
                End If
            Next I
    
            PlayerI = 1

            frmServer.PlayerTimer.Enabled = True
            frmServer.tmrPlayerSave.Enabled = False

            PLYRSAVE_TIMER = 0
        End If
    Else
        PLYRSAVE_TIMER = 0
    End If
End Sub

Function IsAlphaNumeric(TestString As String) As Boolean
    Dim LoopID As Integer
    Dim sChar As String

    IsAlphaNumeric = False

    If LenB(TestString) > 0 Then
        For LoopID = 1 To Len(TestString)
            sChar = Mid$(TestString, LoopID, 1)
            If Not sChar Like "[0-9A-Za-z]" Then
                Exit Function
            End If
        Next

        IsAlphaNumeric = True
    End If
End Function

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function
