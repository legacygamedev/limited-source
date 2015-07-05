Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

' Version constants
Public Const CLIENT_MAJOR = 0
Public Const CLIENT_MINOR = 0
Public Const CLIENT_REVISION = 1
' Before and including this version have an old updater.
Public Const CLIENT_MAJOR_UPDATE = 0
Public Const CLIENT_MINOR_UPDATE = 0
Public Const CLIENT_REVISION_UPDATE = 0

'Compression class
Public objCompression As New cCompression

' Used for respawning items
Public SpawnSeconds As Long

'Used for debug window
Public showDebug As Boolean

' Used for weather effects
Public GameWeather As Long
Public WeatherSeconds As Long
Public GameTime As Long
Public TimeSeconds As Long

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveNPCHPTimer As Long
Public PoisonTimer As Long

' Used for logging
Public ServerLog As Boolean

' Used for day or night
Public blnNight As Boolean

'used to determin if you can use (') broadcasts
Public blnBroadcast As Boolean

Sub InitServer()
Dim IPMask As String
Dim i As Long
Dim f As Long
    Call initColours
    Randomize Timer
    
    'Load the port and game name!
    LoadInitData
    
    
    blnBroadcast = False
    
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = "After Darkness Server" & vbNullChar
    ' Add to the sys tray
    Call Shell_NotifyIcon(NIM_ADD, nid)
    
    ' Init atmosphere
    GameWeather = WEATHER_NONE
    WeatherSeconds = 0
    GameTime = TIME_DAY
    TimeSeconds = 0
    
    ' Check if the maps directory is there, if its not make it
    If LCase(Dir(App.Path & "\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\maps")
    End If
    
    ' Check if the accounts directory is there, if its not make it
    If LCase(Dir(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir(App.Path & "\accounts")
    End If
    
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    UPDATER_END_CHAR = END_CHAR
    
    ServerLog = False
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT
    'frmServer.UpdaterSocket(0).LocalPort = GAME_UPDATE_PORT
        
    ' Init all the player sockets
    For i = 1 To MAX_PLAYERS
        Call SetStatus("Initializing player array...")
        Call ClearPlayer(i)
        
        Load frmServer.Socket(i)
    Next i
    
    ' Init all the player sockets
   ' For i = 1 To MAX_PLAYERS
    '    Call SetStatus("Initializing update array...")
   '     Load frmServer.UpdaterSocket(i)
   ' Next i
    
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
    Call SetStatus("Loading paths...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading prayers...")
    Call LoadPrayers
    Call SetStatus("Loading guilds...")
    Call LoadGuilds
    Call SetStatus("Loading Signs...")
    Call LoadSigns
    Call SetStatus("Loading Quests...")
    Call LoadQuests
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    
        
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f
        Close #f
    End If
    
    ' Start listening
    frmServer.Socket(0).Listen
    'frmServer.UpdaterSocket(0).Listen
    
    
    Call UpdateCaption
    
    frmLoad.Visible = False
    frmServer.Show
    
    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
End Sub

Sub DestroyServer()
Dim i As Long

    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = "After Darkness Server" & vbNullChar
    ' Add to the sys tray
    Call Shell_NotifyIcon(NIM_DELETE, nid)
 
    frmLoad.Visible = True
    frmServer.Visible = False
    
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
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
    Call SetStatus("Unloading sockets and timers...")
    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next i

    End
End Sub

Sub SetStatus(ByVal status As String)
    frmLoad.lblStatus.Caption = status
End Sub

Sub ServerLogic()
Dim i As Long

    ' Check for disconnections
    For i = 1 To MAX_PLAYERS
        If frmServer.Socket(i).State > 7 Then
            Call CloseSocket(i)
        End If
    Next i
        
    Call CheckGiveHP
    Call GameAI
    Call CheckPoison
End Sub

Sub CheckSpawnMapItems()
Dim x As Long, y As Long

    ' Used for map item respawning
    SpawnSeconds = SpawnSeconds + 1
    
    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    If SpawnSeconds >= 120 Then
        ' 2 minutes have passed
        For y = 1 To MAX_MAPS
            ' Make sure no one is on the map when it respawns
            If PlayersOnMap(y) = False Then
                ' Clear out unnecessary junk
                For x = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(x, y)
                Next x
                If map(y).Respawn = True Then
                    ' Spawn the items
                    Call SpawnMapItems(y)
                    Call SendMapItemsToAll(y)
                End If
            End If
            DoEvents
        Next y
        
        SpawnSeconds = 0
    End If
End Sub

Sub GameAI()
Dim i As Long, x As Long, y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long, target As Long
Dim didwalk As Boolean
            
    'WeatherSeconds = WeatherSeconds + 1
    TimeSeconds = TimeSeconds + 1
    
    ' Lets change the weather if its time to
    If WeatherSeconds >= 60 Then
        i = Int(Rnd * 3)
        If i <> GameWeather Then
            GameWeather = i
            Call SendWeatherToAll
        End If
        WeatherSeconds = 0
    End If
    
    ' Check if we need to switch from day to night or night to day
    If TimeSeconds >= 3600 Then
        If GameTime = TIME_DAY Then
            GameTime = TIME_NIGHT
            SendNight (True)
        Else
            GameTime = TIME_DAY
            SendNight (False)
        End If
        
        Call SendTimeToAll
        TimeSeconds = 0
    End If
            
    For y = 1 To MAX_MAPS
        If PlayersOnMap(y) = YES Then
            TickCount = GetTickCount
            
            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            If TickCount > TempTile(y).DoorTimer + 5000 Then
                For y1 = 0 To MAX_MAPY
                    For x1 = 0 To MAX_MAPX
                        If map(y).Tile(x1, y1).type = TILE_TYPE_KEY And TempTile(y).DoorOpen(x1, y1) = YES Then
                            TempTile(y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(y, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If
                    Next x1
                Next y1
            End If
            
            For x = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(y, x).num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If map(y).Npc(x) > 0 And MapNpc(y, x).num > 0 Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = y And MapNpc(y, x).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITER Then
                                    n = Npc(NpcNum).Range
                                    
                                    DistanceX = MapNpc(y, x).x - GetPlayerX(i)
                                    DistanceY = MapNpc(y, x).y - GetPlayerY(i)
                                    
                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                                                Call PlayerMsg(i, "A " & Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), RGB_SayColor)
                                            End If
                                            
                                            MapNpc(y, x).target = i
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
                                                                        
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If map(y).Npc(x) > 0 And MapNpc(y, x).num > 0 Then
                    target = MapNpc(y, x).target
                    
                    ' Check to see if its time for the npc to walk
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        ' Check to see if we are following a player or not
                        If target > 0 Then
                            'MapNpc(y, x).TargetRunAway = 0
                            ' Check if the player is even playing, if so follow'm
                            If IsPlaying(target) And GetPlayerMap(target) = y Then
                                didwalk = False
                                
                                i = Int(Rnd * 5)
                                
                                ' Lets move the npc
                                Select Case i
                                    Case 0
                                        ' Up
                                        If MapNpc(y, x).y > GetPlayerY(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, x).y < GetPlayerY(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, x).x > GetPlayerX(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, x).x < GetPlayerX(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                    
                                    Case 1
                                        ' Right
                                        If MapNpc(y, x).x < GetPlayerX(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, x).x > GetPlayerX(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, x).y < GetPlayerY(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, x).y > GetPlayerY(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        
                                    Case 2
                                        ' Down
                                        If MapNpc(y, x).y < GetPlayerY(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, x).y > GetPlayerY(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, x).x < GetPlayerX(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, x).x > GetPlayerX(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                    
                                    Case 3
                                        ' Left
                                        If MapNpc(y, x).x > GetPlayerX(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, x).x < GetPlayerX(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, x).y > GetPlayerY(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, x).y < GetPlayerY(target) And didwalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                didwalk = True
                                            End If
                                        End If
                                End Select
                                
                                
                            
                                ' Check if we can't move and if player is behind something and if we can just switch dirs
                                If Not didwalk Then
                                    If MapNpc(y, x).x - 1 = GetPlayerX(target) And MapNpc(y, x).y = GetPlayerY(target) Then
                                        If MapNpc(y, x).Dir <> DIR_LEFT Then
                                            Call NpcDir(y, x, DIR_LEFT)
                                        End If
                                        didwalk = True
                                    End If
                                    If MapNpc(y, x).x + 1 = GetPlayerX(target) And MapNpc(y, x).y = GetPlayerY(target) Then
                                        If MapNpc(y, x).Dir <> DIR_RIGHT Then
                                            Call NpcDir(y, x, DIR_RIGHT)
                                        End If
                                        didwalk = True
                                    End If
                                    If MapNpc(y, x).x = GetPlayerX(target) And MapNpc(y, x).y - 1 = GetPlayerY(target) Then
                                        If MapNpc(y, x).Dir <> DIR_UP Then
                                            Call NpcDir(y, x, DIR_UP)
                                        End If
                                        didwalk = True
                                    End If
                                    If MapNpc(y, x).x = GetPlayerX(target) And MapNpc(y, x).y + 1 = GetPlayerY(target) Then
                                        If MapNpc(y, x).Dir <> DIR_DOWN Then
                                            Call NpcDir(y, x, DIR_DOWN)
                                        End If
                                        didwalk = True
                                    End If
                                    
                                    ' We could not move so player must be behind something, walk randomly.
                                    If Not didwalk Then
                                        i = Int(Rnd * 2)
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
                                            If CanNpcMove(y, x, i) Then
                                                Call NpcMove(y, x, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                MapNpc(y, x).target = 0
                            End If
                        Else
                            i = Int(Rnd * 4)
                            If i = 1 Then
                                i = Int(Rnd * 4)
                                If CanNpcMove(y, x, i) Then
                                    Call NpcMove(y, x, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If map(y).Npc(x) > 0 And MapNpc(y, x).num > 0 Then
                    target = MapNpc(y, x).target
                    
                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                        ' Is the target playing and on the same map?
                        If IsPlaying(target) And GetPlayerMap(target) = y Then
                            ' Can the npc attack the player?
                            If CanNpcAttackPlayer(x, target) Then
                                If Not CanPlayerBlockHit(target) Then
                                        Damage = Npc(NpcNum).str - GetPlayerProtection(target)
                                        If Damage > 0 Then
                                            Call NpcAttackPlayer(x, target, Damage)
                                        Else
                                            Call PlayerMsg(target, "The " & Trim(Npc(NpcNum).Name) & "'s hit didn't even phase you!", RGB_HelpColor)
                                        End If
                                Else
                                    Call PlayerMsg(target, "Your " & Trim(Item(GetPlayerInvItemNum(target, GetPlayerShieldSlot(target))).Name) & " blocks the " & Trim(Npc(NpcNum).Name) & "'s hit!", RGB_HelpColor)
                                End If
                            End If
                        Else
                            ' Player left map or game, set target to 0
                            MapNpc(y, x).target = 0
                        End If
                    End If
                    
                End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(y, x).num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(y, x).HP > 0 Then
                        MapNpc(y, x).HP = MapNpc(y, x).HP + GetNpcHPRegen(NpcNum)
                    
                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(y, x).HP > GetNpcMaxHP(NpcNum) Then
                            MapNpc(y, x).HP = GetNpcMaxHP(NpcNum)
                        End If
                    End If
                End If
                    
                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(y, x).num = 0 And map(y).Npc(x) > 0 Then
                    If TickCount > MapNpc(y, x).SpawnWait + (Npc(map(y).Npc(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, y)
                    End If
                End If
            Next x
        End If
        DoEvents
    Next y
    
    'Call movePets
    
    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If
End Sub

Sub CheckGiveHP()
Dim i As Long, n As Long

    If GetTickCount > GiveHPTimer + 10000 Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call SetPlayerHP(i, GetPlayerHP(i) + GetPlayerHPRegen(i))
                If player(i).Char(player(i).CharNum).lastSentHP <> GetPlayerHP(i) Then
                    SendHP (i)
                    player(i).Char(player(i).CharNum).lastSentHP = GetPlayerHP(i)
                End If
                Call SetPlayerMP(i, GetPlayerMP(i) + GetPlayerMPRegen(i))
                If player(i).Char(player(i).CharNum).lastSentMP <> GetPlayerMP(i) Then
                    SendMP (i)
                    player(i).Char(player(i).CharNum).lastSentMP = GetPlayerMP(i)
                End If
                Call SetPlayerSP(i, GetPlayerSP(i) + GetPlayerSPRegen(i))
                If player(i).Char(player(i).CharNum).lastSentSP <> GetPlayerSP(i) Then
                    SendSP (i)
                    player(i).Char(player(i).CharNum).lastSentSP = GetPlayerSP(i)
                End If
            End If
            DoEvents
        Next i
        
        GiveHPTimer = GetTickCount
    End If
End Sub

Sub CheckPoison()
    Dim i As Long, n As Long

    If GetTickCount > PoisonTimer + 10000 Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If player(i).Char(player(i).CharNum).Poison = True Then
                    Call SetPlayerHP(i, GetPlayerHP(i) - player(i).Char(player(i).CharNum).Poison_vital)
                    Call SetPlayerPoisonLength(i, GetPlayerPoisonLength(i) - 1)
                    Call SendHP(i)
                End If
            End If
            DoEvents
        Next i
        
        PoisonTimer = GetTickCount
    End If
End Sub

Sub PlayerSaveTimer()
Static MinPassed As Long
Dim i As Long

    MinPassed = MinPassed + 1
    If MinPassed >= 10 Then
        If TotalOnlinePlayers > 0 Then
            Call TextAdd(frmServer.txtText, "Saving all online players...", True)
            Call GlobalMsg("Saving all online players...", Blue)
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    Call SavePlayer(i, False)
                End If
                DoEvents
            Next i
        End If
        
        MinPassed = 0
    End If
End Sub

