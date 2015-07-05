Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

' Version constants
Public Const CLIENT_MAJOR = 3
Public Const CLIENT_MINOR = 0
Public Const CLIENT_REVISION = 3

' Used for respawning items
Public SpawnSeconds As Long

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

' Used for logging
Public ServerLog As Boolean

Sub InitServer()
Dim IPMask As String
Dim i As Long
Dim f As Long
    
    Randomize Timer
    
    nid.cbSize = Len(nid)
    nid.hwnd = frmServer.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = GAME_NAME & " Server" & vbNullChar
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
    
    ServerLog = False
    
    GAME_PORT = ReadIniValue(App.Path & "\Core Files\Configuration.ini", "server information", "PORT")
    GAME_NAME = ReadIniValue(App.Path & "\Core Files\Configuration.ini", "game information", "game name")
    GAME_WEBSITE = ReadIniValue(App.Path & "\Core Files\Configuration.ini", "game information", "game website")
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT
        
    ' Init all the player sockets
    For i = 1 To MAX_PLAYERS
        Call SetStatus("Initializing player array...")
        Call ClearPlayer(i)
        
        Load frmServer.Socket(i)
    Next i
    
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
    Call SetStatus("Loading classes...")
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
    
    Call UpdateCaption
    
    frmLoad.Visible = False
    frmServer.Show
    
    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
End Sub

Sub DestroyServer()
Dim i As Long

    nid.cbSize = Len(nid)
    nid.hwnd = frmServer.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = GAME_NAME & " Server" & vbNullChar
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

Sub SetStatus(ByVal Status As String)
    frmLoad.lblStatus.Caption = Status
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
End Sub

Sub CheckSpawnMapItems()
Dim X As Long, y As Long

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
                For X = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(X, y)
                Next X
                    
                ' Spawn the items
                Call SpawnMapItems(y)
                Call SendMapItemsToAll(y)
            End If
            DoEvents
        Next y
        
        SpawnSeconds = 0
    End If
End Sub

Sub GameAI()
Dim i As Long, X As Long, y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long, Target As Long
Dim DidWalk As Boolean
            
    'WeatherSeconds = WeatherSeconds + 1
    'TimeSeconds = TimeSeconds + 1
    
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
    If TimeSeconds >= 60 Then
        If GameTime = TIME_DAY Then
            GameTime = TIME_NIGHT
        Else
            GameTime = TIME_DAY
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
                        If Map(y).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(y).DoorOpen(x1, y1) = YES Then
                            TempTile(y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(y, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If
                    Next x1
                Next y1
            End If
            
            For X = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(y, X).Num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(X) > 0 And MapNpc(y, X).Num > 0 Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = y And MapNpc(y, X).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITER Then
                                    n = Npc(NpcNum).Range
                                    
                                    DistanceX = MapNpc(y, X).X - GetPlayerX(i)
                                    DistanceY = MapNpc(y, X).y - GetPlayerY(i)
                                    
                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                                                Call PlayerMsg(i, "A " & Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), SayColor)
                                            End If
                                            
                                            MapNpc(y, X).Target = i
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
                If Map(y).Npc(X) > 0 And MapNpc(y, X).Num > 0 Then
                    Target = MapNpc(y, X).Target
                    
                    ' Check to see if its time for the npc to walk
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                            ' Check if the player is even playing, if so follow'm
                            If IsPlaying(Target) And GetPlayerMap(Target) = y Then
                                DidWalk = False
                                
                                i = Int(Rnd * 5)
                                
                                ' Lets move the npc
                                Select Case i
                                    Case 0
                                        ' Up
                                        If MapNpc(y, X).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_UP) Then
                                                Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, X).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_DOWN) Then
                                                Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_LEFT) Then
                                                Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_RIGHT) Then
                                                Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 1
                                        ' Right
                                        If MapNpc(y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_RIGHT) Then
                                                Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_LEFT) Then
                                                Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, X).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_DOWN) Then
                                                Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, X).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_UP) Then
                                                Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        
                                    Case 2
                                        ' Down
                                        If MapNpc(y, X).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_DOWN) Then
                                                Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, X).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_UP) Then
                                                Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_RIGHT) Then
                                                Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_LEFT) Then
                                                Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 3
                                        ' Left
                                        If MapNpc(y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_LEFT) Then
                                                Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_RIGHT) Then
                                                Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, X).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_UP) Then
                                                Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, X).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_DOWN) Then
                                                Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                End Select
                                
                                
                            
                                ' Check if we can't move and if player is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(y, X).X - 1 = GetPlayerX(Target) And MapNpc(y, X).y = GetPlayerY(Target) Then
                                        If MapNpc(y, X).Dir <> DIR_LEFT Then
                                            Call NpcDir(y, X, DIR_LEFT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, X).X + 1 = GetPlayerX(Target) And MapNpc(y, X).y = GetPlayerY(Target) Then
                                        If MapNpc(y, X).Dir <> DIR_RIGHT Then
                                            Call NpcDir(y, X, DIR_RIGHT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, X).X = GetPlayerX(Target) And MapNpc(y, X).y - 1 = GetPlayerY(Target) Then
                                        If MapNpc(y, X).Dir <> DIR_UP Then
                                            Call NpcDir(y, X, DIR_UP)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, X).X = GetPlayerX(Target) And MapNpc(y, X).y + 1 = GetPlayerY(Target) Then
                                        If MapNpc(y, X).Dir <> DIR_DOWN Then
                                            Call NpcDir(y, X, DIR_DOWN)
                                        End If
                                        DidWalk = True
                                    End If
                                    
                                    ' We could not move so player must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
                                            If CanNpcMove(y, X, i) Then
                                                Call NpcMove(y, X, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                MapNpc(y, X).Target = 0
                            End If
                        Else
                            i = Int(Rnd * 4)
                            If i = 1 Then
                                i = Int(Rnd * 4)
                                If CanNpcMove(y, X, i) Then
                                    Call NpcMove(y, X, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(X) > 0 And MapNpc(y, X).Num > 0 Then
                    Target = MapNpc(y, X).Target
                    
                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                        ' Is the target playing and on the same map?
                        If IsPlaying(Target) And GetPlayerMap(Target) = y Then
                            ' Can the npc attack the player?
                            If CanNpcAttackPlayer(X, Target) Then
                                If Not CanPlayerBlockHit(Target) Then
                                    Damage = Npc(NpcNum).STR - GetPlayerProtection(Target)
                                    If Damage > 0 Then
                                        Call NpcAttackPlayer(X, Target, Damage)
                                    Else
                                        Call PlayerMsg(Target, "The " & Trim(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                    End If
                                Else
                                    Call PlayerMsg(Target, "Your " & Trim(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                End If
                            End If
                        Else
                            ' Player left map or game, set target to 0
                            MapNpc(y, X).Target = 0
                        End If
                    End If
                End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(y, X).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(y, X).HP > 0 Then
                        MapNpc(y, X).HP = MapNpc(y, X).HP + GetNpcHPRegen(NpcNum)
                    
                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(y, X).HP > GetNpcMaxHP(NpcNum) Then
                            MapNpc(y, X).HP = GetNpcMaxHP(NpcNum)
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
                If MapNpc(y, X).Num = 0 And Map(y).Npc(X) > 0 Then
                    If TickCount > MapNpc(y, X).SpawnWait + (Npc(Map(y).Npc(X)).SpawnSecs * 1000) Then
                        Call SpawnNpc(X, y)
                    End If
                End If
            Next X
        End If
        DoEvents
    Next y
    
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
                Call SendHP(i)
                Call SetPlayerMP(i, GetPlayerMP(i) + GetPlayerMPRegen(i))
                Call SendMP(i)
                Call SetPlayerSP(i, GetPlayerSP(i) + GetPlayerSPRegen(i))
                Call SendSP(i)
            End If
            DoEvents
        Next i
        
        GiveHPTimer = GetTickCount
    End If
End Sub

Sub PlayerSaveTimer()
Static MinPassed As Long
Dim i As Long

    MinPassed = MinPassed + 1
    If MinPassed >= 60 Then
        If TotalOnlinePlayers > 0 Then
            'Call TextAdd(frmServer.txtText, "Saving all online players...", True)
            'Call GlobalMsg("Saving all online players...", Pink)
            'For i = 1 To MAX_PLAYERS
            '    If IsPlaying(i) Then
            '        Call SavePlayer(i)
            '    End If
            '    DoEvents
            'Next i
            PlayerI = 1
            frmServer.PlayerTimer.Enabled = True
            frmServer.tmrPlayerSave.Enabled = False
        End If
       
        MinPassed = 0
    End If
End Sub

