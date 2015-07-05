Attribute VB_Name = "modGeneral"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

Public Declare Function GetTickCount _
   Lib "kernel32" () As Long

' Version constants
Public Const CLIENT_MAJOR = 1
Public Const CLIENT_MINOR = 1
Public Const CLIENT_REVISION = 1

' Security password
Public Const SEC_CODE = "89h89hr98hewf9wfnd3nf98b9s8enfs09fn390jnf83n"

' Used for respawning items
Public SpawnSeconds As Long

' Used for weather effects
Public GameWeather As Long
Public WeatherSeconds As Long
Public GameTime As Long
Public TimeSeconds As Long
Public RainIntensity As Long
Public GameClock As String
Public Gamespeed As Long
Public Hours As Integer

Public TimeDisable As Boolean

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean
Public CurrentLoad As Long

Sub CheckGiveHP()
Dim i As Long

    If GetTickCount > GiveHPTimer + 10000 Then
        For i = 1 To MAX_PLAYERS

            If IsPlaying(i) Then
                If GetPlayerHP(i) <= GetPlayerMaxHP(i) And GetPlayerHP(i) > 0 Then
                    Call SetPlayerHP(i, GetPlayerHP(i) + GetPlayerHPRegen(i))
                    Call SendHP(i)
                End If
                If GetPlayerMP(i) <= GetPlayerMaxMP(i) And GetPlayerMP(i) > 0 Then
                    Call SetPlayerMP(i, GetPlayerMP(i) + GetPlayerMPRegen(i))
                    Call SendMP(i)
                End If
                If GetPlayerSP(i) <= GetPlayerMaxSP(i) And GetPlayerSP(i) > 0 Then
                    Call SetPlayerSP(i, GetPlayerSP(i) + GetPlayerSPRegen(i))
                    Call SendSP(i)
                End If
            End If
            DoEvents

        Next
        GiveHPTimer = GetTickCount
    End If
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
                Next

                ' Spawn the items
                Call SpawnMapItems(y)
                Call SendMapItemsToAll(y)
            End If
            DoEvents

        Next
        SpawnSeconds = 0
    End If
End Sub

Sub DestroyServer()
Dim i As Long

    Call Shell_NotifyIcon(NIM_DELETE, nid)
    Call SetStatus("Shutting down...")
    frmLoad.Visible = True
    frmServer.Visible = False
    DoEvents

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
        Call SetStatus("Unloading sockets and timers... " & i & "/" & MAX_PLAYERS)
        DoEvents

        Unload frmServer.Socket(i)
    Next

    'If frmServer.chkChat.value = Checked Then
    '    Call SetStatus("Saving chat logs...")
    '    Call SaveLogs
    'End If
    End
End Sub

Sub GameAI()
Dim i As Long, X As Long, y As Long, N As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, TickCount As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, npcnum As Long, Target As Long
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

                        If Map(y).Tile(x1, y1).Type = TILE_TYPE_DOOR And TempTile(y).DoorOpen(x1, y1) = YES Then
                            TempTile(y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(y, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If
                    Next
                Next
            End If
            For X = 1 To MAX_MAP_NPCS
                npcnum = MapNpc(y, X).num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(X) > 0 And MapNpc(y, X).num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(npcnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(npcnum).Behavior = NPC_BEHAVIOR_GUARD Then
                        For i = 1 To MAX_PLAYERS

                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = y And MapNpc(y, X).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITER Then
                                    N = Npc(npcnum).Range
                                    DistanceX = MapNpc(y, X).X - GetPlayerX(i)
                                    DistanceY = MapNpc(y, X).y - GetPlayerY(i)

                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1

                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= N And DistanceY <= N Then
                                        If Npc(npcnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            If Trim$(Npc(npcnum).AttackSay) <> "" Then
                                                Call PlayerMsg(i, "A " & Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay) & "", SayColor)
                                            End If
                                            MapNpc(y, X).TargetType = TARGET_TYPE_PLAYER
                                            MapNpc(y, X).Target = i
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
                ' Make sure theres a npc with the map
                If Map(y).Npc(X) > 0 And MapNpc(y, X).num > 0 Then
                    Target = MapNpc(y, X).Target

                    ' Check to see if its time for the npc to walk
                    If Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then

                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                            If MapNpc(y, X).TargetType = TARGET_TYPE_PLAYER Then

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

                                ' Check if the pet is even playing, if so follow'm
                                If IsPlaying(Target) And Player(Target).Pet.Map = y Then
                                    DidWalk = False
                                    i = Int(Rnd * 5)

                                    ' Lets move the npc
                                    Select Case i

                                        Case 0

                                            ' Up
                                            If MapNpc(y, X).y > Player(Target).Pet.y And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_UP) Then
                                                    Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNpc(y, X).y < Player(Target).Pet.y And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_DOWN) Then
                                                    Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNpc(y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_LEFT) Then
                                                    Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNpc(y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_RIGHT) Then
                                                    Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 1

                                            ' Right
                                            If MapNpc(y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_RIGHT) Then
                                                    Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNpc(y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_LEFT) Then
                                                    Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNpc(y, X).y < Player(Target).Pet.y And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_DOWN) Then
                                                    Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNpc(y, X).y > Player(Target).Pet.y And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_UP) Then
                                                    Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 2

                                            ' Down
                                            If MapNpc(y, X).y < Player(Target).Pet.y And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_DOWN) Then
                                                    Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNpc(y, X).y > Player(Target).Pet.y And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_UP) Then
                                                    Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNpc(y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_RIGHT) Then
                                                    Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNpc(y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_LEFT) Then
                                                    Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 3

                                            ' Left
                                            If MapNpc(y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_LEFT) Then
                                                    Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNpc(y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_RIGHT) Then
                                                    Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNpc(y, X).y > Player(Target).Pet.y And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_UP) Then
                                                    Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNpc(y, X).y < Player(Target).Pet.y And DidWalk = False Then
                                                If CanNpcMove(y, X, DIR_DOWN) Then
                                                    Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                    End Select

                                    ' Check if we can't move and if pet is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNpc(y, X).X - 1 = Player(Target).Pet.X And MapNpc(y, X).y = Player(Target).Pet.y Then
                                            If MapNpc(y, X).Dir <> DIR_LEFT Then
                                                Call NpcDir(y, X, DIR_LEFT)
                                            End If
                                            DidWalk = True
                                        End If

                                        If MapNpc(y, X).X + 1 = Player(Target).Pet.X And MapNpc(y, X).y = Player(Target).Pet.y Then
                                            If MapNpc(y, X).Dir <> DIR_RIGHT Then
                                                Call NpcDir(y, X, DIR_RIGHT)
                                            End If
                                            DidWalk = True
                                        End If

                                        If MapNpc(y, X).X = Player(Target).Pet.X And MapNpc(y, X).y - 1 = Player(Target).Pet.y Then
                                            If MapNpc(y, X).Dir <> DIR_UP Then
                                                Call NpcDir(y, X, DIR_UP)
                                            End If
                                            DidWalk = True
                                        End If

                                        If MapNpc(y, X).X = Player(Target).Pet.X And MapNpc(y, X).y + 1 = Player(Target).Pet.y Then
                                            If MapNpc(y, X).Dir <> DIR_DOWN Then
                                                Call NpcDir(y, X, DIR_DOWN)
                                            End If
                                            DidWalk = True
                                        End If

                                        ' We could not move so pet must be behind something, walk randomly.
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

                ' //////////////////////////////////////////////////////
                ' // This is used for npcs to attack players and pets //
                ' //////////////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(X) > 0 And MapNpc(y, X).num > 0 Then
                    Target = MapNpc(y, X).Target

                    If MapNpc(y, X).TargetType <> TARGET_TYPE_LOCATION And MapNpc(y, X).TargetType <> TARGET_TYPE_NPC Then

                        ' Check if the npc can attack the targeted player player
                        If Target > 0 Then
                            If MapNpc(y, X).TargetType = TARGET_TYPE_PLAYER Then

                                ' Is the target playing and on the same map?
                                If IsPlaying(Target) And GetPlayerMap(Target) = y Then

                                    ' Can the npc attack the player?
                                    If CanNpcAttackPlayer(X, Target) Then
                                        If Not CanPlayerBlockHit(Target) Then
                                            Damage = Npc(npcnum).STR - GetPlayerProtection(Target) + (Rnd * 5) - 2

                                            If Damage > 0 Then
                                                Call NpcAttackPlayer(X, Target, Damage)
                                            Else
                                                Call BattleMsg(Target, "The " & Trim$(Npc(npcnum).Name) & " couldn't hurt you!", BrightBlue, 1)

                                                'Call PlayerMsg(Target, "The " & Trim$(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                            End If
                                        Else
                                            Call BattleMsg(Target, "You blocked the " & Trim$(Npc(npcnum).Name) & "'s hit!", BrightCyan, 1)

                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                        End If
                                    End If
                                Else

                                    ' Player left map or game, set target to 0
                                    MapNpc(y, X).Target = 0
                                End If
                            Else

                                ' Is the target playing and on the same map?
                                If IsPlaying(Target) And Player(Target).Pet.Map = y Then

                                    ' Can the npc attack the pet?
                                    If CanNpcAttackPet(X, Target) Then
                                        Damage = Npc(npcnum).STR - Player(Target).Pet.Level + (Rnd * 5) - 2

                                        If Damage > 0 Then
                                            Call NpcAttackPet(X, Target, Damage)
                                        End If
                                    End If
                                Else

                                    ' Pet left map or game, set target to 0
                                    MapNpc(y, X).Target = 0
                                End If
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(y, X).num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(y, X).HP > 0 Then
                        MapNpc(y, X).HP = MapNpc(y, X).HP + GetNpcHPRegen(npcnum)

                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(y, X).HP > GetNpcMaxHP(npcnum) Then
                            MapNpc(y, X).HP = GetNpcMaxHP(npcnum)
                        End If
                        Call SendDataToMap(y, "NPCHP" & SEP_CHAR & X & SEP_CHAR & MapNpc(y, X).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(y, X).num) & SEP_CHAR & END_CHAR)
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).str > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(y, X).num = 0 And Map(y).Npc(X) > 0 Then
                    If TickCount > MapNpc(y, X).SpawnWait + (Npc(Map(y).Npc(X)).SpawnSecs * 1000) Then
                        Call SpawnNpc(X, y)
                    End If
                End If

                If MapNpc(y, X).num > 0 Then

                    ' If the NPC hasn't been fighting, why send it's HP?
                    If GetTickCount < MapNpc(y, X).LastAttack + 6000 Then
                        Call SendDataToMap(y, "NPCHP" & SEP_CHAR & X & SEP_CHAR & MapNpc(y, X).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(y, X).num) & SEP_CHAR & END_CHAR)
                    End If
                End If
            Next
        End If
        DoEvents

    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

    ' //////////////////////////////////////////////////////////
    ' // Used for moving pets (it took a while it get going!) //
    ' //////////////////////////////////////////////////////////
    For X = 1 To MAX_PLAYERS
    
    If Player(X).CorpseMap > 0 Then
         If GetTickCount > CLng(Player(X).CorpseTimer + CLng((400000))) Then
          Call ClearCorpse(X)
          Call SendCorpseToAll(X)
         End If
        End If

        If Player(X).Pet.Alive = YES Then
            x1 = Player(X).Pet.X
            y1 = Player(X).Pet.y
            x2 = Player(X).Pet.XToGo
            y2 = Player(X).Pet.YToGo

            If Player(X).Pet.Target > 0 Then
                If Player(X).Pet.TargetType = TARGET_TYPE_PLAYER Then
                    x2 = GetPlayerX(Player(X).Pet.Target)
                    y2 = GetPlayerY(Player(X).Pet.Target)
                End If

                If Player(X).Pet.TargetType = TARGET_TYPE_NPC Then
                    If CanPetAttackNpc(X, Player(X).Pet.Target) Then
                        Damage = Player(X).Pet.Level - Npc(Player(X).Pet.Target).STR + (Rnd * 5) - 2

                        If Damage > 0 Then
                            Call PetAttackNpc(X, Player(X).Pet.Target, Damage)
                            x2 = x1
                            y2 = y1
                        End If
                    Else
                        x2 = MapNpc(Player(X).Pet.Map, Player(X).Pet.Target).X
                        y2 = MapNpc(Player(X).Pet.Map, Player(X).Pet.Target).y
                    End If
                End If
            Else

                If Player(X).Pet.Map = GetPlayerMap(X) Or Player(X).Pet.MapToGo = 0 Then
                    If Player(X).Pet.XToGo = -1 Or Player(X).Pet.YToGo = -1 Then
                        i = Int(Rnd * 4)

                        If i = 1 Then
                            i = Int(Rnd * 4)

                            If i = DIR_UP Then
                                y2 = y1 - 1
                                x2 = Player(X).Pet.X
                            End If

                            If i = DIR_DOWN Then
                                y2 = y1 + 1
                                x2 = Player(X).Pet.X
                            End If

                            If i = DIR_RIGHT Then
                                x2 = x1 + 1
                                y2 = Player(X).Pet.y
                            End If

                            If i = DIR_LEFT Then
                                x2 = x1 - 1
                                y2 = Player(X).Pet.y
                            End If

                            If Not IsValid(x2, y2) Then
                                x2 = x1
                                y2 = y1
                            End If
                            If Grid(Player(X).Pet.Map).Loc(x2, y2).Blocked = True Then
                                x2 = x1
                                y2 = y1
                            End If
                        Else
                            x2 = x1
                            y2 = y1
                        End If
                    End If
                Else

                    If Map(Player(X).Pet.Map).Up = Player(X).Pet.MapToGo Then
                        y2 = y1 - 1
                    Else

                        If Map(Player(X).Pet.Map).Down = Player(X).Pet.MapToGo Then
                            y2 = y1 + 1
                        Else

                            If Map(Player(X).Pet.Map).Left = Player(X).Pet.MapToGo Then
                                x2 = x1 - 1
                            Else

                                If Map(Player(X).Pet.Map).Right = Player(X).Pet.MapToGo Then
                                    x2 = x1 + 1
                                Else
                                    i = Int(Rnd * 4)

                                    If i = 1 Then
                                        i = Int(Rnd * 4)

                                        If i = DIR_UP Then y2 = y1 - 1
                                        If i = DIR_DOWN Then y2 = y1 + 1
                                        If i = DIR_RIGHT Then x2 = x1 + 1
                                        If i = DIR_LEFT Then x2 = x1 - 1
                                        If Not IsValid(x2, y2) Then
                                            x2 = x1
                                            y2 = y1
                                        End If
                                        If Grid(Player(X).Pet.Map).Loc(x2, y2).Blocked = True Then
                                            x2 = x1
                                            y2 = y1
                                        End If
                                    Else
                                        x2 = x1
                                        y2 = y1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If x1 < x2 Then

                ' RIGHT not left
                If y1 < y2 Then

                    ' DOWN not up
                    If x2 - x1 > y2 - y1 Then

                        ' RIGHT not down
                        If CanPetMove(X, DIR_RIGHT) Then

                            ' RIGHT works
                            Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                        Else

                            If CanPetMove(X, DIR_DOWN) Then

                                ' DOWN works and right doesn't
                                Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                            Else

                                ' Nothing works, random time
                                i = Int(Rnd * 4)

                                If CanPetMove(X, i) Then
                                    Call PetMove(X, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    Else

                        If x2 - x1 <> y2 - y1 Then

                            ' DOWN not right
                            If CanPetMove(X, DIR_DOWN) Then

                                ' DOWN works
                                Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                            Else

                                If CanPetMove(X, DIR_RIGHT) Then

                                    ' RIGHT works and down doesn't
                                    Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    i = Int(Rnd * 4)

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        Else

                            ' Both are equal
                            If CanPetMove(X, DIR_RIGHT) Then

                                ' RIGHT works
                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN and RIGHT work
                                    i = (Int(Rnd * 2) * 2) + 1

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If
                                Else

                                    ' RIGHT works only
                                    Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                End If
                            Else

                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN works only
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    i = Int(Rnd * 4)

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else

                    If y1 <> y2 Then

                        ' UP not down
                        If x2 - x1 > y1 - y2 Then

                            ' RIGHT not up
                            If CanPetMove(X, DIR_RIGHT) Then

                                ' RIGHT works
                                Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                            Else

                                If CanPetMove(X, DIR_UP) Then

                                    ' UP works and right doesn't
                                    Call PetMove(X, DIR_UP, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    i = Int(Rnd * 4)

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        Else

                            If x2 - x1 <> y1 - y2 Then

                                ' UP not right
                                If CanPetMove(X, DIR_UP) Then

                                    ' UP works
                                    Call PetMove(X, DIR_UP, MOVING_WALKING)
                                Else

                                    If CanPetMove(X, DIR_RIGHT) Then

                                        ' RIGHT works and up doesn't
                                        Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            Else

                                ' Both are equal
                                If CanPetMove(X, DIR_RIGHT) Then

                                    ' RIGHT works
                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP and RIGHT work
                                        i = Int(Rnd * 2) * 3

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    Else

                                        ' RIGHT works only
                                        Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                    End If
                                Else

                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP works only
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else

                        ' Target is horizontal
                        If CanPetMove(X, DIR_RIGHT) Then

                            ' RIGHT works
                            Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                        Else

                            ' Right doesn't work
                            If CanPetMove(X, DIR_UP) Then
                                If CanPetMove(X, DIR_DOWN) Then

                                    ' UP and DOWN work
                                    i = Int(Rnd * 2)
                                    Call PetMove(X, i, MOVING_WALKING)
                                Else

                                    ' Only UP works
                                    Call PetMove(X, DIR_UP, MOVING_WALKING)
                                End If
                            Else

                                If CanPetMove(X, DIR_DOWN) Then

                                    ' Only DOWN works
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    ' Nothing works, only left is left (heh)
                                    If CanPetMove(X, DIR_LEFT) Then
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    Else

                                        ' Nothing works at all, let it die
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else

                If x1 <> x2 Then

                    ' LEFT not right
                    If y1 < y2 Then

                        ' DOWN not up
                        If x1 - x2 > y2 - y1 Then

                            ' LEFT not down
                            If CanPetMove(X, DIR_LEFT) Then

                                ' LEFT works
                                Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                            Else

                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN works and left doesn't
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    i = Int(Rnd * 4)

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        Else

                            If x1 - x2 <> y2 - y1 Then

                                ' DOWN not left
                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN works
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' LEFT works and down doesn't
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            Else

                                ' Both are equal
                                If CanPetMove(X, DIR_LEFT) Then

                                    ' LEFT works
                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' DOWN and LEFT work
                                        i = Int(Rnd * 2) + 1
                                        Call PetMove(X, i, MOVING_WALKING)
                                    Else

                                        ' LEFT works only
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    End If
                                Else

                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' DOWN works only
                                        Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else

                        If y1 <> y2 Then

                            ' UP not down
                            If x1 - x2 > y1 - y2 Then

                                ' LEFT not up
                                If CanPetMove(X, DIR_LEFT) Then

                                    ' LEFT works
                                    Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                Else

                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP works and left doesn't
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            Else

                                If x1 - x2 <> y1 - y2 Then

                                    ' UP not LEFT
                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP works
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        If CanPetMove(X, DIR_LEFT) Then

                                            ' LEFT works and up doesn't
                                            Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                        Else

                                            ' Nothing works, random time
                                            i = Int(Rnd * 4)

                                            If CanPetMove(X, i) Then
                                                Call PetMove(X, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                Else

                                    ' Both are equal
                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' LEFT works
                                        If CanPetMove(X, DIR_UP) Then

                                            ' UP and LEFT work
                                            i = Int(Rnd * 2) * 2
                                            Call PetMove(X, i, MOVING_WALKING)
                                        Else

                                            ' LEFT works only
                                            Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                        End If
                                    Else

                                        If CanPetMove(X, DIR_UP) Then

                                            ' UP works only
                                            Call PetMove(X, DIR_UP, MOVING_WALKING)
                                        Else

                                            ' Nothing works, random time
                                            i = Int(Rnd * 4)

                                            If CanPetMove(X, i) Then
                                                Call PetMove(X, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else

                            ' Target is horizontal
                            If CanPetMove(X, DIR_LEFT) Then

                                ' LEFT works
                                Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                            Else

                                ' LEFT doesn't work
                                If CanPetMove(X, DIR_UP) Then
                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' UP and DOWN work
                                        i = Int(Rnd * 2)
                                        Call PetMove(X, i, MOVING_WALKING)
                                    Else

                                        ' Only UP works
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    End If
                                Else

                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' Only DOWN works
                                        Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                    Else

                                        ' Nothing works, only right is left (heh)
                                        If CanPetMove(X, DIR_RIGHT) Then
                                            Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                        Else

                                            ' Nothing works at all, let it die
                                            Player(X).Pet.MapToGo = Player(X).Pet.Map
                                            Player(X).Pet.XToGo = -1
                                            Player(X).Pet.YToGo = -1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else

                    ' Target is vertical
                    If y1 < y2 Then

                        ' DOWN not up
                        If CanPetMove(X, DIR_DOWN) Then
                            Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                        Else

                            ' Down doesn't work
                            If CanPetMove(X, DIR_RIGHT) Then
                                If CanPetMove(X, DIR_LEFT) Then

                                    ' RIGHT and LEFT work
                                    i = Int((Rnd * 2) + 2)
                                    Call PetMove(X, i, MOVING_WALKING)
                                Else

                                    ' RIGHT works only
                                    Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                End If
                            Else

                                If CanPetMove(X, DIR_LEFT) Then

                                    ' LEFT works only
                                    Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                Else

                                    ' Nothing works, lets try up
                                    If CanPetMove(X, DIR_UP) Then
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        ' Nothing at all works, let it die
                                        Player(X).Pet.MapToGo = Player(X).Pet.Map
                                        Player(X).Pet.XToGo = -1
                                        Player(X).Pet.YToGo = -1
                                    End If
                                End If
                            End If
                        End If
                    Else

                        If y1 <> y2 Then

                            ' UP not down
                            If CanPetMove(X, DIR_UP) Then
                                Call PetMove(X, DIR_UP, MOVING_WALKING)
                            Else

                                ' UP doesn't work
                                If CanPetMove(X, DIR_RIGHT) Then
                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' RIGHT and LEFT work
                                        i = Int((Rnd * 2) + 2)
                                        Call PetMove(X, i, MOVING_WALKING)
                                    Else

                                        ' RIGHT works only
                                        Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                    End If
                                Else

                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' LEFT works only
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    Else

                                        ' Nothing works, lets try down
                                        If CanPetMove(X, DIR_DOWN) Then
                                            Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                        Else

                                            ' Nothing at all works, let it die
                                            Player(X).Pet.MapToGo = Player(X).Pet.Map
                                            Player(X).Pet.XToGo = -1
                                            Player(X).Pet.YToGo = -1
                                        End If
                                    End If
                                End If
                            End If
                        Else

                            ' Question:
                            '   What do we do now?
                            ' Answer:
                            Player(X).Pet.MapToGo = Player(X).Pet.Map
                            Player(X).Pet.XToGo = -1
                            Player(X).Pet.YToGo = -1

                            ' Explaination:
                            '   If y1 - y2 = 0 and x1 - x2 = 0...
                            '   We must be at the location we want to move to!
                            '   Cancel the movement for the future
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub InitServer()
Dim i As Long
Dim f As Long
Dim stringy As String

    CurrentLoad = 0
    Randomize Timer
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = GAME_NAME & " Server" & vbNullChar

    ' Add to the sys tray
    Call Shell_NotifyIcon(NIM_ADD, nid)

    ' Init atmosphe
    GameWeather = WEATHER_NONE
    WeatherSeconds = 0
    GameTime = TIME_DAY
    TimeSeconds = 0
    RainIntensity = 25

    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\maps", vbDirectory)) <> "maps" Then
        Call MkDir$(App.Path & "\Maps")
    End If

    If LCase$(Dir$(App.Path & "\logs", vbDirectory)) <> "logs" Then
        Call MkDir$(App.Path & "\Logs")
    End If
    
    If LCase$(Dir$(App.Path & "\quests", vbDirectory)) <> "quests" Then
        Call MkDir$(App.Path & "\Quests")
    End If

    ' Check if the accounts directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir$(App.Path & "\Accounts")
    End If

    If LCase$(Dir$(App.Path & "\npcs", vbDirectory)) <> "npcs" Then
        Call MkDir$(App.Path & "\Npcs")
    End If

    If LCase$(Dir$(App.Path & "\items", vbDirectory)) <> "items" Then
        Call MkDir$(App.Path & "\Items")
    End If

    If LCase$(Dir$(App.Path & "\spells", vbDirectory)) <> "spells" Then
        Call MkDir$(App.Path & "\Spells")
    End If

    If LCase$(Dir$(App.Path & "\shops", vbDirectory)) <> "shops" Then
        Call MkDir$(App.Path & "\Shops")
    End If
    
    If LCase$(Dir$(App.Path & "\banks", vbDirectory)) <> "banks" Then
        Call MkDir$(App.Path & "\Banks")
    End If

    If LCase$(Dir$(App.Path & "\speech", vbDirectory)) <> "speech" Then
        Call MkDir$(App.Path & "\Speech")
    End If
    SEP_CHAR = Chr$(169)
    END_CHAR = Chr$(174)
    NEXT_CHAR = Chr$(171)
    ServerLog = True

    If Not FileExist("Data.ini") Then
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "GameName", "Chaos Engine"
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "WebSite", ""
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "Port", 4000
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "HPRegen", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "MPRegen", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SPRegen", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "Scrolling", 1

        'SpecialPutVar App.Path & "\Data.ini", "CONFIG", "AutoTurn", 0
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SCRIPTING", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "PAPERDOLL", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SPRITESIZE", 1
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_PLAYERS", 25
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_ITEMS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_NPCS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_SHOPS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_SPELLS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_MAPS", 200
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS", 20
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_GUILDS", 20
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS", 10
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS", 10
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_LEVEL", 500
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_PARTIES", 20
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS", 4
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_SPEECH", 25
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_ELEMENTS", 20
    End If

    If Not FileExist("Stats.ini") Then
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerLevel", 10
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerstr", 10
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerDef", 0
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerMagi", 0
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerSpeed", 0
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerLevel", 10
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerstr", 0
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerDef", 0
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerMagi", 10
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerSpeed", 0
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerLevel", 10
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerstr", 0
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerDef", 0
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerMagi", 0
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerSpeed", 20
    End If
    If Not FileExist("News.ini") Then
        PutVar App.Path & "\News.ini", "DATA", "ServerNews", "News:Change this in the news folder"
    End If
    Call SetStatus("Loading settings...")
    AddHP.Level = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerLevel"))
    AddHP.STR = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerstr"))
    AddHP.DEF = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerDef"))
    AddHP.Magi = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerMagi"))
    AddHP.Speed = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerSpeed"))
    AddMP.Level = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerLevel"))
    AddMP.STR = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerstr"))
    AddMP.DEF = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerDef"))
    AddMP.Magi = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerMagi"))
    AddMP.Speed = Val(GetVar(App.Path & "\Stats.ini", "MP", ""))
    AddSP.Level = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerLevel"))
    AddSP.STR = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerstr"))
    AddSP.DEF = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerDef"))
    AddSP.Magi = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerMagi"))
    AddSP.Speed = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerSpeed"))
    GAME_NAME = Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "GameName"))
    MAX_PLAYERS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_PLAYERS"))
    MAX_ITEMS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_ITEMS"))
    MAX_NPCS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_NPCS"))
    MAX_SHOPS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_SHOPS"))
    MAX_SPELLS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPELLS"))
    MAX_MAPS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAPS"))
    MAX_MAP_ITEMS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS"))
    MAX_GUILDS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILDS"))
    MAX_GUILD_MEMBERS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS"))
    MAX_EMOTICONS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS"))
    MAX_LEVEL = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_LEVEL"))
    MAX_PARTIES = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_PARTIES"))
    MAX_PARTY_MEMBERS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS"))
    MAX_SPEECH = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPEECH"))
    SCRIPTING = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SCRIPTING"))
    MAX_ELEMENTS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_ELEMENTS"))
    PAPERDOLL = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "PAPERDOLL"))
    SPRITESIZE = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SPRITESIZE"))
    MAX_MAPX = 30
    MAX_MAPY = 30

    If GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling") = 0 Then
        MAX_MAPX = 19
        MAX_MAPY = 13
    ElseIf GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling") = 1 Then
        MAX_MAPX = 30
        MAX_MAPY = 30
    End If
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim Item(0 To MAX_ITEMS) As ItemRec
    ReDim Npc(0 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim Grid(1 To MAX_MAPS) As GridRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    ReDim Party(1 To MAX_PARTIES) As PartyRec
    ReDim Speech(1 To MAX_SPEECH) As SpeechRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim Element(0 To MAX_ELEMENTS) As ElementRec
    For i = 1 To MAX_GUILDS
        ReDim Guild(i).Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
    Next
    For i = 1 To MAX_PARTIES
        ReDim Party(i).Member(1 To MAX_PARTY_MEMBERS) As Long
    Next
    For i = 1 To MAX_MAPS
        ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim TempTile(i).DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
        ReDim Grid(i).Loc(0 To MAX_MAPX, 0 To MAX_MAPY) As MapGridRec
    Next
    ReDim Experience(1 To MAX_LEVEL) As Long
    START_MAP = 1
    START_X = MAX_MAPX / 2
    START_Y = MAX_MAPY / 2
    GAME_PORT = GetVar(App.Path & "\Data.ini", "CONFIG", "Port")

    'SCRIPTING
    If SCRIPTING = 1 Then
        Call SetStatus("Loading scripts...")
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    End If

    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT

     ' Init all the player sockets
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next i
    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next

    If Not FileExist("CMessages.ini") Then
        For i = 1 To 6
            PutVar App.Path & "\CMessages.ini", "MESSAGES", "Title" & i, "Custom Msg"
            PutVar App.Path & "\CMessages.ini", "MESSAGES", "Message" & i, ""
        Next
    End If
    For i = 1 To 6
        CMessages(i).Title = GetVar(App.Path & "\CMessages.ini", "MESSAGES", "Title" & i)
        CMessages(i).Message = GetVar(App.Path & "\CMessages.ini", "MESSAGES", "Message" & i)
        frmServer.CustomMsg(i - 1).Caption = CMessages(i).Title
    Next
    
    frmGuide.lstTopics.Clear
    i = 1
    Do While FileExist("Guide\" & i & ".txt")
        f = FreeFile
        Open App.Path & "\Guide\" & i & ".txt" For Input As #f
            Line Input #f, stringy
            frmGuide.lstTopics.AddItem (stringy)
        Close #f
        i = i + 1
    Loop
    frmGuide.lstTopics.Selected(0) = True
    Call SetStatus("Loading emoticons...")
    Call LoadEmos
    Call SetStatus("Loading arrows...")
    Call LoadArrows
    Call SetStatus("Loading exp...")
    Call LoadExps
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
    Call SetStatus("Loading quests...")
    Call LoadQuests
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading speeches...")
    Call LoadSpeeches
    Call SetStatus("Loading elements...")
    Call LoadElements
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Setting up the grid...")
    Call SetUpGrid
    frmServer.MapList.Clear
    For i = 1 To MAX_MAPS
        frmServer.MapList.AddItem i & ": " & Map(i).Name
    Next
    frmServer.MapList.Selected(0) = True

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f
        Close #f
    End If
    
    'Load wordfilter
    Call LoadWordfilter

    ' Start listening
    frmServer.Socket(0).Listen
    
    Call UpdateCaption
    frmLoad.Visible = False
    frmServer.Show
    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
End Sub

Sub PlayerSaveTimer()
    Static MinPassed As Long
    MinPassed = MinPassed + 1

    If MinPassed >= 60 Then
        If TotalOnlinePlayers > 0 Then

            'Call TextAdd(frmServer.txtText(0), "Saving all online players...", True)
            'Call GlobalMsg("Saving all online players...", Pink)
            'For i = 1 To MAX_PLAYERS
            ' If IsPlaying(i) Then
            ' Call SavePlayer(i)
            ' End If
            ' DoEvents
            'Next
            PlayerI = 1
            frmServer.PlayerTimer.Enabled = True
            frmServer.tmrPlayerSave.Enabled = False
        End If
        MinPassed = 0
    End If
End Sub

Sub ServerLogic()
    Dim i As Long

    ' Check for disconnections
    For i = 1 To MAX_PLAYERS

        If frmServer.Socket(i).State > 7 Then
            Call CloseSocket(i)
        End If

    Next

    Call CheckGiveHP
    Call GameAI
End Sub

Sub SetStatus(ByVal Caption As String)
Dim s As String
  
    s = vbNewLine & Caption
    frmLoad.txtStatus.SelText = s
End Sub
