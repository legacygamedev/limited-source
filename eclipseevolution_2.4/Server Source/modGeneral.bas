Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

' Debugging only - tracks what packets use how much bandwidth
Public Const DebugPacketsOut As Boolean = True
Public PacketsOut() As Long

' Version constants
Public Const CLIENT_MAJOR = 3
Public Const CLIENT_MINOR = 1
Public Const CLIENT_REVISION = 4

' Used for respawning items
Public SpawnSeconds As Long

' Used for weather effects
Public GameWeather As Long
Public WeatherSeconds As Long
Public GameTime As Long
Public GameCycle As Long
Public TimeSeconds As Long
Public RainIntensity As Long
Public GameClock As String
Public Gamespeed As Long
Public Hours As Integer
Public Seconds As Long
Public Minutes As Integer

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

' Used for classes
Public ClassesOn As Byte

'Used for loading progress bar - change this for the number of times the bar is incremented!
Public Const PROG_BAR_MAX = 19

Public Wierd As Long
Public TimeDisable As Boolean


Sub CheckGiveHP()

  Dim i As Long

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
            
        Next i

        GiveHPTimer = GetTickCount
    End If

End Sub

Sub CheckSpawnMapItems()

  Dim X As Long
  Dim Y As Long

    ' Used for map item respawning
    SpawnSeconds = SpawnSeconds + 1

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////

    If SpawnSeconds >= 120 Then
        ' 2 minutes have passed

        For Y = 1 To MAX_MAPS
            ' Make sure no one is on the map when it respawns

            If PlayersOnMap(Y) = False Then
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

Sub DestroyServer()

  Dim i As Long
  Dim i2
  Dim Percent As Integer
  
    Call Shell_NotifyIcon(NIM_DELETE, nid)
    
    If DebugPacketsOut Then
        For i = 0 To 255
            WriteINI "OUT", STR$(i), STR$(PacketsOut(i)), App.Path & "\Logs\PacketsOut.ini"
            i2 = i2 + PacketsOut(i)
        Next i
        WriteINI "OUT", "TOTAL", STR$(i2), App.Path & "\Logs\PacketsOut.ini"
        Erase PacketsOut
    End If

    Call SetStatus("Shutting down...")
    frmLoad.Visible = True
    frmServer.Visible = False
    DoEvents

    Call SetStatus("Unloading sockets and timers...")

    For i = 1 To MAX_PLAYERS
        Percent = i / MAX_PLAYERS * 100
        UpdatePercentage Percent, "Unloading sockets and timers..."

        Unload frmServer.Socket(i)
    Next i

    End

End Sub

Sub GameAI()

    On Error Resume Next
  Dim i As Integer
  Dim X As Integer
  Dim Y As Integer
  Dim n As Integer
  Dim x1 As Integer
  Dim y1 As Integer
  Dim TickCount As Long
  Dim Damage As Long
  Dim DistanceX As Integer
  Dim DistanceY As Integer
  Dim NpcNum As Integer
  Dim Target As Integer
  Dim DidWalk As Boolean
  Dim npcn As Integer

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
    End If

    For Y = 1 To MAX_MAPS

        If PlayersOnMap(Y) = YES Then
            TickCount = GetTickCount

            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////

            If TickCount > TempTile(Y).DoorTimer + 5000 Then

                For y1 = 0 To MAX_MAPY
                    For x1 = 0 To MAX_MAPX

                        If Map(Y).tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(Y).DoorOpen(x1, y1) = YES Then
                            TempTile(Y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(Y, PacketID.MapKey & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If

                        If Map(Y).tile(x1, y1).Type = TILE_TYPE_DOOR And TempTile(Y).DoorOpen(x1, y1) = YES Then
                            TempTile(Y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(Y, PacketID.MapKey & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If

                    Next x1
                Next y1
            End If

            For X = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(Y, X).num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////

                If 0 + MapNpc(Y, X).owner <> 0 Then
                    If 0 + GetPlayerMap(MapNpc(Y, X).owner) <> Y Then
                        npcn = 1

                        Do While npcn <= MAX_MAP_NPCS

                            If MapNpc(Y, X).owner = 0 Then
                                npcn = MAX_MAP_NPCS + 1
                             Else

                                If 0 + Map(GetPlayerMap(MapNpc(Y, X).owner)).Npc(npcn) = 0 Then
                                    MapNpc(Y, X).owner = 0
                                    Call ScriptSpawnNpc(npcn, Y, MapNpc(Y, X).X, MapNpc(Y, X).Y, 0)
                                 Else
                                    npcn = npcn + 1
                                End If

                            End If
                        Loop

                    End If
                 Else
                    ' Make sure theres a npc with the map

                    If Map(Y).Npc(X) > 0 And MapNpc(Y, X).num > 0 Then
                        ' If the npc is a attack on sight, search for a player on the map

                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then

                            For i = 1 To MAX_PLAYERS

                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = Y And MapNpc(Y, X).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITER Then
                                        n = Npc(NpcNum).Range

                                        DistanceX = MapNpc(Y, X).X - GetPlayerX(i)
                                        DistanceY = MapNpc(Y, X).Y - GetPlayerY(i)

                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1

                                        ' Are they in range?  if so GET'M!

                                        If DistanceX <= n And DistanceY <= n Then
                                            If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                If Trim(Npc(NpcNum).AttackSay) <> "" Then
                                                    Call PlayerMsg(i, "A " & Trim(Npc(NpcNum).Name) & " : " & Trim(Npc(NpcNum).AttackSay) & "", SayColor)
                                                End If

                                                MapNpc(Y, X).Target = i
                                            End If

                                        End If
                                    End If
                                End If
                            Next i

                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map

                If Map(Y).Npc(X) > 0 And MapNpc(Y, X).num > 0 Then

                    If 0 + MapNpc(Y, X).owner <> 0 Then
                        Target = 0 + Player(MapNpc(Y, X).owner).Target
                     Else
                        Target = MapNpc(Y, X).Target
                    End If

                    ' Check to see if its time for the npc to walk

                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        ' Check to see if we are following a player or not

                        If Target > 0 Then
                            ' Check if the player is even playing, if so follow'm

                            If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                DidWalk = False

                                i = Int(Rnd * 5)

                                ' Lets move the npc

                                Select Case i
                                 Case 0
                                    ' Up

                                    If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_UP) Then
                                            Call NPCMove(Y, X, DIR_UP, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Down

                                    If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                            Call NPCMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Left

                                    If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                            Call NPCMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Right

                                    If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                            Call NPCMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If

                                 Case 1
                                    ' Right

                                    If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                            Call NPCMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Left

                                    If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                            Call NPCMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Down

                                    If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                            Call NPCMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Up

                                    If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_UP) Then
                                            Call NPCMove(Y, X, DIR_UP, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If

                                 Case 2
                                    ' Down

                                    If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                            Call NPCMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Up

                                    If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_UP) Then
                                            Call NPCMove(Y, X, DIR_UP, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Right

                                    If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                            Call NPCMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Left

                                    If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                            Call NPCMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If

                                 Case 3
                                    ' Left

                                    If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                            Call NPCMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Right

                                    If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                            Call NPCMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Up

                                    If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_UP) Then
                                            Call NPCMove(Y, X, DIR_UP, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                    ' Down

                                    If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                            Call NPCMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                            DidWalk = True
                                        End If

                                    End If
                                End Select

                                ' Check if we can't move and if player is behind something and if we can just switch dirs

                                If Not DidWalk Then
                                    If MapNpc(Y, X).X - 1 = GetPlayerX(Target) And MapNpc(Y, X).Y = GetPlayerY(Target) Then
                                        If MapNpc(Y, X).Dir <> DIR_LEFT Then
                                            Call NPCDir(Y, X, DIR_LEFT)
                                        End If

                                        DidWalk = True
                                    End If

                                    If MapNpc(Y, X).X + 1 = GetPlayerX(Target) And MapNpc(Y, X).Y = GetPlayerY(Target) Then
                                        If MapNpc(Y, X).Dir <> DIR_RIGHT Then
                                            Call NPCDir(Y, X, DIR_RIGHT)
                                        End If

                                        DidWalk = True
                                    End If

                                    If MapNpc(Y, X).X = GetPlayerX(Target) And MapNpc(Y, X).Y - 1 = GetPlayerY(Target) Then
                                        If MapNpc(Y, X).Dir <> DIR_UP Then
                                            Call NPCDir(Y, X, DIR_UP)
                                        End If

                                        DidWalk = True
                                    End If

                                    If MapNpc(Y, X).X = GetPlayerX(Target) And MapNpc(Y, X).Y + 1 = GetPlayerY(Target) Then
                                        If MapNpc(Y, X).Dir <> DIR_DOWN Then
                                            Call NPCDir(Y, X, DIR_DOWN)
                                        End If

                                        DidWalk = True
                                    End If

                                    ' We could not move so player must be behind something, walk randomly.

                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)

                                        If i = 1 Then
                                            i = Int(Rnd * 4)

                                            If CanNpcMove(Y, X, i) Then
                                                Call NPCMove(Y, X, i, MOVING_WALKING)
                                            End If

                                        End If
                                    End If
                                End If
                             Else
                                MapNpc(Y, X).Target = 0
                            End If

                         Else

                            If 0 + MapNpc(Y, X).owner <> 0 Then
                                If GetPlayerTargetNpc(MapNpc(Y, X).owner) <> 0 Then
                                    If MapNpc(Y, GetPlayerTargetNpc(MapNpc(Y, X).owner)).X < MapNpc(Y, X).X Then
                                        If CanNpcMove(Y, X, 2) Then Call NPCMove(Y, X, 2, 1)
                                        Exit Sub
                                    End If

                                    If MapNpc(Y, GetPlayerTargetNpc(MapNpc(Y, X).owner)).X > MapNpc(Y, X).X Then
                                        If CanNpcMove(Y, X, 3) Then Call NPCMove(Y, X, 3, 1)
                                        Exit Sub
                                    End If

                                    If MapNpc(Y, GetPlayerTargetNpc(MapNpc(Y, X).owner)).Y < MapNpc(Y, X).Y - 1 Then
                                        If CanNpcMove(Y, X, 0) Then Call NPCMove(Y, X, 0, 1)
                                        Exit Sub
                                    End If

                                    If MapNpc(Y, GetPlayerTargetNpc(MapNpc(Y, X).owner)).Y > MapNpc(Y, X).Y + 1 Then
                                        If CanNpcMove(Y, X, 1) Then Call NPCMove(Y, X, 1, 1)
                                        Exit Sub
                                    End If

                                 Else

                                    If Player(MapNpc(Y, X).owner).Char(Player(MapNpc(Y, X).owner).CharNum).X < MapNpc(Y, X).X Then
                                        If CanNpcMove(Y, X, 2) Then Call NPCMove(Y, X, 2, 1)
                                        Exit Sub
                                    End If

                                    If Player(MapNpc(Y, X).owner).Char(Player(MapNpc(Y, X).owner).CharNum).X > MapNpc(Y, X).X Then
                                        If CanNpcMove(Y, X, 3) Then Call NPCMove(Y, X, 3, 1)
                                        Exit Sub
                                    End If

                                    If Player(MapNpc(Y, X).owner).Char(Player(MapNpc(Y, X).owner).CharNum).Y < MapNpc(Y, X).Y - 1 Then
                                        If CanNpcMove(Y, X, 0) Then Call NPCMove(Y, X, 0, 1)
                                        Exit Sub
                                    End If

                                    If Player(MapNpc(Y, X).owner).Char(Player(MapNpc(Y, X).owner).CharNum).Y > MapNpc(Y, X).Y + 1 Then
                                        If CanNpcMove(Y, X, 1) Then Call NPCMove(Y, X, 1, 1)
                                        Exit Sub
                                    End If

                                End If
                             Else
                                i = Int(Rnd * 4)

                                If i = 1 Then
                                    i = Int(Rnd * 4)

                                    If CanNpcMove(Y, X, i) Then
                                        Call NPCMove(Y, X, i, MOVING_WALKING)
                                    End If

                                End If
                            End If
                        End If
                    End If

                    ' /////////////////////////////////////////////
                    ' // This is used for npcs to attack players //
                    ' /////////////////////////////////////////////
                    ' Make sure theres a npc with the map

                    If Map(Y).Npc(X) > 0 And MapNpc(Y, X).num > 0 Then
                        If 0 + MapNpc(Y, X).owner <> 0 Then
                            Target = 0 + Player(MapNpc(Y, X).owner).Target
                         Else
                            Target = MapNpc(Y, X).Target
                        End If

                        ' Check if the npc can attack the targeted player player

                        If Target > 0 Then
                            If 0 + MapNpc(Y, X).owner <> 0 Then
                                If GetPlayerMap(MapNpc(Y, X).owner) = Y Then
                                    If MapNpc(GetPlayerTargetNpc(MapNpc(Y, X).owner)).X = 1 Then
                                        If CanAttributeNpcAttackNpc(Y, X, MapNpc(Y, X).X, MapNpc(Y, X).Y) Then
                                            'pet attacking npc
                                            Damage = Int(Npc(Y).STR * 2) - Int(Npc(GetPlayerTargetNpc(MapNpc(Y, X).owner)).DEF / 2)

                                            If Damage > 0 Then
                                                MapNpc(GetPlayerTargetNpc(MapNpc(Y, X).owner)).HP = MapNpc(GetPlayerTargetNpc(MapNpc(Y, X).owner)).HP - Damage
                                            End If

                                            'npc attacking pet
                                            Damage = Int(Npc(GetPlayerTargetNpc(MapNpc(Y, X).owner)).STR * 2) - Int(Npc(Y).DEF / 2)

                                            If Damage > 0 Then
                                                MapNpc(Y).HP = MapNpc(Y).HP - Damage
                                            End If

                                        End If
                                    End If
                                End If
                             Else
                                ' Is the target playing and on the same map?

                                If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                    ' Can the npc attack the player?

                                    If CanNpcAttackPlayer(X, Target) Then
                                        If Not CanPlayerBlockHit(Target) Then
                                            Damage = Npc(NpcNum).STR - GetPlayerProtection(Target)

                                            If Damage > 0 Then
                                                Call NpcAttackPlayer(X, Target, Damage)
                                             Else
                                                Call BattleMsg(Target, "The " & Trim(Npc(NpcNum).Name) & " could not hurt you.", BrightBlue, 1)

                                                'Call PlayerMsg(Target, "The " & Trim(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                            End If

                                         Else
                                            Call BattleMsg(Target, "you blocked " & Trim(Npc(NpcNum).Name) & "'s hit.", BrightCyan, 1)

                                            'Call PlayerMsg(Target, "Your " & Trim(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                        End If

                                    End If
                                 Else
                                    ' Player left map or game, set target to 0
                                    MapNpc(Y, X).Target = 0
                                End If

                            End If
                        End If
                    End If

                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp

                If MapNpc(Y, X).num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(Y, X).HP > 0 Then
                        MapNpc(Y, X).HP = MapNpc(Y, X).HP + GetNpcHPRegen(NpcNum)

                        ' Check if they have more then they should and if so just set it to max

                        If MapNpc(Y, X).HP > GetNpcMaxhp(NpcNum) Then
                            MapNpc(Y, X).HP = GetNpcMaxhp(NpcNum)
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

                If MapNpc(Y, X).num = 0 And Map(Y).Npc(X) > 0 Then
                    If TickCount > MapNpc(Y, X).SpawnWait + (Npc(Map(Y).Npc(X)).SpawnSecs * 1000) Then
                        Call SpawnNPC(X, Y)
                    End If

                End If

                If MapNpc(Y, X).num > 0 Then
                    Call SendDataToMap(Y, PacketID.NPCHP & SEP_CHAR & X & SEP_CHAR & MapNpc(Y, X).HP & SEP_CHAR & GetNpcMaxhp(MapNpc(Y, X).num) & SEP_CHAR & END_CHAR)
                End If

            Next X

        End If
        DoEvents
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

Function GetPlayerTargetNpc(ByVal index As Long)

    If index > 0 Then
        GetPlayerTargetNpc = Player(index).targetnpc
    End If

End Function

Sub IncrementBar()

    On Error Resume Next
    'Increment prog bar
    frmLoad.loadProgressBar.Value = frmLoad.loadProgressBar.Value + 1
    On Error GoTo 0
    
End Sub

Sub InitServer()

  Dim i As Integer
  Dim n As Integer
  Dim f As Long

    Randomize Timer

    'Must be done before any packets are sent/received
    InitPacketIDs

    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = GAME_NAME & " Server" & vbNullChar
    ' Add to the sys tray
    'Call Shell_NotifyIcon(NIM_ADD, nid)

    ' Init atmosphere
    GameWeather = WEATHER_NONE
    WeatherSeconds = 0
    GameCycle = YES
    TimeSeconds = 0
    RainIntensity = 25

    ' Init progress bar
    frmLoad.loadProgressBar.Max = PROG_BAR_MAX

    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\maps", vbDirectory)) <> "maps" Then Call MkDir(App.Path & "\maps")
    If LCase$(Dir$(App.Path & "\logs", vbDirectory)) <> "logs" Then Call MkDir(App.Path & "\Logs")

    ' Check if the accounts directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\accounts", vbDirectory)) <> "accounts" Then Call MkDir(App.Path & "\accounts")
    If LCase$(Dir$(App.Path & "\npcs", vbDirectory)) <> "npcs" Then Call MkDir(App.Path & "\Npcs")
    If LCase$(Dir$(App.Path & "\items", vbDirectory)) <> "items" Then Call MkDir(App.Path & "\Items")
    If LCase$(Dir$(App.Path & "\spells", vbDirectory)) <> "spells" Then Call MkDir(App.Path & "\Spells")
    If LCase$(Dir$(App.Path & "\shops", vbDirectory)) <> "shops" Then Call MkDir(App.Path & "\Shops")
    If LCase$(Dir$(App.Path & "\banks", vbDirectory)) <> "banks" Then Call MkDir(App.Path & "\Banks")
    If LCase$(Dir$(App.Path & "\skills", vbDirectory)) <> "skills" Then Call MkDir(App.Path & "\Skills")
    If LCase$(Dir$(App.Path & "\quests", vbDirectory)) <> "quests" Then Call MkDir(App.Path & "\Quests")

    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)

    ServerLog = True

    If Not FileExist("Data.ini") Then
        PutVar App.Path & "\Data.ini", "CONFIG", "GameName", ""
        PutVar App.Path & "\Data.ini", "CONFIG", "WebSite", ""
        PutVar App.Path & "\Data.ini", "CONFIG", "Port", 4000
        PutVar App.Path & "\Data.ini", "CONFIG", "HPRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "MPRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "SPRegen", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "Scrolling", 1
        'PutVar App.Path & "\Data.ini", "CONFIG", "AutoTurn", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 1
        PutVar App.Path & "\Data.ini", "CONFIG", "PaperDoll", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "32x64", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "Custom", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "NonToScroll", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "TeToEe", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "Level", 0
        PutVar App.Path & "\Data.ini", "CONFIG", "ENCRYPT_PASS", ""
        PutVar App.Path & "\Data.ini", "CONFIG", "ENCRYPT_TYPE", "BMP"
        PutVar App.Path & "\Data.ini", "CONFIG", "mousebug", "0"
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
        PutVar App.Path & "\Data.ini", "CONFIG", "Classes", 1
    End If

    If Not FileExist("Stats.ini") Then
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

    If Not FileExist("News.ini") Then
        PutVar App.Path & "\News.ini", "DATA", "ServerNews", "News:Change this in the news folder"
    End If

    If Not FileExist("Tiles.ini") Then

        For i = 0 To 100
            PutVar App.Path & "\Tiles.ini", "Names", "Tile" & i, CStr(i)
        Next i

    End If

    Call SetStatus("Loading settings...")

    On Error GoTo LoadErr
    AddHP.Level = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerLevel"))
    AddHP.STR = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerStr"))
    AddHP.DEF = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerDef"))
    AddHP.Magi = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerMagi"))
    AddHP.Speed = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerSpeed"))
    AddMP.Level = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerLevel"))
    AddMP.STR = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerStr"))
    AddMP.DEF = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerDef"))
    AddMP.Magi = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerMagi"))
    AddMP.Speed = Val(GetVar(App.Path & "\Stats.ini", "MP", ""))
    AddSP.Level = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerLevel"))
    AddSP.STR = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerStr"))
    AddSP.DEF = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerDef"))
    AddSP.Magi = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerMagi"))
    AddSP.Speed = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerSpeed"))

    GAME_NAME = Trim(GetVar(App.Path & "\Data.ini", "CONFIG", "GameName"))
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
    Scripting = GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting")
    MAX_PARTY_MEMBERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS")
    MAX_ELEMENTS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_ELEMENTS")
    MAX_SCRIPTSPELLS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SCRIPTSPELLS")
    Paperdoll = GetVar(App.Path & "\Data.ini", "CONFIG", "PaperDoll")
    Spritesize = GetVar(App.Path & "\Data.ini", "CONFIG", "SpriteSize")
    STAT1 = Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "Stat1"))
    STAT2 = Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "Stat2"))
    STAT3 = Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "Stat3"))
    STAT4 = Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "Stat4"))
    CUSTOM_SPRITE = GetVar(App.Path & "\Data.ini", "CONFIG", "Custom")
    ENCRYPT_PASS = GetVar(App.Path & "\Data.ini", "CONFIG", "ENCRYPT_PASS")
    ENCRYPT_TYPE = GetVar(App.Path & "\Data.ini", "CONFIG", "ENCRYPT_TYPE")
    MAX_SKILLS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SKILLS")
    MAX_QUESTS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_QUESTS")
    ClassesOn = GetVar(App.Path & "\Data.ini", "CONFIG", "Classes")

    MAX_MAPX = 30
    MAX_MAPY = 30

    If GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling") = 0 Then
        IS_SCROLLING = 0
        MAX_MAPX = 19
        MAX_MAPY = 14
    ElseIf GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling") = 1 Then
        IS_SCROLLING = 1
        MAX_MAPX = 30
        MAX_MAPY = 30
    End If

    GoTo LoadSuccess

LoadErr:
    Call MsgBox("Error reading from Data.ini or Stats.ini. Check to make sure they are set up correctly! Error: " & Err.number & " (" & Err.Description & ")", vbCritical)
    End

LoadSuccess:

    'Restore error handling
    On Error GoTo 0

    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim MapPackets(1 To MAX_MAPS) As String
    ReDim Temp_Map(1 To MAX_MAPS) As Temp_MapRec
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim Item(0 To MAX_ITEMS) As ItemRec
    ReDim skill(1 To MAX_SKILLS) As SkillRec
    ReDim Quest(0 To MAX_QUESTS) As QuestRec
    ReDim Npc(0 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim Element(0 To MAX_ELEMENTS) As ElementRec
    If DebugPacketsOut Then ReDim PacketsOut(0 To 255)

    For i = 1 To MAX_GUILDS
        ReDim Guild(i).Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
    Next i

    For i = 1 To MAX_MAPS
        ReDim Map(i).tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim Temp_Map(i).tile(0 To MAX_MAPX, 0 To MAX_MAPY) As Temp_TileRec
        ReDim TempTile(i).DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
    Next i

    For i = 1 To MAX_PLAYERS
        For n = 1 To MAX_CHARS
            ReDim Player(i).Char(n).SkillLvl(1 To MAX_SKILLS)
            ReDim Player(i).Char(n).SkillExp(1 To MAX_SKILLS)
        Next n
    Next i

    ReDim Experience(1 To MAX_LEVEL) As Long

    START_MAP = 1
    START_X = MAX_MAPX / 2
    START_Y = MAX_MAPY / 2

    GAME_PORT = GetVar(App.Path & "\Data.ini", "CONFIG", "Port")

    Set CTimers = New Collection

    Call IncrementBar

    On Error GoTo ScriptErr

    'Scripting
    frmServer.lblScriptOn.caption = "Scripts are: Off"
    ' Check for Main.txt

    If Not FileExist("\Scripts\Main.txt") Then
        Call MsgBox("Main.txt not found. Scripts disabled.", vbExclamation)
        Scripting = 0
    End If

    ' Continue happily

    If Scripting = 1 Then
        Call SetStatus("Loading scripts...")
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        frmServer.lblScriptOn.caption = "Scripts are: On"
    End If

    Call IncrementBar

    GoTo ScriptsGood

ScriptErr:
    If MsgBox("Unknown error occured loading scripts, disabled. Err: " & Err.number & ", Desc: " & Err.Description, vbOKCancel) = vbCancel Then Call DestroyServer

ScriptsGood:

    On Error GoTo 0

    ' Check the open port
    Call CheckOpenPort(GAME_PORT)

    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT

    ' Init all the player sockets
    Call SetStatus("Initializing player array...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next i

    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next i

    Call IncrementBar

    If Not FileExist("CMessages.ini") Then

        For i = 1 To 6
            PutVar App.Path & "\CMessages.ini", "MESSAGES", "Title" & i, "Custom Msg"
            PutVar App.Path & "\CMessages.ini", "MESSAGES", "Message" & i, ""
        Next i

    End If

    For i = 1 To 6
        CMessages(i).Title = GetVar(App.Path & "\CMessages.ini", "MESSAGES", "Title" & i)
        CMessages(i).Message = GetVar(App.Path & "\CMessages.ini", "MESSAGES", "Message" & i)
        frmServer.CustomMsg(i - 1).caption = CMessages(i).Title
    Next i

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
    Call SetStatus("Clearing skills...")
    Call ClearSkills
    Call SetStatus("Clearing quests...")
    Call ClearQuests
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing exp...")
    Call ClearExps
    Call SetStatus("Clearing emoticons...")
    Call ClearEmos
    Call IncrementBar
    Call SetStatus("Loading emoticons...")
    Call IncrementBar
    Call LoadEmos
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
    Call LoadExps
    Call SetStatus("Loading classes...")
    Call IncrementBar
    Call LoadClasses
    'Call SetStatus("Loading first class advandcement...")
    'Call LoadClasses2
    'Call SetStatus("Loading second class advandcement...")
    'Call Loadclasses3
    Call SetStatus("Loading maps...")
    Call IncrementBar
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call IncrementBar
    Call LoadItems
    Call SetStatus("Loading skills...")
    Call IncrementBar
    Call LoadSkills
    Call SetStatus("Loading quests...")
    Call IncrementBar
    Call LoadQuests
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

    For i = 1 To MAX_MAPS
        frmServer.MapList.AddItem i & ": " & Map(i).Name
    Next i

    frmServer.MapList.Selected(0) = True

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it

    If Not FileExist("accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f
        Close #f
    End If

    'ASGARD
    'Load wordfilter
    Call LoadWordfilter

    'Error handling for 'Address in use' error
    Err.Clear
    On Error Resume Next

    ' Start listening
    frmServer.Socket(0).Listen

    'RTE 10048 occured

    If Err.number = 10048 Then
        Call MsgBox("A server on this port is already running! Please change the port or close the other server.", vbCritical)
    End If

    'Restore error handling
    On Error GoTo 0

    Call UpdateCaption

    frmLoad.Visible = False
    frmServer.Show

    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
    frmServer.tmrScriptedTimer.Enabled = True

End Sub

Sub PlayerSaveTimer()

  Static MinPassed As Long
  Dim i As Long
  Dim settingsLocation As Long

    settingsLocation = Val(GetVar("Data.ini", "Config", "SaveTime"))
    MinPassed = MinPassed + 1

    If MinPassed >= Val(settingsLocation) And Val(settingsLocation) <> 0 Then

        For i = 1 To MAX_PLAYERS

            If IsPlaying(i) Then
                Call SavePlayer(i)
            End If

        Next i

        PlayerI = 1
        frmServer.PlayerTimer.Enabled = True
        frmServer.tmrPlayerSave.Enabled = False
    End If

    MinPassed = 0

End Sub

Sub ScriptedTimer()

  Dim X As Long
  Dim n As Long
  Dim CustomTimer As clsCTimers

    n = 0
    X = CTimers.Count

    For Each CustomTimer In CTimers
        n = n + 1

        If GetTickCount > CustomTimer.tmrWait Then
            MyScript.ExecuteStatement "Scripts\Main.txt", CustomTimer.Name ' & " " & Index & "," & PointType

            If CTimers.Count < X Then
                n = n - X - CTimers.Count
                X = CTimers.Count
            End If

            If n > 0 Then CTimers.Item(n).tmrWait = GetTickCount + CustomTimer.Interval Else Exit For
        End If

    Next CustomTimer

End Sub

Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, ByVal spawn_x As Long, ByVal spawn_y As Long, ByVal NpcNum As Long)

    '                         NPC_index               map_number          X spawn          y spawn            NPC_number
  Dim packet As String
  Dim i As Long

    ' Check for subscript out of range

    If MapNpcNum < 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    If NpcNum = 0 Then
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        MapNpc(MapNum, MapNpcNum).num = 0
        Map(MapNum).Npc(MapNpcNum) = 0
        MapNpc(MapNum, MapNpcNum).Target = 0
        MapNpc(MapNum, MapNpcNum).HP = 0
        MapNpc(MapNum, MapNpcNum).MP = 0
        MapNpc(MapNum, MapNpcNum).SP = 0
        MapNpc(MapNum, MapNpcNum).Dir = 0
        MapNpc(MapNum, MapNpcNum).X = 0
        MapNpc(MapNum, MapNpcNum).Y = 0

        'Packet = PacketID.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(mapnum, MapNpcNum).num & SEP_CHAR & MapNpc(mapnum, MapNpcNum).x & SEP_CHAR & MapNpc(mapnum, MapNpcNum).y & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(mapnum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
        'Call SendDataToMap(mapnum, Packet)
        Call SaveMap(MapNum)
    End If

    'MapNpc(mapnum, MapNpcNum).num = 0
    'MapNpc(mapnum, MapNpcNum).SpawnWait = GetTickCount
    'MapNpc(mapnum, MapNpcNum).HP = 0
    'Call SendDataToMap(mapnum, PacketID.NPCDead & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)

    Map(MapNum).Revision = Map(MapNum).Revision + 1

    MapNpc(MapNum, MapNpcNum).num = NpcNum
    Map(MapNum).Npc(MapNpcNum) = NpcNum

    MapNpc(MapNum, MapNpcNum).Target = 0

    MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxhp(NpcNum)
    MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
    MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)

    MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

    MapNpc(MapNum, MapNpcNum).X = spawn_x
    MapNpc(MapNum, MapNpcNum).Y = spawn_y

    packet = PacketID.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, packet)

    Call SaveMap(MapNum)

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call SendDataTo(i, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & SEP_CHAR & END_CHAR)
        End If

    Next i

End Sub

Sub ServerLogic()

  Dim i As Long

    ' Check for disconnections

    For i = 1 To MAX_PLAYERS
        On Error Resume Next

        If frmServer.Socket(i).State > 7 Then
            Call CloseSocket(i)
        End If

    Next i

    Call CheckGiveHP
    Call GameAI
    Call ScriptedTimer

End Sub

Sub SetStatus(ByVal Status As String)

    frmLoad.lblStatus.caption = Status

End Sub

