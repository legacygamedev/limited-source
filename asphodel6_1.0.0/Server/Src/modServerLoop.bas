Attribute VB_Name = "modServerLoop"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

Sub ServerLoop()
Dim i As Long
Dim Tick As Currency

Dim X As Long
Dim Y As Long

Dim tmr500 As Currency
Dim tmr1000 As Currency
Dim tmr2500 As Currency
Dim tmr60000 As Currency

Dim LastUpdateSavePlayers As Currency
Dim LastUpdateMapSpawnItems As Currency
Dim LastUpdatePlayerVitals As Currency

    ServerOnline = True
    
    Do While ServerOnline
        Tick = GetTickCountNew
        
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If Map(GetPlayerMap(i)).Tile(GetPlayerX(i), GetPlayerY(i)).Type = Tile_Type.Damage_ Then
                    If TempPlayer(i).DOT_Tile < GetTickCountNew Then
                        DirectDamagePlayer i, GetPlayerMaxVital(i, HP) * (Map(GetPlayerMap(i)).Tile(GetPlayerX(i), GetPlayerY(i)).Data1 * 0.01)
                        TempPlayer(i).DOT_Tile = GetTickCountNew + Map(GetPlayerMap(i)).Tile(GetPlayerX(i), GetPlayerY(i)).Data2
                    End If
                End If
            End If
        Next
        
        ' Check for disconnections every half second
        If tmr500 < Tick Then
        
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            
            UpdateNpcAI
            
            tmr500 = GetTickCountNew + 500
        End If
        
        If tmr1000 < Tick Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            ' Check if a player is unmuted every second
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If Player(i).Char(TempPlayer(i).CharNum).Muted Then
                        If Player(i).Char(TempPlayer(i).CharNum).MuteTime < Tick Then
                            Player(i).Char(TempPlayer(i).CharNum).Muted = False
                            Player(i).Char(TempPlayer(i).CharNum).MuteTime = 0
                            PlayerMsg i, "Your mute time has ran out! You are now unmuted.", Color.BrightGreen
                            TextAdd frmServer.txtText, GetPlayerName(i) & "'s time has ran out! He's been unmuted."
                            UpdatePlayerTable i
                        End If
                    End If
                End If
            Next
            
            tmr1000 = GetTickCountNew + 1000
        End If
        
        If tmr2500 < Tick Then
            For i = 1 To MAX_MAPS
                ' This is used for closing doors
                If TempTile(i).DoorTimer < Tick Then
                    For X = 0 To MAX_MAPX
                        For Y = 0 To MAX_MAPY
                            If Map(i).Tile(X, Y).Type = Tile_Type.Key_ Then
                                If TempTile(i).DoorOpen(X, Y) = YES Then
                                    TempTile(i).DoorOpen(X, Y) = NO
                                    Call SendDataToMap(i, SMapKey & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & END_CHAR)
                                End If
                            End If
                        Next
                    Next
                End If
            Next
            tmr2500 = GetTickCountNew + 2500
        End If
        
        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If LastUpdatePlayerVitals < Tick Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCountNew + 5000
        End If
        
        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If LastUpdateMapSpawnItems < Tick Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCountNew + 300000
        End If
        
        ' Checks to save players every 10 minutes - Can be tweaked
        If LastUpdateSavePlayers < Tick Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCountNew + 600000
        End If
        
        DoEvents
        Sleep 1
        
    Loop
    
    DestroyServer
    
End Sub

Private Sub UpdateMapSpawnItems()
Dim X As Long
Dim Y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For Y = 1 To MAX_MAPS
        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(Y) Then
            ' Clear out unnecessary junk
            For X = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(X, Y)
            Next
            
            ' Spawn the items
            Call SpawnMapItems(Y)
            Call SendMapItemsToAll(Y)
        End If
    Next
    
End Sub

Private Sub UpdateNpcAI()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim n As Long
Dim x1 As Long
Dim y1 As Long
Dim TickCount As Currency
Dim Damage As Long
Dim NpcNum As Long
Dim Target As Long
Dim DidWalk As Boolean

    For Y = 1 To MAX_MAPS
        If PlayersOnMap(Y) Then
            TickCount = GetTickCountNew
            
            For X = 1 To UBound(MapSpawn(Y).Npc)
                NpcNum = MapNpc(Y).MapNpc(X).Num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapSpawn(Y).Npc(X).Num > 0 Then
                    If NpcNum > 0 Then
                        ' If the npc is a attack on sight, search for a player on the map
                        If Npc(NpcNum).Behavior = NPC_Behavior.AttackOnSight Or Npc(NpcNum).Behavior = NPC_Behavior.Guard Then
                            For i = 1 To MAX_PLAYERS
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = Y Then
                                        If MapNpc(Y).MapNpc(X).Target = 0 Then
                                            n = Npc(NpcNum).Range
                                            
                                            If GetPlayerAccess(i) > 0 Then
                                                If frmServer.chkAdminSafety = vbChecked Then
                                                    GoTo SkipNpcTargetting
                                                End If
                                            End If
                                            
                                            ' Are they in range?  if so GET'M!
                                            If IsInRange(MapNpc(Y).MapNpc(X).X, MapNpc(Y).MapNpc(X).Y, GetPlayerX(i), GetPlayerY(i), Npc(NpcNum).Range) Then
                                                If Npc(NpcNum).Behavior = NPC_Behavior.AttackOnSight Or GetPlayerPK(i) = YES Then
                                                    If LenB(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                                                        Call PlayerMsg(i, "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                                                    End If
                                                    
                                                    MapNpc(Y).MapNpc(X).Target = i
                                                End If
                                            End If
SkipNpcTargetting:
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapSpawn(Y).Npc(X).Num > 0 And NpcNum > 0 Then
                    Target = MapNpc(Y).MapNpc(X).Target
                    
                    ' Check to see if its time for the npc to walk
                    If Npc(NpcNum).Behavior <> NPC_Behavior.ShopKeeper Then
                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                            ' Check if the player is even playing, if so follow'm
                            If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                DidWalk = False
                                
                                If GetPlayerAccess(MapNpc(Y).MapNpc(X).Target) > 0 Then
                                    If frmServer.chkAdminSafety = vbChecked Then
                                        MapNpc(Y).MapNpc(X).Target = 0
                                        GoTo AnotherSkipper
                                    End If
                                End If
                                
                                i = Random(0, 3)
                                
                                ' Lets move the npc
                                Select Case i
                                    Case 0
                                        ' Up
                                        If MapNpc(Y).MapNpc(X).Y > GetPlayerY(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Up_) Then
                                                    Call NpcMove(Y, X, E_Direction.Up_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(Y).MapNpc(X).Y < GetPlayerY(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Down_) Then
                                                    Call NpcMove(Y, X, E_Direction.Down_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(Y).MapNpc(X).X > GetPlayerX(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Left_) Then
                                                    Call NpcMove(Y, X, E_Direction.Left_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(Y).MapNpc(X).X < GetPlayerX(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Right_) Then
                                                    Call NpcMove(Y, X, E_Direction.Right_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                    
                                    Case 1
                                        ' Right
                                        If MapNpc(Y).MapNpc(X).X < GetPlayerX(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Right_) Then
                                                    Call NpcMove(Y, X, E_Direction.Right_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(Y).MapNpc(X).X > GetPlayerX(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Left_) Then
                                                    Call NpcMove(Y, X, E_Direction.Left_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(Y).MapNpc(X).Y < GetPlayerY(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Down_) Then
                                                    Call NpcMove(Y, X, E_Direction.Down_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(Y).MapNpc(X).Y > GetPlayerY(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Up_) Then
                                                    Call NpcMove(Y, X, E_Direction.Up_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        
                                    Case 2
                                        ' Down
                                        If MapNpc(Y).MapNpc(X).Y < GetPlayerY(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Down_) Then
                                                    Call NpcMove(Y, X, E_Direction.Down_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(Y).MapNpc(X).Y > GetPlayerY(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Up_) Then
                                                    Call NpcMove(Y, X, E_Direction.Up_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(Y).MapNpc(X).X < GetPlayerX(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Right_) Then
                                                    Call NpcMove(Y, X, E_Direction.Right_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(Y).MapNpc(X).X > GetPlayerX(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Left_) Then
                                                    Call NpcMove(Y, X, E_Direction.Left_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                    
                                    Case 3
                                        ' Left
                                        If MapNpc(Y).MapNpc(X).X > GetPlayerX(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Left_) Then
                                                    Call NpcMove(Y, X, E_Direction.Left_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(Y).MapNpc(X).X < GetPlayerX(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Right_) Then
                                                    Call NpcMove(Y, X, E_Direction.Right_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(Y).MapNpc(X).Y > GetPlayerY(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Up_) Then
                                                    Call NpcMove(Y, X, E_Direction.Up_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(Y).MapNpc(X).Y < GetPlayerY(Target) Then
                                            If Not DidWalk Then
                                                If CanNpcMove(Y, X, E_Direction.Down_) Then
                                                    Call NpcMove(Y, X, E_Direction.Down_, MovementType.Walking)
                                                    DidWalk = True
                                                End If
                                            End If
                                        End If
                                End Select
                                
                                
                            
                                ' Check if we can't move and if player is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(Y).MapNpc(X).X - 1 = GetPlayerX(Target) Then
                                        If MapNpc(Y).MapNpc(X).Y = GetPlayerY(Target) Then
                                            If MapNpc(Y).MapNpc(X).Dir <> E_Direction.Left_ Then
                                                Call NpcDir(Y, X, E_Direction.Left_)
                                            End If
                                            DidWalk = True
                                        End If
                                    End If
                                    If MapNpc(Y).MapNpc(X).X + 1 = GetPlayerX(Target) Then
                                        If MapNpc(Y).MapNpc(X).Y = GetPlayerY(Target) Then
                                            If MapNpc(Y).MapNpc(X).Dir <> E_Direction.Right_ Then
                                                Call NpcDir(Y, X, E_Direction.Right_)
                                            End If
                                            DidWalk = True
                                        End If
                                    End If
                                    If MapNpc(Y).MapNpc(X).X = GetPlayerX(Target) Then
                                        If MapNpc(Y).MapNpc(X).Y - 1 = GetPlayerY(Target) Then
                                            If MapNpc(Y).MapNpc(X).Dir <> E_Direction.Up_ Then
                                                Call NpcDir(Y, X, E_Direction.Up_)
                                            End If
                                            DidWalk = True
                                        End If
                                    End If
                                    If MapNpc(Y).MapNpc(X).X = GetPlayerX(Target) Then
                                        If MapNpc(Y).MapNpc(X).Y + 1 = GetPlayerY(Target) Then
                                            If MapNpc(Y).MapNpc(X).Dir <> E_Direction.Down_ Then
                                                Call NpcDir(Y, X, E_Direction.Down_)
                                            End If
                                            DidWalk = True
                                        End If
                                    End If
                                    
                                    ' We could not move so player must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Random(1, 2)
                                        If i = 1 Then
                                            i = Random(0, 4)
                                            If CanNpcMove(Y, X, i) Then
                                                Call NpcMove(Y, X, i, MovementType.Walking)
                                            End If
                                        End If
                                    End If
AnotherSkipper:
                                End If
                            Else
                                MapNpc(Y).MapNpc(X).Target = 0
                            End If
                        Else
                            i = Random(1, 2)
                            If i = 1 Then
                                i = Random(0, 4)
                                If CanNpcMove(Y, X, i) Then
                                    Call NpcMove(Y, X, i, MovementType.Walking)
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapSpawn(Y).Npc(X).Num > 0 Then
                    If NpcNum > 0 Then
                        Target = MapNpc(Y).MapNpc(X).Target
                        
                        ' Check if the npc can attack the targeted player player
                        If Target > 0 Then
                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                ' Can the npc attack the player?
                                If CanNpcAttackPlayer(X, Target) Then
                                    If LenB(Trim$(Npc(NpcNum).Sound(NpcSound.Attack_))) > 0 Then
                                        SendSound Y, Trim$(Npc(NpcNum).Sound(NpcSound.Attack_))
                                    End If
                                    If Not CanPlayerBlockHit(Target) Then
                                        Damage = Npc(NpcNum).Stat(Stats.Strength) - GetPlayerProtection(Target)
                                        If Damage > 0 Then
                                            Call NpcAttackPlayer(X, Target, Damage)
                                        Else
                                            Call PlayerMsg(Target, "The " & Trim$(Npc(NpcNum).Name) & "'s hit didn't even phase you!", Color.BrightBlue)
                                        End If
                                    Else
                                        Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerEquipmentSlot(Target, Shield))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", Color.BrightCyan)
                                    End If
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(Y).MapNpc(X).Target = 0
                            End If
                        End If
                    End If
                End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If NpcNum > 0 Then
                    If GiveNPCHPTimer < TickCount Then
                        If MapNpc(Y).MapNpc(X).Vital(Vitals.HP) > 0 Then
                            MapNpc(Y).MapNpc(X).Vital(Vitals.HP) = MapNpc(Y).MapNpc(X).Vital(Vitals.HP) + GetNpcVitalRegen(NpcNum, Vitals.HP)
                            
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(Y).MapNpc(X).Vital(Vitals.HP) > GetNpcMaxVital(NpcNum, Vitals.HP) Then
                                MapNpc(Y).MapNpc(X).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
                            End If
                            
                            SendNPCVital Y, X
                        End If
                        GiveNPCHPTimer = GetTickCountNew + 10000
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If NpcNum = 0 Then
                    If MapSpawn(Y).Npc(X).Num > 0 Then
                        If MapNpc(Y).MapNpc(X).SpawnWait < TickCount Then
                            Call SpawnNpc(X, Y)
                        End If
                    End If
                End If
            Next
        End If
    Next
    
End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        UpdatePlayerVital i
    Next
    
End Sub

Public Sub UpdatePlayerVital(ByVal Index As Long)

    If IsPlaying(Index) Then
        'If GetPlayerVital(Index, Vitals.HP) <> GetPlayerMaxVital(Index, Vitals.HP) Then
            Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + GetPlayerVitalRegen(Index, Vitals.HP))
            Call SendVital(Index, Vitals.HP)
        'End If
        'If GetPlayerVital(Index, Vitals.MP) <> GetPlayerMaxVital(Index, Vitals.MP) Then
            Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + GetPlayerVitalRegen(Index, Vitals.MP))
            Call SendVital(Index, Vitals.MP)
        'End If
        'If GetPlayerVital(Index, Vitals.SP) <> GetPlayerMaxVital(Index, Vitals.SP) Then
            Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) + GetPlayerVitalRegen(Index, Vitals.SP))
            Call SendVital(Index, Vitals.SP)
        'End If
    End If
    
End Sub

Private Sub UpdateSavePlayers()
Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd(frmServer.txtText, "Saving all online players...")
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call SavePlayer(i)
            End If
        Next
        Call GlobalMsg("All online players saved.", Color.Pink)
    End If
End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server is shutting down in " & Secs & " seconds!", Color.BrightBlue)
        Call TextAdd(frmServer.txtText, "Automated server shut down in " & Secs & " seconds.")
    End If
    
    Secs = Secs - 1
    
    If Secs <= 0 Then
        Call GlobalMsg("Server shut down! Good bye.", Color.BrightRed)
        ServerOnline = False
    End If

End Sub
