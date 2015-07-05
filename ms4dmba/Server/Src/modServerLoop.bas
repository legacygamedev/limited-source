Attribute VB_Name = "modServerLoop"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
Dim i As Long
Dim Tick As Long
Dim tmr500 As Long
Dim tmr1000 As Long

Dim LastUpdateSavePlayers As Long
Dim LastUpdateMapSpawnItems As Long
Dim LastUpdatePlayerVitals As Long

    ServerOnline = True
    
    Do While ServerOnline
        Tick = GetTickCount
        
        '/////////////////////////////////////////////
        '// Checks if it's time to update something //
        '/////////////////////////////////////////////
        
        ' Check for disconnections every half second
        If Tick > tmr500 Then
        
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            
            UpdateNpcAI
           
            tmr500 = GetTickCount + 500
        End If
        
        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            tmr1000 = GetTickCount + 1000
        End If
                
        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If
        
        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If
        
        ' Checks to save players every 10 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 600000
        End If
        
        Sleep 1
        DoEvents
        
    Loop
End Sub

Private Sub UpdateMapSpawnItems()
Dim x As Long
Dim y As Long
    
    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For y = 1 To MAX_MAPS
        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(y) Then
            ' Clear out unnecessary junk
            For x = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(x, y)
            Next
                
            ' Spawn the items
            Call SpawnMapItems(y)
            Call SendMapItemsToAll(y)
        End If
        DoEvents
    Next
End Sub

Private Sub UpdateNpcAI()
Dim i As Long
Dim x As Long
Dim y As Long
Dim n As Long
Dim x1 As Long
Dim y1 As Long
Dim TickCount As Long
Dim Damage As Long
Dim DistanceX As Long
Dim DistanceY As Long
Dim NpcNum As Long
Dim Target As Long
Dim DidWalk As Boolean
Dim Buffer As clsBuffer
            
    For y = 1 To MAX_MAPS
        If PlayersOnMap(y) = YES Then
            TickCount = GetTickCount
            
            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            If TickCount > TempTile(y).DoorTimer + 5000 Then
                For x1 = 0 To Map(y).MaxX
                    For y1 = 0 To Map(y).MaxY
                        If Map(y).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(y).DoorOpen(x1, y1) = YES Then
                            TempTile(y).DoorOpen(x1, y1) = NO
                            
                            Set Buffer = New clsBuffer
                            
                            Buffer.WriteLong SMapKey
                            Buffer.WriteLong x1
                            Buffer.WriteLong y1
                            Buffer.WriteLong 0
                            
                            SendDataToMap y, Buffer.ToArray()
                            
                            Set Buffer = Nothing
                            
                        End If
                    Next
                Next
            End If
            
            For x = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(y).Npc(x).Num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(x) > 0 And MapNpc(y).Npc(x).Num > 0 Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = y And MapNpc(y).Npc(x).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                    n = Npc(NpcNum).Range
                                    
                                    DistanceX = MapNpc(y).Npc(x).x - GetPlayerX(i)
                                    DistanceY = MapNpc(y).Npc(x).y - GetPlayerY(i)
                                    
                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            If LenB(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                                                Call PlayerMsg(i, "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                                            End If
                                            
                                            MapNpc(y).Npc(x).Target = i
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
                If Map(y).Npc(x) > 0 And MapNpc(y).Npc(x).Num > 0 Then
                    Target = MapNpc(y).Npc(x).Target
                    
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
                                        If MapNpc(y).Npc(x).y > GetPlayerY(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y).Npc(x).y < GetPlayerY(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y).Npc(x).x > GetPlayerX(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y).Npc(x).x < GetPlayerX(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 1
                                        ' Right
                                        If MapNpc(y).Npc(x).x < GetPlayerX(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y).Npc(x).x > GetPlayerX(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y).Npc(x).y < GetPlayerY(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y).Npc(x).y > GetPlayerY(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        
                                    Case 2
                                        ' Down
                                        If MapNpc(y).Npc(x).y < GetPlayerY(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y).Npc(x).y > GetPlayerY(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y).Npc(x).x < GetPlayerX(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y).Npc(x).x > GetPlayerX(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 3
                                        ' Left
                                        If MapNpc(y).Npc(x).x > GetPlayerX(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y).Npc(x).x < GetPlayerX(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y).Npc(x).y > GetPlayerY(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y).Npc(x).y < GetPlayerY(Target) And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                End Select
                                
                                
                            
                                ' Check if we can't move and if player is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(y).Npc(x).x - 1 = GetPlayerX(Target) And MapNpc(y).Npc(x).y = GetPlayerY(Target) Then
                                        If MapNpc(y).Npc(x).Dir <> DIR_LEFT Then
                                            Call NpcDir(y, x, DIR_LEFT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y).Npc(x).x + 1 = GetPlayerX(Target) And MapNpc(y).Npc(x).y = GetPlayerY(Target) Then
                                        If MapNpc(y).Npc(x).Dir <> DIR_RIGHT Then
                                            Call NpcDir(y, x, DIR_RIGHT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y).Npc(x).x = GetPlayerX(Target) And MapNpc(y).Npc(x).y - 1 = GetPlayerY(Target) Then
                                        If MapNpc(y).Npc(x).Dir <> DIR_UP Then
                                            Call NpcDir(y, x, DIR_UP)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y).Npc(x).x = GetPlayerX(Target) And MapNpc(y).Npc(x).y + 1 = GetPlayerY(Target) Then
                                        If MapNpc(y).Npc(x).Dir <> DIR_DOWN Then
                                            Call NpcDir(y, x, DIR_DOWN)
                                        End If
                                        DidWalk = True
                                    End If
                                    
                                    ' We could not move so player must be behind something, walk randomly.
                                    If Not DidWalk Then
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
                                MapNpc(y).Npc(x).Target = 0
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
                If Map(y).Npc(x) > 0 And MapNpc(y).Npc(x).Num > 0 Then
                    Target = MapNpc(y).Npc(x).Target
                    
                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                        ' Is the target playing and on the same map?
                        If IsPlaying(Target) And GetPlayerMap(Target) = y Then
                            ' Can the npc attack the player?
                            If CanNpcAttackPlayer(x, Target) Then
                                If Not CanPlayerBlockHit(Target) Then
                                    Damage = Npc(NpcNum).Stat(Stats.Strength) - GetPlayerProtection(Target)
                                    Call NpcAttackPlayer(x, Target, Damage)
                                Else
                                    Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerEquipmentSlot(Target, Shield))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                End If
                            End If
                        Else
                            ' Player left map or game, set target to 0
                            MapNpc(y).Npc(x).Target = 0
                        End If
                    End If
                End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(y).Npc(x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(y).Npc(x).Vital(Vitals.HP) > 0 Then
                        MapNpc(y).Npc(x).Vital(Vitals.HP) = MapNpc(y).Npc(x).Vital(Vitals.HP) + GetNpcVitalRegen(NpcNum, Vitals.HP)
                    
                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(y).Npc(x).Vital(Vitals.HP) > GetNpcMaxVital(NpcNum, Vitals.HP) Then
                            MapNpc(y).Npc(x).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
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
                If MapNpc(y).Npc(x).Num = 0 And Map(y).Npc(x) > 0 Then
                    If TickCount > MapNpc(y).Npc(x).SpawnWait + (Npc(Map(y).Npc(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, y)
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
End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                Call SendVital(i, Vitals.HP)
            End If
            If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                Call SendVital(i, Vitals.MP)
            End If
            If GetPlayerVital(i, Vitals.SP) <> GetPlayerMaxVital(i, Vitals.SP) Then
                Call SetPlayerVital(i, Vitals.SP, GetPlayerVital(i, Vitals.SP) + GetPlayerVitalRegen(i, Vitals.SP))
                Call SendVital(i, Vitals.SP)
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")
        Call GlobalMsg("Saving all online players...", Pink)
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call SavePlayer(i)
            End If
            DoEvents
        Next
    End If
End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If
    
    Secs = Secs - 1
    
    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub

