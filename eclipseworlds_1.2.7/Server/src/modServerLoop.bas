Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Integer, x As Integer, n As Integer
    Dim Tick As Long
    Dim TickCPS As Long, CPS As Long, tmr25 As Long, tmr500 As Long, tmr1000 As Long, FrameTime As Long
    Dim LastUpdateSavePlayers As Long, LastUpdateMapSpawnItems As Long, LastUpdateVitals As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = timeGetTime
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr25 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    ' Check if they've completed casting, and if so set the actual spell going
                    If tempplayer(i).SpellBuffer.Spell > 0 Then
                        If timeGetTime > tempplayer(i).SpellBuffer.Timer + (Spell(Account(i).Chars(GetPlayerChar(i)).Spell(tempplayer(i).SpellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, tempplayer(i).SpellBuffer.Spell, tempplayer(i).SpellBuffer.target, tempplayer(i).SpellBuffer.TType
                        End If
                    End If
                    
                    ' Check if need to turn off stunned
                    If tempplayer(i).StunDuration > 0 Then
                        If timeGetTime > tempplayer(i).StunTimer + (tempplayer(i).StunDuration * 1000) Then
                            tempplayer(i).StunDuration = 0
                            tempplayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    
                    ' Check if we need to reset the spell CD
                    For x = 1 To MAX_PLAYER_SPELLS
                        If GetPlayerSpell(i, x) > 0 Then
                            If GetPlayerSpellCD(i, x) > 0 Then
                                If GetPlayerSpellCD(i, x) <= timeGetTime Then
                                    Call SetPlayerSpellCD(i, x, 0)
                                    Call SendSpellCooldown(i, x)
                                End If
                            End If
                        End If
                    Next
                    
                    ' Check regen timer
                    If tempplayer(i).StopRegen Then
                        If tempplayer(i).StopRegenTimer + 5000 < timeGetTime Then
                            tempplayer(i).StopRegen = False
                            tempplayer(i).StopRegenTimer = 0
                        End If
                    End If
                    
                    ' Check PvP timer
                    If tempplayer(i).PVPTimer > 0 And tempplayer(i).PVPTimer < timeGetTime Then
                        tempplayer(i).PVPTimer = 0
                        If IsPlaying(i) Then
                            Call SendPlayerPK(i)
                        ElseIf tempplayer(i).HasLogged Then
                            Call LeftGame(i)
                        End If
                    ElseIf tempplayer(i).HasLogged Then
                        Call LeftGame(i)
                    End If
                    
                    ' HoT and DoT logic
                    For x = 1 To MAX_DOTS
                        HandleDoT_Player i, x
                        HandleHoT_Player i, x
                    Next
                End If
            Next
            
            For i = 1 To MAX_MAPS
                For x = 1 To Map(i).NPC_HighIndex
                    ' Check if they've completed casting, and if so set the actual spell going
                    If MapNPC(i).NPC(x).SpellBuffer.Spell > 0 Then
                        If timeGetTime > MapNPC(i).NPC(x).SpellBuffer.Timer + (Spell(MapNPC(i).NPC(x).SpellBuffer.Spell).CastTime * 1000) Then
                            If MapNPC(i).NPC(x).targetType = TARGET_TYPE_PLAYER Then
                                Call NPCSpellPlayer(x, MapNPC(i).NPC(x).SpellBuffer.target)
                            ElseIf MapNPC(i).NPC(x).targetType = TARGET_TYPE_NPC Then
                                Call NPCSpellNPC(x, MapNPC(i).NPC(x).SpellBuffer.target, i)
                            End If
                            Call ClearNPCSpellBuffer(i, x)
                        End If
                    End If
                    
                    ' Check regen timer
                    If MapNPC(i).NPC(x).StopRegen Then
                        If MapNPC(i).NPC(x).StopRegenTimer + 5000 < timeGetTime Then
                            MapNPC(i).NPC(x).StopRegen = False
                            MapNPC(i).NPC(x).StopRegenTimer = 0
                        End If
                    End If
                Next
            Next

            If GameCPS > 0 Then
                frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            End If
            UpdateEventLogic
            tmr25 = timeGetTime + 25
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To Player_HighIndex

                If frmServer.Socket(i).State <> sckConnected Then
                    If tempplayer(i).PVPTimer < timeGetTime Then
                        Call LeftGame(i)
                    End If
                End If
            Next
            
            UpdateMapLogic
            tmr500 = timeGetTime + 500
        End If

        If Tick > tmr1000 Then
            If IsShuttingDown Then
                Call HandleShutdown
            End If
            
            ' Update the form labels, and reset the packets per second
            frmServer.lblPackIn.Caption = Trim$(STR$(PacketsIn))
            frmServer.lblPackOut.Caption = Trim$(STR$(PacketsOut))
            PacketsIn = 0
            PacketsOut = 0

            ' Update the Server Online Time
            ServerSeconds = ServerSeconds + 1

            If ServerSeconds > 59 Then
                ServerMinutes = ServerMinutes + 1
                ServerSeconds = 0

                If ServerMinutes > 59 Then
                    ServerMinutes = 0
                    ServerHours = ServerHours + 1
                End If
            End If
            
            frmServer.lblTime.Caption = Trim$(STR$(ServerHours)) & ":" & Trim$(STR$(ServerMinutes)) & ":" & Trim$(STR$(ServerSeconds))

            Call ConsumeLogic
        
            tmr1000 = timeGetTime + 1000
        End If

        ' Checks to update player vitals every 5 seconds
        If Tick > LastUpdateVitals Then
            UpdatePlayerVitals
            LastUpdateVitals = timeGetTime + 5000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = timeGetTime + 300000
        End If

        ' Checks to save players every minute
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = timeGetTime + 60000
        End If

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub UpdateMapSpawnItems()
    Dim x As Long
    Dim Y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For Y = 1 To MAX_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(Y) Then

            ' Clear out unnecessary junk
            For x = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(x, Y)
            Next

            ' Spawn the items
            Call SpawnMapItems(Y)
        End If
        DoEvents
    Next
End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, x As Long, MapNum As Integer, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, NPCNum As Long
    Dim target As Long, targetType As Byte, DidWalk As Boolean, buffer As clsBuffer, Resource_Index As Long
    Dim targetX As Long, targetY As Long, Target_Verify As Boolean, ii As Long

    For MapNum = 1 To MAX_MAPS
        ' Items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(MapNum, i).Num > 0 Then
                If Not Trim$(MapItem(MapNum, i).playerName) = vbNullString Then
                    ' Make item public
                    If MapItem(MapNum, i).PlayerTimer < timeGetTime Then
                        ' Make it public
                        MapItem(MapNum, i).playerName = vbNullString
                        MapItem(MapNum, i).PlayerTimer = 0
                        
                        ' Send updates to everyone
                        SendMapItemToMap MapNum, i
                    End If
                    
                    ' Despawn item
                    If MapItem(MapNum, i).CanDespawn Then
                        If MapItem(MapNum, i).DespawnTimer < timeGetTime Then
                            ' Despawn it
                            ClearMapItem i, MapNum
                            
                            ' Send updates to everyone
                            SendMapItemToMap MapNum, i
                        End If
                    End If
                End If
            End If
        Next
        
        ' Check for DoTs + Hots
        For i = 1 To Map(MapNum).NPC_HighIndex
            If MapNPC(MapNum).NPC(i).Num > 0 Then
                For x = 1 To MAX_DOTS
                    HandleDoT_NPC MapNum, i, x
                    HandleHoT_NPC MapNum, i, x
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(MapNum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(MapNum).Resource_Count
                Resource_Index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(i).x, ResourceCache(MapNum).ResourceData(i).Y).Data1

                If Resource_Index > 0 Then
                    If ResourceCache(MapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(MapNum).ResourceData(i).Cur_Reward < 1 Then  ' dead or fucked up
                        If ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_Index).RespawnTime * 1000) < timeGetTime Then
                            ResourceCache(MapNum).ResourceData(i).ResourceTimer = timeGetTime
                            ResourceCache(MapNum).ResourceData(i).ResourceState = 0 ' Normal
                            
                            ' Re-set health to resource root
                            ResourceCache(MapNum).ResourceData(i).Cur_Reward = Random(Resource(Resource_Index).Reward_Min, Resource(Resource_Index).Reward_Max)
                        End If
                    End If
                End If
            Next
            
            SendResourceCacheToMap MapNum
        End If

        If PlayersOnMap(MapNum) = YES Then
            TickCount = timeGetTime
            
            For x = 1 To Map(MapNum).NPC_HighIndex
                NPCNum = MapNPC(MapNum).NPC(x).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).NPC(x) > 0 And MapNPC(MapNum).NPC(x).Num > 0 Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(NPCNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(NPCNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        ' Make sure it's not stunned
                        If Not MapNPC(MapNum).NPC(x).StunDuration > 0 Then
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And MapNPC(MapNum).NPC(x).target = 0 And GetPlayerAccess(i) <= STAFF_MODERATOR Then
                                        n = NPC(NPCNum).Range
                                        DistanceX = MapNPC(MapNum).NPC(x).x - GetPlayerX(i)
                                        DistanceY = MapNPC(MapNum).NPC(x).Y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range
                                        If DistanceX <= n And DistanceY <= n Then
                                            If NPC(NPCNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(NPCNum).Behavior = NPC_BEHAVIOR_GUARD And GetPlayerPK(i) = PLAYER_KILLER Then
                                                MapNPC(MapNum).NPC(x).targetType = TARGET_TYPE_PLAYER
                                                MapNPC(MapNum).NPC(x).target = i
                                                tempplayer(i).PVPTimer = timeGetTime + 15000
                                                Call SendMapNPCTarget(MapNum, x, MapNPC(MapNum).NPC(x).target, MapNPC(MapNum).NPC(x).targetType)
                                                
                                                If Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                                                    Call SendChatBubble(MapNum, x, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), White)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            
                            ' Check if target was found for NPC targetting
                            If MapNPC(MapNum).NPC(x).target = 0 Then
                                ' Make sure it belongs to a faction
                                If NPC(NPCNum).Faction > 0 Then
                                    ' Search for npc of another faction to target
                                    For i = 1 To Map(MapNum).NPC_HighIndex
                                        ' Exist
                                        If MapNPC(MapNum).NPC(i).Num > 0 Then
                                            ' Different npc
                                            If MapNPC(MapNum).NPC(i).Num <> MapNPC(MapNum).NPC(x).Num Then
                                                ' Not friendly or shopkeeper
                                                If Not NPC(MapNPC(MapNum).NPC(i).Num).Behavior = NPC_BEHAVIOR_QUEST Then
                                                    ' Different faction
                                                    If NPC(MapNPC(MapNum).NPC(i).Num).Faction > 0 Then
                                                        If Not NPC(MapNPC(MapNum).NPC(i).Num).Faction = NPC(NPCNum).Faction Then
                                                            n = NPC(NPCNum).Range
                                                            DistanceX = MapNPC(MapNum).NPC(x).x - CLng(MapNPC(MapNum).NPC(i).x)
                                                            DistanceY = MapNPC(MapNum).NPC(x).Y - CLng(MapNPC(MapNum).NPC(i).Y)
                                                            
                                                            ' Make sure we get a positive value
                                                            If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                            If DistanceY < 0 Then DistanceY = DistanceY * -1
                                                            
                                                            ' Are they in range
                                                            If DistanceX <= n And DistanceY <= n Then
                                                                If NPC(NPCNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(NPCNum).Behavior = NPC_BEHAVIOR_GUARD Then
                                                                    MapNPC(MapNum).NPC(x).targetType = TARGET_TYPE_NPC
                                                                    MapNPC(MapNum).NPC(x).target = i
                                                                    Call SendMapNPCTarget(MapNum, x, MapNPC(MapNum).NPC(x).target, MapNPC(MapNum).NPC(x).targetType)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If
                
                Target_Verify = False
                        
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure there's a npc with the map
                If Map(MapNum).NPC(x) > 0 And MapNPC(MapNum).NPC(x).Num > 0 Then
                    If MapNPC(MapNum).NPC(x).StunDuration > 0 Then
                        ' Check if we can unstun them
                        If timeGetTime > MapNPC(MapNum).NPC(x).StunTimer + (MapNPC(MapNum).NPC(x).StunDuration * 1000) Then
                            MapNPC(MapNum).NPC(x).StunDuration = 0
                            MapNPC(MapNum).NPC(x).StunTimer = 0
                        End If
                    Else
                        target = MapNPC(MapNum).NPC(x).target
                        targetType = MapNPC(MapNum).NPC(x).targetType
      
                        ' Check to see if its time for the npc to walk
                        If Not NPC(NPCNum).Behavior = NPC_BEHAVIOR_QUEST Then
                            If targetType = 1 Then ' Player
                                ' Check to see if we are following a player or not
                                If target > 0 Then
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                        DidWalk = False
                                        Target_Verify = True
                                        targetY = GetPlayerY(target)
                                        targetX = GetPlayerX(target)
                                    Else
                                        MapNPC(MapNum).NPC(x).targetType = TARGET_TYPE_NONE
                                        MapNPC(MapNum).NPC(x).target = 0
                                        Call SendMapNPCTarget(MapNum, x, 0, 0)
                                    End If
                                End If
                            
                            ElseIf targetType = 2 Then ' NPC
                                If target > 0 Then
                                    If MapNPC(MapNum).NPC(target).Num > 0 Then
                                        DidWalk = False
                                        Target_Verify = True
                                        targetX = MapNPC(MapNum).NPC(target).x
                                        targetY = MapNPC(MapNum).NPC(target).Y
                                    Else
                                        MapNPC(MapNum).NPC(x).targetType = TARGET_TYPE_NONE
                                        MapNPC(MapNum).NPC(x).target = 0
                                        Call SendMapNPCTarget(MapNum, x, 0, 0)
                                    End If
                                End If
                            End If
                            
                            If Target_Verify Then
                                If IsOneBlockAway(targetX, targetY, CLng(MapNPC(MapNum).NPC(x).x), CLng(MapNPC(MapNum).NPC(x).Y)) = False Then
                                    ' Up Left
                                    If MapNPC(MapNum).NPC(x).Y > targetY And MapNPC(MapNum).NPC(x).x > targetX And Not DidWalk Then
                                        If MapNPC(MapNum).NPC(x).x > targetX Then
                                            If CanNPCMove(MapNum, x, DIR_UPLEFT) Then
                                                Call NPCMove(MapNum, x, DIR_UPLEFT, MOVING_RUNNING)
                                                DidWalk = True
                                            End If
                                        End If
                                    End If
                                    
                                    ' Up right
                                    If MapNPC(MapNum).NPC(x).Y > targetY And MapNPC(MapNum).NPC(x).x < targetX And Not DidWalk Then
                                        If MapNPC(MapNum).NPC(x).x < targetX Then
                                            If CanNPCMove(MapNum, x, DIR_UPRIGHT) Then
                                                Call NPCMove(MapNum, x, DIR_UPRIGHT, MOVING_RUNNING)
                                                DidWalk = True
                                            End If
                                        End If
                                    End If
                                    
                                    ' Down Left
                                    If MapNPC(MapNum).NPC(x).Y < targetY And MapNPC(MapNum).NPC(x).x > targetX And Not DidWalk Then
                                        If MapNPC(MapNum).NPC(x).x > targetX Then
                                            If CanNPCMove(MapNum, x, DIR_DOWNLEFT) Then
                                                Call NPCMove(MapNum, x, DIR_DOWNLEFT, MOVING_RUNNING)
                                                DidWalk = True
                                            End If
                                        End If
                                    End If
                                    
                                    ' Down Right
                                    If MapNPC(MapNum).NPC(x).Y < targetY And MapNPC(MapNum).NPC(x).x < targetX And Not DidWalk Then
                                        If MapNPC(MapNum).NPC(x).x < targetX Then
                                            If CanNPCMove(MapNum, x, DIR_DOWNRIGHT) Then
                                                Call NPCMove(MapNum, x, DIR_DOWNRIGHT, MOVING_RUNNING)
                                                DidWalk = True
                                            End If
                                        End If
                                    End If
                                    
                                    ' Left
                                    If MapNPC(MapNum).NPC(x).x > targetX And Not DidWalk Then
                                        If CanNPCMove(MapNum, x, DIR_LEFT) Then
                                            Call NPCMove(MapNum, x, DIR_LEFT, MOVING_RUNNING)
                                            DidWalk = True
                                        End If
                                    End If
                                    
                                    ' Right
                                    If MapNPC(MapNum).NPC(x).x < targetX And Not DidWalk Then
                                        If CanNPCMove(MapNum, x, DIR_RIGHT) Then
                                            Call NPCMove(MapNum, x, DIR_RIGHT, MOVING_RUNNING)
                                            DidWalk = True
                                        End If
                                    End If
                                    
                                    ' Up
                                    If MapNPC(MapNum).NPC(x).Y > targetY And Not DidWalk Then
                                        If CanNPCMove(MapNum, x, DIR_UP) Then
                                            Call NPCMove(MapNum, x, DIR_UP, MOVING_RUNNING)
                                            DidWalk = True
                                        End If
                                    End If
                                    
                                    ' Down
                                    If MapNPC(MapNum).NPC(x).Y < targetY And Not DidWalk Then
                                        If CanNPCMove(MapNum, x, DIR_DOWN) Then
                                            Call NPCMove(MapNum, x, DIR_DOWN, MOVING_RUNNING)
                                            DidWalk = True
                                        End If
                                    End If
        
                                    ' We could not move so Target must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        ii = 0

                                        Do While DidWalk = False And ii <= 10 ' Try 10 times to move
                                            i = Int(Rnd * 8)
                                            
                                            If CanNPCMove(MapNum, x, i) Then
                                                Call NPCMove(MapNum, x, i, MOVING_RUNNING)
                                                DidWalk = True
                                            End If

                                            ii = ii + 1
                                        Loop

                                        ' We could not move so target must be behind something, change direction to face them
                                        If Not DidWalk Then
                                            ' Up Left
                                            If MapNPC(MapNum).NPC(x).Y > targetY And MapNPC(MapNum).NPC(x).x > targetX And Not DidWalk Then
                                                If MapNPC(MapNum).NPC(x).x > targetX Then
                                                    Call NPCDir(MapNum, x, DIR_UPLEFT)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' Up right
                                            If MapNPC(MapNum).NPC(x).Y > targetY And MapNPC(MapNum).NPC(x).x < targetX And Not DidWalk Then
                                                If MapNPC(MapNum).NPC(x).x < targetX Then
                                                    Call NPCDir(MapNum, x, DIR_UPRIGHT)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' Down Left
                                            If MapNPC(MapNum).NPC(x).Y < targetY And MapNPC(MapNum).NPC(x).x > targetX And Not DidWalk Then
                                                If MapNPC(MapNum).NPC(x).x > targetX Then
                                                    Call NPCDir(MapNum, x, DIR_DOWNLEFT)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' Down Right
                                            If MapNPC(MapNum).NPC(x).Y < targetY And MapNPC(MapNum).NPC(x).x < targetX And Not DidWalk Then
                                                If MapNPC(MapNum).NPC(x).x < targetX Then
                                                    Call NPCDir(MapNum, x, DIR_DOWNRIGHT)
                                                    DidWalk = True
                                                End If
                                            End If
                                            
                                            ' Left
                                            If MapNPC(MapNum).NPC(x).x > targetX And Not DidWalk Then
                                                Call NPCDir(MapNum, x, DIR_LEFT)
                                                DidWalk = True
                                            End If
                                            
                                            ' Right
                                            If MapNPC(MapNum).NPC(x).x < targetX And Not DidWalk Then
                                                Call NPCDir(MapNum, x, DIR_RIGHT)
                                                DidWalk = True
                                            End If
                                            
                                            ' Up
                                            If MapNPC(MapNum).NPC(x).Y > targetY And Not DidWalk Then
                                                Call NPCDir(MapNum, x, DIR_UP)
                                                DidWalk = True
                                            End If
                                            
                                            ' Down
                                            If MapNPC(MapNum).NPC(x).Y < targetY And Not DidWalk Then
                                                Call NPCDir(MapNum, x, DIR_DOWN)
                                                DidWalk = True
                                            End If
                                        End If
                                    End If
                                Else
                                    Call NPCDir(MapNum, x, GetNpcDir(targetX, targetY, CLng(MapNPC(MapNum).NPC(x).x), CLng(MapNPC(MapNum).NPC(x).Y)))
                                End If
                            Else
                                i = Int(Rnd * 8)
                                If i = 1 Then
                                    i = Int(Rnd * 8)
    
                                    If CanNPCMove(MapNum, x, i) Then
                                        Call NPCMove(MapNum, x, i, MOVING_RUNNING)
                                        DidWalk = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for NPCs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).NPC(x) > 0 And MapNPC(MapNum).NPC(x).Num > 0 Then
                    target = MapNPC(MapNum).NPC(x).target
                    targetType = MapNPC(MapNum).NPC(x).targetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                        If targetType = TARGET_TYPE_PLAYER Then ' Player
                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                TryNPCAttackPlayer x, target
                            Else
                                ' Player left map or game, set target to 0
                                MapNPC(MapNum).NPC(x).target = 0
                                MapNPC(MapNum).NPC(x).targetType = TARGET_TYPE_NONE
                                Call SendMapNPCTarget(MapNum, x, 0, 0)
                            End If
                        ElseIf targetType = TARGET_TYPE_NPC Then
                            If MapNPC(MapNum).NPC(target).Num > 0 Then ' NPC exists
                                Call TryNPCAttackNPC(MapNum, x, target)
                            Else
                                ' NPC is dead or non-existant
                                MapNPC(MapNum).NPC(x).target = 0
                                MapNPC(MapNum).NPC(x).targetType = TARGET_TYPE_NONE
                                Call SendMapNPCTarget(MapNum, x, 0, 0)
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNPC(MapNum).NPC(x).StopRegen Then
                    If MapNPC(MapNum).NPC(x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNPC(MapNum).NPC(x).Vital(Vitals.HP) < GetNPCMaxVital(NPCNum, Vitals.HP) Then
                            MapNPC(MapNum).NPC(x).Vital(Vitals.HP) = MapNPC(MapNum).NPC(x).Vital(Vitals.HP) + GetNPCVitalRegen(NPCNum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNPC(MapNum).NPC(x).Vital(Vitals.HP) > GetNPCMaxVital(NPCNum, Vitals.HP) Then
                                MapNPC(MapNum).NPC(x).Vital(Vitals.HP) = GetNPCMaxVital(NPCNum, Vitals.HP)
                            End If
                        End If
                        
                        If MapNPC(MapNum).NPC(x).Vital(Vitals.MP) < GetNPCMaxVital(NPCNum, Vitals.MP) Then
                            MapNPC(MapNum).NPC(x).Vital(Vitals.MP) = MapNPC(MapNum).NPC(x).Vital(Vitals.MP) + GetNPCVitalRegen(NPCNum, Vitals.MP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNPC(MapNum).NPC(x).Vital(Vitals.MP) > GetNPCMaxVital(NPCNum, Vitals.MP) Then
                                MapNPC(MapNum).NPC(x).Vital(Vitals.MP) = GetNPCMaxVital(NPCNum, Vitals.MP)
                            End If
                        End If
                        Call SendMapNPCVitals(MapNum, x)
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNPC(MapNum).NPC(x).Num = 0 And Map(MapNum).NPC(x) > 0 Then
                    If TickCount > MapNPC(MapNum).NPC(x).SpawnWait + (NPC(Map(MapNum).NPC(x)).SpawnSecs * 1000) Then
                        Call SpawnNPC(x, MapNum)
                    End If
                End If
            Next
        End If
        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If timeGetTime > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = timeGetTime
    End If
End Sub

Sub UpdatePlayerVitals()
    Dim i As Long
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not tempplayer(i).StopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, HP)
                    
                    SendActionMsg GetPlayerMap(i), "+" & GetPlayerVitalRegen(i, Vitals.HP), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(i) * 32, GetPlayerY(i) * 32, 1
                    
                    ' Send vitals to party if in one
                    If tempplayer(i).InParty > 0 Then SendPartyVitals tempplayer(i).InParty, i
                End If
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, MP)
                    
                    SendActionMsg GetPlayerMap(i), "+" & GetPlayerVitalRegen(i, Vitals.MP), BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(i) * 32, GetPlayerY(i) * 32, 1
                    
                    ' Send vitals to party if in one
                    If tempplayer(i).InParty > 0 Then SendPartyVitals tempplayer(i).InParty, i
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                Call SaveAccount(i)
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

Private Sub ConsumeLogic()
    Dim i As Long
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If tempplayer(i).VitalCycle(HP) > 0 Then
                If tempplayer(i).VitalPotion(HP) > 0 Then
                    ' Don't heal if we're already full health
                    If Account(i).Chars(GetPlayerChar(i)).Vital(HP) < GetPlayerMaxVital(i, HP) Then
                        Account(i).Chars(GetPlayerChar(i)).Vital(HP) = Account(i).Chars(GetPlayerChar(i)).Vital(HP) + Round(Item(tempplayer(i).VitalPotion(HP)).AddHP / Item(tempplayer(i).VitalPotion(HP)).Data1)
                        
                        ' Prevent overhealing
                        If Account(i).Chars(GetPlayerChar(i)).Vital(HP) > GetPlayerMaxVital(i, HP) Then
                            Account(i).Chars(GetPlayerChar(i)).Vital(HP) = GetPlayerMaxVital(i, HP)
                        End If
                        
                        Call SendActionMsg(GetPlayerMap(i), "+" & Round(Item(tempplayer(i).VitalPotion(HP)).AddHP / Item(tempplayer(i).VitalPotion(HP)).Data1), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(i) * 32, GetPlayerY(i) * 32)
                        
                        ' Send the vital
                        Call SendVital(i, HP)
                    End If
                End If
                
                ' Lower the cycle by 1
                tempplayer(i).VitalCycle(HP) = tempplayer(i).VitalCycle(HP) - 1
                
                ' Clear out old data if the cycle is over
                If tempplayer(i).VitalCycle(HP) = 0 Then
                    tempplayer(i).VitalPotion(HP) = 0
                End If
            End If
            
            If tempplayer(i).VitalCycle(MP) > 0 Then
                If tempplayer(i).VitalPotion(MP) > 0 Then
                    ' Don't heal if we're already full mana
                    If Account(i).Chars(GetPlayerChar(i)).Vital(MP) < GetPlayerMaxVital(i, MP) Then
                        Account(i).Chars(GetPlayerChar(i)).Vital(MP) = Account(i).Chars(GetPlayerChar(i)).Vital(MP) + Round(Item(tempplayer(i).VitalPotion(MP)).AddMP / Item(tempplayer(i).VitalPotion(MP)).Data1)
                        
                        ' Prevent overhealing
                        If Account(i).Chars(GetPlayerChar(i)).Vital(MP) > GetPlayerMaxVital(i, MP) Then
                            Account(i).Chars(GetPlayerChar(i)).Vital(MP) = GetPlayerMaxVital(i, MP)
                        End If
                        
                        Call SendActionMsg(GetPlayerMap(i), "+" & Round(Item(tempplayer(i).VitalPotion(MP)).AddMP / Item(tempplayer(i).VitalPotion(MP)).Data1), BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(i) * 32, GetPlayerY(i) * 32)
                        
                        ' Send the vital
                        Call SendVital(i, MP)
                    End If
                End If
                
                ' Lower the cycle by 1
                tempplayer(i).VitalCycle(MP) = tempplayer(i).VitalCycle(MP) - 1
                
                ' Clear out old data if the cycle is over
                If tempplayer(i).VitalCycle(MP) = 0 Then
                    tempplayer(i).VitalPotion(MP) = 0
                End If
            End If
        End If
    Next
End Sub

Function CanEventMoveTowardsPlayer(PlayerID As Long, MapNum As Long, eventID As Long) As Long
    Dim i As Long, x As Long, Y As Long, x1 As Long, y1 As Long, DidWalk As Boolean, WalkThrough As Long
    Dim tim As Long, sX As Long, sY As Long, pos() As Long, reachable As Boolean, j As Long, LastSum As Long, Sum As Long, FX As Long, FY As Long
    Dim path() As Vector, LastX As Long, LastY As Long, did As Boolean
    
    ' This does not work for global events so this MUST be a player one....
    ' This Event returns a direction, 4 is not a valid direction so we assume fail unless otherwise told.
    CanEventMoveTowardsPlayer = 4
    
    If PlayerID <= 0 Or PlayerID > Player_HighIndex Then Exit Function
    If MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > tempplayer(PlayerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(PlayerID)
    Y = GetPlayerY(PlayerID)
    x1 = tempplayer(PlayerID).EventMap.EventPages(eventID).x
    y1 = tempplayer(PlayerID).EventMap.EventPages(eventID).Y
    WalkThrough = Map(MapNum).Events(tempplayer(PlayerID).EventMap.EventPages(eventID).eventID).Pages(tempplayer(PlayerID).EventMap.EventPages(eventID).PageID).WalkThrough
    ' Add option for pathfinding to random guessing option.
    
    If PathfindingType = 1 Then
        i = Int(Rnd * 5)
        DidWalk = False
        
        ' Lets move the event
        Select Case i
            Case 0
                ' Up
                If y1 > Y And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        DidWalk = True
                    End If
                End If
        
                ' Down
                If y1 < Y And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        DidWalk = True
                    End If
                End If
        
                ' Left
                If x1 > x And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        DidWalk = True
                    End If
                End If
        
                ' Right
                If x1 < x And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        DidWalk = True
                    End If
                End If
        
            Case 1
                ' Right
                If x1 < x And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        DidWalk = True
                    End If
                End If
                
                ' Left
                If x1 > x And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        DidWalk = True
                    End If
                End If
                
                ' Down
                If y1 < Y And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        DidWalk = True
                    End If
                End If
                
                ' Up
                If y1 > Y And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        DidWalk = True
                    End If
                End If
        
            Case 2
                ' Down
                If y1 < Y And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        DidWalk = True
                    End If
                End If
                
                ' Up
                If y1 > Y And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        DidWalk = True
                    End If
                End If
                
                ' Right
                If x1 < x And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        DidWalk = True
                    End If
                End If
                
                ' Left
                If x1 > x And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        DidWalk = True
                    End If
                End If
        
            Case 3
                ' Left
                If x1 > x And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        DidWalk = True
                    End If
                End If
                
                ' Right
                If x1 < x And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        DidWalk = True
                    End If
                End If
                
                ' Up
                If y1 > Y And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        DidWalk = True
                    End If
                End If
                
                ' Down
                If y1 < Y And Not DidWalk Then
                    If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        DidWalk = True
                    End If
                End If
        End Select
        
        CanEventMoveTowardsPlayer = Random(0, 3)
    ElseIf PathfindingType = 2 Then
        ' Initialization phase
        tim = 0
        sX = x1
        sY = y1
        FX = x
        FY = Y
        
        ReDim pos(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
        
        pos = MapBlocks(MapNum).Blocks
        
        For i = 1 To tempplayer(PlayerID).EventMap.CurrentEvents
            If tempplayer(PlayerID).EventMap.EventPages(i).Visible Then
                If tempplayer(PlayerID).EventMap.EventPages(i).WalkThrough = 1 Then
                    pos(tempplayer(PlayerID).EventMap.EventPages(i).x, tempplayer(PlayerID).EventMap.EventPages(i).Y) = 9
                End If
            End If
        Next
        
        pos(sX, sY) = 100 + tim
        pos(FX, FY) = 2
        
        ' Reset reachable
        reachable = False
        
        'Do while reachable is false... if its set true in progress, we jump out
        ' If the path is decided unreachable in process, we will use exit sub. Not proper,
        ' But faster ;-)
        Do While reachable = False
            ' We loop through all squares
            For j = 0 To Map(MapNum).MaxY
                For i = 0 To Map(MapNum).MaxX
                    ' If j = 10 And i = 0 Then MsgBox "hi!"
                    ' If they are to be extended, the pointer TIM is on them
                    If pos(i, j) = 100 + tim Then
                    ' The part is to be extended, so do it
                        ' We have to make sure that there is a pos(i+1,j) BEFORE we actually use it,
                        ' Because then we get error... If the square is on side, we dont test for this one!
                        If i < Map(MapNum).MaxX Then
                            ' If there isnt a wall, or any other... thing
                            If pos(i + 1, j) = 0 Then
                                ' Expand it, and make its pos equal to tim+1, so the next time we make this loop,
                                ' It will exapand that square too! This is crucial part of the program
                                pos(i + 1, j) = 100 + tim + 1
                            ElseIf pos(i + 1, j) = 2 Then
                                ' If the position is no 0 but its 2 (FINISH) then Reachable = true!!! We found end
                                reachable = True
                            End If
                        End If
                    
                        ' This is the same as the last one, as i said a lot of copy paste work and editing that
                        ' This is simply another side that we have to test for... so instead of i+1 we have i-1
                        ' Its actually pretty same then... I wont comment it therefore, because its only repeating
                        ' Same thing with minor changes to check sides
                        If i > 0 Then
                            If pos((i - 1), j) = 0 Then
                                pos(i - 1, j) = 100 + tim + 1
                            ElseIf pos(i - 1, j) = 2 Then
                                reachable = True
                            End If
                        End If
                    
                        If j < Map(MapNum).MaxY Then
                            If pos(i, j + 1) = 0 Then
                                pos(i, j + 1) = 100 + tim + 1
                            ElseIf pos(i, j + 1) = 2 Then
                                reachable = True
                            End If
                        End If
                    
                        If j > 0 Then
                            If pos(i, j - 1) = 0 Then
                                pos(i, j - 1) = 100 + tim + 1
                            ElseIf pos(i, j - 1) = 2 Then
                                reachable = True
                            End If
                        End If
                    End If
                    DoEvents
                Next i
            Next j
            
            ' If the reachable is STILL false, then
            If reachable = False Then
                ' Reset sum
                Sum = 0
                For j = 0 To Map(MapNum).MaxY
                    For i = 0 To Map(MapNum).MaxX
                    ' We add up ALL the squares
                    Sum = Sum + pos(i, j)
                    Next i
                Next j
                
                'Now if the sum is euqal to the last sum, its not reachable, if it isnt, then we store
                ' Sum to lastsum
                If Sum = LastSum Then
                    CanEventMoveTowardsPlayer = 4
                    Exit Function
                Else
                    LastSum = Sum
                End If
            End If
            
            ' We increase the pointer to point to the next squares to be expanded
            tim = tim + 1
        Loop
        
        ' We work backwards to find the way...
        LastX = FX
        LastY = FY
        
        ReDim path(tim + 1)
        
        ' The following code may be a little bit confusing but ill try my best to explain it.
        ' We are working backwards to find ONE of the shortest ways back to Start.
        ' So we repeat the loop until the LastX and LastY arent in start. Look in the code to see
        ' How LastX and LasY change
        Do While LastX <> sX Or LastY <> sY
            ' We decrease tim by one, and then we are finding any adjacent square to the final one, that
            ' Has that value. So lets say the tim would be 5, because it takes 5 steps to get to the target.
            'Now everytime we decrease that, so we make it 4, and we look for any adjacent square that has
            ' That value. When we find it, we just color it yellow as for the solution
            
            tim = tim - 1
            ' Reset did to false
            did = False
            
            ' If we arent on edge
            If LastX < Map(MapNum).MaxX Then
                ' Check the square on the right of the solution. Is it a tim-1 one? or just a blank one
                If pos(LastX + 1, LastY) = 100 + tim Then
                    ' If it, then make it yellow, and change did to true
                    LastX = LastX + 1
                    did = True
                End If
            End If
            
            ' This will then only work if the previous part didnt execute, and did is still false. THen
            ' we want to check another square, the on left. Is it a tim-1 one ?
            If did = False Then
                If LastX > 0 Then
                    If pos(LastX - 1, LastY) = 100 + tim Then
                        LastX = LastX - 1
                        did = True
                    End If
                End If
            End If
            
            ' We check the one below it
            If did = False Then
                If LastY < Map(MapNum).MaxY Then
                    If pos(LastX, LastY + 1) = 100 + tim Then
                        LastY = LastY + 1
                        did = True
                    End If
                End If
            End If
            
            ' And above it. One of these have to be it, since we have found the solution, we know that already
            ' there is a way back.
            If did = False Then
                If LastY > 0 Then
                    If pos(LastX, LastY - 1) = 100 + tim Then
                        LastY = LastY - 1
                    End If
                End If
            End If
            
            path(tim).x = LastX
            path(tim).Y = LastY
            
            ' Now we loop back and decrease tim, and look for the next square with lower value
            DoEvents
        Loop
        
        ' Ok we got a path. Now, lets look at the first step and see what direction we should take.
        If path(1).x > LastX Then
            CanEventMoveTowardsPlayer = DIR_RIGHT
        ElseIf path(1).Y > LastY Then
            CanEventMoveTowardsPlayer = DIR_DOWN
        ElseIf path(1).Y < LastY Then
            CanEventMoveTowardsPlayer = DIR_UP
        ElseIf path(1).x < LastX Then
            CanEventMoveTowardsPlayer = DIR_LEFT
        End If
    End If
End Function

Function CanEventMoveAwayFromPlayer(PlayerID As Long, MapNum As Long, eventID As Long) As Long
    Dim i As Long, x As Long, Y As Long, x1 As Long, y1 As Long, DidWalk As Boolean, WalkThrough As Long
    
    ' This does not work for global events so this MUST be a player one....
    ' This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    CanEventMoveAwayFromPlayer = 5
    
    If PlayerID <= 0 Or PlayerID > Player_HighIndex Then Exit Function
    If MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > tempplayer(PlayerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(PlayerID)
    Y = GetPlayerY(PlayerID)
    x1 = tempplayer(PlayerID).EventMap.EventPages(eventID).x
    y1 = tempplayer(PlayerID).EventMap.EventPages(eventID).Y
    WalkThrough = Map(MapNum).Events(tempplayer(PlayerID).EventMap.EventPages(eventID).eventID).Pages(tempplayer(PlayerID).EventMap.EventPages(eventID).PageID).WalkThrough
    
    i = Int(Rnd * 5)
    DidWalk = False
    
    ' Lets move the event
    Select Case i
        Case 0
        ' Up
        If y1 > Y And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                CanEventMoveAwayFromPlayer = DIR_DOWN
                Exit Function
                DidWalk = True
            End If
        End If

        ' Down
        If y1 < Y And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                CanEventMoveAwayFromPlayer = DIR_UP
                Exit Function
                DidWalk = True
            End If
        End If

        ' Left
        If x1 > x And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                CanEventMoveAwayFromPlayer = DIR_RIGHT
                Exit Function
                DidWalk = True
            End If
        End If

        ' Right
        If x1 < x And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                CanEventMoveAwayFromPlayer = DIR_LEFT
                Exit Function
                DidWalk = True
            End If
        End If

    Case 1
        ' Right
        If x1 < x And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                CanEventMoveAwayFromPlayer = DIR_LEFT
                Exit Function
                DidWalk = True
            End If
        End If
        
        ' Left
        If x1 > x And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                CanEventMoveAwayFromPlayer = DIR_RIGHT
                Exit Function
                DidWalk = True
            End If
        End If
        
        ' Down
        If y1 < Y And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                CanEventMoveAwayFromPlayer = DIR_UP
                Exit Function
                DidWalk = True
            End If
        End If
        
        ' Up
        If y1 > Y And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                CanEventMoveAwayFromPlayer = DIR_DOWN
                Exit Function
                DidWalk = True
            End If
        End If

    Case 2
        ' Down
        If y1 < Y And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                CanEventMoveAwayFromPlayer = DIR_UP
                Exit Function
                DidWalk = True
            End If
        End If
        
        ' Up
        If y1 > Y And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                CanEventMoveAwayFromPlayer = DIR_DOWN
                Exit Function
                DidWalk = True
            End If
        End If
        
        ' Right
        If x1 < x And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                CanEventMoveAwayFromPlayer = DIR_LEFT
                Exit Function
                DidWalk = True
            End If
        End If
        
        ' Left
        If x1 > x And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                CanEventMoveAwayFromPlayer = DIR_RIGHT
                Exit Function
                DidWalk = True
            End If
        End If

    Case 3
        ' Left
        If x1 > x And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                CanEventMoveAwayFromPlayer = DIR_RIGHT
                Exit Function
                DidWalk = True
            End If
        End If
        
        ' Right
        If x1 < x And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                CanEventMoveAwayFromPlayer = DIR_LEFT
                Exit Function
                DidWalk = True
            End If
        End If
        
        ' Up
        If y1 > Y And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                CanEventMoveAwayFromPlayer = DIR_DOWN
                Exit Function
                DidWalk = True
            End If
        End If
        
        ' Down
        If y1 < Y And Not DidWalk Then
            If CanEventMove(PlayerID, MapNum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                CanEventMoveAwayFromPlayer = DIR_UP
                Exit Function
                DidWalk = True
            End If
        End If

    End Select
    
    CanEventMoveAwayFromPlayer = Random(0, 3)
End Function

Function GetDirToPlayer(PlayerID As Long, MapNum As Long, eventID As Long) As Long
    Dim i As Long, x As Long, Y As Long, x1 As Long, y1 As Long, DidWalk As Boolean, WalkThrough As Long, distance As Long
    
    ' This does not work for global events so this MUST be a player one....
    ' This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    If PlayerID <= 0 Or PlayerID > Player_HighIndex Then Exit Function
    If MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > tempplayer(PlayerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(PlayerID)
    Y = GetPlayerY(PlayerID)
    x1 = tempplayer(PlayerID).EventMap.EventPages(eventID).x
    y1 = tempplayer(PlayerID).EventMap.EventPages(eventID).Y
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_RIGHT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_LEFT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If Y - y1 > 0 Then
        If Y - y1 > distance Then
            i = DIR_DOWN
            distance = Y - y1
        End If
    ElseIf Y - y1 < 0 Then
        If ((Y - y1) * -1) > distance Then
            i = DIR_UP
            distance = ((Y - y1) * -1)
        End If
    End If
    
    GetDirToPlayer = i
End Function

Function GetDirAwayFromPlayer(PlayerID As Long, MapNum As Long, eventID As Long) As Long
    Dim i As Long, x As Long, Y As Long, x1 As Long, y1 As Long, DidWalk As Boolean, WalkThrough As Long, distance As Long
    
    ' This does not work for global events so this MUST be a player one....
    ' This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    If PlayerID <= 0 Or PlayerID > Player_HighIndex Then Exit Function
    If MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > tempplayer(PlayerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(PlayerID)
    Y = GetPlayerY(PlayerID)
    x1 = tempplayer(PlayerID).EventMap.EventPages(eventID).x
    y1 = tempplayer(PlayerID).EventMap.EventPages(eventID).Y
    
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_LEFT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_RIGHT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If Y - y1 > 0 Then
        If Y - y1 > distance Then
            i = DIR_UP
            distance = Y - y1
        End If
    ElseIf Y - y1 < 0 Then
        If ((Y - y1) * -1) > distance Then
            i = DIR_DOWN
            distance = ((Y - y1) * -1)
        End If
    End If
    
    GetDirAwayFromPlayer = i
End Function

Function GetNpcDir(x As Long, Y As Long, x1 As Long, y1 As Long) As Long
    Dim i As Long, distance As Long
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_RIGHT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_LEFT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If Y - y1 > 0 Then
        If Y - y1 > distance Then
            i = DIR_DOWN
            distance = Y - y1
        End If
    ElseIf Y - y1 < 0 Then
        If ((Y - y1) * -1) > distance Then
            i = DIR_UP
            distance = ((Y - y1) * -1)
        End If
    End If
    
    GetNpcDir = i
End Function

Public Sub UpdateMapBlock(ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long, ByRef blocked As Boolean)
    If blocked Then
        MapBlocks(MapNum).Blocks(x, Y) = 9
    Else
        MapBlocks(MapNum).Blocks(x, Y) = 0
    End If
End Sub

Public Sub CacheMapBlocks(ByVal MapNum As Integer)
    Dim x As Long, Y As Long
    
    ReDim Preserve MapBlocks(MapNum).Blocks(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    
    For x = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            If NPCTileIsOpen(MapNum, x, Y) = False Then
                MapBlocks(MapNum).Blocks(x, Y) = 9
            End If
        Next
    Next
End Sub

Function IsOneBlockAway(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Boolean
    If x1 = x2 Then
        If y1 = y2 - 1 Or y1 = y2 + 1 Then
            IsOneBlockAway = True
        Else
            IsOneBlockAway = False
        End If
    ElseIf y1 = y2 Then
        If x1 = x2 - 1 Or x1 = x2 + 1 Then
            IsOneBlockAway = True
        Else
            IsOneBlockAway = False
        End If
    Else
        IsOneBlockAway = False
    End If
End Function
