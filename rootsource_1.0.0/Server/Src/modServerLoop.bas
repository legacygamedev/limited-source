Attribute VB_Name = "modServerLoop"
Option Explicit

' ********************************************
' **               rootSource               **
' ********************************************

' halts thread of execution
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public ServerOnline As Boolean ' Used for server loop
Private GiveNpcHPTimer As Long  ' Used for Npc HP regeneration

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Private Secs As Long

Public Sub ServerLoop()
Dim i As Long
Dim X As Long
Dim y As Long

Dim Tick As Long

Dim tmr500 As Long
Dim tmr1000 As Long

Dim LastUpdateSavePlayers As Long
Dim LastUpdateMapSpawnItems As Long
Dim LastUpdatePlayerVitals As Long

Dim Buffer As clsBuffer

    Do While ServerOnline
        Tick = GetTickCount
        
        If Tick > tmr500 Then
            ' Check for disconnections
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            
            ' Process Npc AI
            UpdateNpcAI
           
            tmr500 = GetTickCount + 500
        End If
        
        If Tick > tmr1000 Then
            ' Handle shutting down server
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            ' Handles closing doors
            For i = 1 To MAX_MAPS
                If Tick > TempTile(i).DoorTimer + 5000 Then
                    For X = 0 To MAX_MAPX
                        For y = 0 To MAX_MAPY
                            If Map(i).Tile(X, y).Type = TILE_TYPE_KEY Then
                                If TempTile(i).DoorOpen(X, y) = YES Then
                                    TempTile(i).DoorOpen(X, y) = NO
                                    Set Buffer = New clsBuffer
                                    Buffer.PreAllocate 14
                                    Buffer.WriteInteger SMapKey
                                    Buffer.WriteLong X
                                    Buffer.WriteLong y
                                    Buffer.WriteLong 0
                                    Call SendDataToMap(i, Buffer.ToArray())
                                    Set Buffer = Nothing
                                End If
                            End If
                        Next
                    Next
                End If
            Next
            
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
Dim X As Long
Dim y As Long
    
    ' This is used for respawning map items
    For y = 1 To MAX_MAPS
        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(y) Then
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
    
End Sub

Private Sub UpdateNpcAI()
Dim i As Long, n As Long
Dim MapNum As Long, MapNpcNum As Long
Dim NpcNum As Long, Target As Long
Dim TickCount As Long
Dim Damage As Long
Dim DistanceX As Long
Dim DistanceY As Long
Dim DidWalk As Boolean
           
    For MapNum = 1 To MAX_MAPS
        If True Then
            TickCount = GetTickCount
           
            For MapNpcNum = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(MapNum, MapNpcNum).Num
               
                ' Make sure theres a Npc with the map
                If NpcNum > 0 Then
                   
                    ' Get the target
                    Target = MapNpc(MapNum, MapNpcNum).Target
                   
                    ' /////////////////////////////////////////
                    ' // This is used for ATTACKING ON SIGHT //
                    ' /////////////////////////////////////////
                    ' If the Npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behavior = Npc_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = Npc_BEHAVIOR_GUARD Then
                        ' First check if they don't have a target before looping...
                        If Target = 0 Then
                            For i = 1 To High_Index
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum Then
                                        If GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                            n = Npc(NpcNum).Range
                                           
                                            DistanceX = MapNpc(MapNum, MapNpcNum).X - GetPlayerX(i)
                                            DistanceY = MapNpc(MapNum, MapNpcNum).y - GetPlayerY(i)
                                           
                                            ' Make sure we get a positive value
                                            If DistanceX < 0 Then DistanceX = -DistanceX
                                            If DistanceY < 0 Then DistanceY = -DistanceY
                                           
                                            ' Are they in range?  if so GET'M!
                                            If DistanceX <= n Then
                                                If DistanceY <= n Then
                                                    If Npc(NpcNum).Behavior = Npc_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                        If LenB(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                                                            Call PlayerMsg(i, "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                                                        End If
                                                       
                                                        MapNpc(MapNum, MapNpcNum).Target = i
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
               
                                                                       
                    ' /////////////////////////////////////////////
                    ' // This is used for Npc walking/targetting //
                    ' /////////////////////////////////////////////
                    ' Check to see if its time for the Npc to walk
                    If Npc(NpcNum).Behavior <> Npc_BEHAVIOR_SHOPKEEPER Then
                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                            ' Check if the player is even playing, if so follow'm
                            If IsPlaying(Target) Then
                                If GetPlayerMap(Target) = MapNum Then
                                    DidWalk = False
                                   
                                    i = Int(Rnd * 4)
                                   
                                    ' Lets move the Npc
                                    Select Case i
                                        Case 0
                                            ' Up
                                            If MapNpc(MapNum, MapNpcNum).y > GetPlayerY(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Down
                                            If MapNpc(MapNum, MapNpcNum).y < GetPlayerY(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Left
                                            If MapNpc(MapNum, MapNpcNum).X > GetPlayerX(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Right
                                            If MapNpc(MapNum, MapNpcNum).X < GetPlayerX(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                       
                                        Case 1
                                            ' Right
                                            If MapNpc(MapNum, MapNpcNum).X < GetPlayerX(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Left
                                            If MapNpc(MapNum, MapNpcNum).X > GetPlayerX(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Down
                                            If MapNpc(MapNum, MapNpcNum).y < GetPlayerY(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Up
                                            If MapNpc(MapNum, MapNpcNum).y > GetPlayerY(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                           
                                        Case 2
                                            ' Down
                                            If MapNpc(MapNum, MapNpcNum).y < GetPlayerY(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Up
                                            If MapNpc(MapNum, MapNpcNum).y > GetPlayerY(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Right
                                            If MapNpc(MapNum, MapNpcNum).X < GetPlayerX(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Left
                                            If MapNpc(MapNum, MapNpcNum).X > GetPlayerX(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                       
                                        Case 3
                                            ' Left
                                            If MapNpc(MapNum, MapNpcNum).X > GetPlayerX(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Right
                                            If MapNpc(MapNum, MapNpcNum).X < GetPlayerX(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Up
                                            If MapNpc(MapNum, MapNpcNum).y > GetPlayerY(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                            ' Down
                                            If MapNpc(MapNum, MapNpcNum).y < GetPlayerY(Target) And Not DidWalk Then
                                                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then
                                                    Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                    End Select
                           
                                    ' Check if we can't move and if player is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNpc(MapNum, MapNpcNum).X - 1 = GetPlayerX(Target) And MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Target) Then
                                            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_LEFT Then
                                                Call NpcDir(MapNum, MapNpcNum, DIR_LEFT)
                                            End If
                                            DidWalk = True
                                        End If
                                    End If
                                   
                                    If Not DidWalk Then
                                        If MapNpc(MapNum, MapNpcNum).X + 1 = GetPlayerX(Target) And MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Target) Then
                                            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_RIGHT Then
                                                Call NpcDir(MapNum, MapNpcNum, DIR_RIGHT)
                                            End If
                                            DidWalk = True
                                        End If
                                    End If
                                   
                                    If Not DidWalk Then
                                        If MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Target) And MapNpc(MapNum, MapNpcNum).y - 1 = GetPlayerY(Target) Then
                                            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_UP Then
                                                Call NpcDir(MapNum, MapNpcNum, DIR_UP)
                                            End If
                                            DidWalk = True
                                        End If
                                    End If
                                   
                                    If Not DidWalk Then
                                        If MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Target) And MapNpc(MapNum, MapNpcNum).y + 1 = GetPlayerY(Target) Then
                                            If MapNpc(MapNum, MapNpcNum).Dir <> DIR_DOWN Then
                                                Call NpcDir(MapNum, MapNpcNum, DIR_DOWN)
                                            End If
                                            DidWalk = True
                                        End If
                                    End If
                                   
                                    ' We could not move so player must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
                                            If CanNpcMove(MapNum, MapNpcNum, i) Then
                                                Call NpcMove(MapNum, MapNpcNum, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                   
                                    ' /////////////////////////////////////////////
                                    ' // This is used for Npcs to attack players //
                                    ' /////////////////////////////////////////////
                                    ' Can the Npc attack the player?
                                    If CanNpcAttackPlayer(MapNpcNum, Target) Then
                                        If Not CanPlayerBlockHit(Target) Then
                                            Damage = Npc(NpcNum).Stat(Stats.Strength) - GetPlayerProtection(Target)
                                            Call NpcAttackPlayer(MapNpcNum, Target, Damage)
                                        Else
                                            Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerEquipmentSlot(Target, Shield))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                        End If
                                    End If
                                Else
                                    MapNpc(MapNum, MapNpcNum).Target = 0
                                End If
                            Else
                                MapNpc(MapNum, MapNpcNum).Target = 0
                            End If
                        Else
                            If Int(Rnd * 4) = 1 Then
                                i = Int(Rnd * 4)
                                If CanNpcMove(MapNum, MapNpcNum, i) Then
                                    Call NpcMove(MapNum, MapNpcNum, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
               
                    ' ////////////////////////////////////////////
                    ' // This is used for regenerating Npc's HP //
                    ' ////////////////////////////////////////////
                    If TickCount > GiveNpcHPTimer + 10000 Then
                        If MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) > 0 Then
                            MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) + GetNpcVitalRegen(NpcNum, Vitals.HP)
                       
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) > GetNpcMaxVital(NpcNum, Vitals.HP) Then
                                MapNpc(MapNum, MapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
                            End If
                        End If
                        GiveNpcHPTimer = TickCount
                    End If
                   
                End If
               
                ' //////////////////////////////////////
                ' // This is used for spawning an Npc //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an Npc or not
                If MapNpc(MapNum, MapNpcNum).Num = 0 Then
                    If Map(MapNum).Npc(MapNpcNum) > 0 Then
                        If TickCount > MapNpc(MapNum, MapNpcNum).SpawnWait + (Npc(Map(MapNum).Npc(MapNpcNum)).SpawnSecs * 1000) Then
                            Call SpawnNpc(MapNpcNum, MapNum)
                        End If
                    End If
                End If
               
            Next
        End If
        DoEvents
    Next

End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long

    For i = 1 To TotalPlayersOnline
        If GetPlayerVital(PlayersOnline(i), Vitals.HP) <> GetPlayerMaxVital(PlayersOnline(i), Vitals.HP) Then
            Call SetPlayerVital(PlayersOnline(i), Vitals.HP, GetPlayerVital(PlayersOnline(i), Vitals.HP) + GetPlayerVitalRegen(PlayersOnline(i), Vitals.HP))
            Call SendVital(PlayersOnline(i), Vitals.HP)
        End If
        If GetPlayerVital(PlayersOnline(i), Vitals.MP) <> GetPlayerMaxVital(PlayersOnline(i), Vitals.MP) Then
            Call SetPlayerVital(PlayersOnline(i), Vitals.MP, GetPlayerVital(PlayersOnline(i), Vitals.MP) + GetPlayerVitalRegen(PlayersOnline(i), Vitals.MP))
            Call SendVital(PlayersOnline(i), Vitals.MP)
        End If
        If GetPlayerVital(PlayersOnline(i), Vitals.SP) <> GetPlayerMaxVital(PlayersOnline(i), Vitals.SP) Then
            Call SetPlayerVital(PlayersOnline(i), Vitals.SP, GetPlayerVital(PlayersOnline(i), Vitals.SP) + GetPlayerVitalRegen(PlayersOnline(i), Vitals.SP))
            Call SendVital(PlayersOnline(i), Vitals.SP)
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
Dim i As Long

    If TotalPlayersOnline > 0 Then
        Call TextAdd("Saving all online players...")
        Call AdminMsg("Saving all online players...", Pink)
        
        For i = 1 To TotalPlayersOnline
            Call SavePlayer(PlayersOnline(i))
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

