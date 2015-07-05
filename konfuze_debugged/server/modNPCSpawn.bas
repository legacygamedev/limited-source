Attribute VB_Name = "modNPCSpawn"
Sub SpawnAttributeNpc(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal MapNum As Long)
Dim Packet As String
Dim NpcNum As Long
Dim i As Long
Dim Spawned As Boolean
Dim BX As Long, BY As Long
Dim BX2 As Long, BY2 As Long
Dim BX3 As Long, BY3 As Long

    If index > Map(MapNum).Tile(x, y).Data2 Then Exit Sub

    ' Check for subscript out of range
    If index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    Spawned = False
    
    NpcNum = Map(MapNum).Tile(x, y).Data1
    'If NpcNum > 0 Then
        If GameTime = TIME_NIGHT Then
            If Npc(NpcNum).SpawnTime = 1 Then
                MapAttributeNpc(MapNum, index, x, y).num = 0
                MapAttributeNpc(MapNum, index, x, y).SpawnWait = GetTickCount
                MapAttributeNpc(MapNum, index, x, y).HP = 0
                Call SendDataToMap(MapNum, "ATTRIBUTENPCDEAD" & SEP_CHAR & index & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        Else
            If Npc(NpcNum).SpawnTime = 2 Then
                MapAttributeNpc(MapNum, index, x, y).num = 0
                MapAttributeNpc(MapNum, index, x, y).SpawnWait = GetTickCount
                MapAttributeNpc(MapNum, index, x, y).HP = 0
                Call SendDataToMap(MapNum, "ATTRIBUTENPCDEAD" & SEP_CHAR & index & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
    
        MapAttributeNpc(MapNum, index, x, y).num = NpcNum
        MapAttributeNpc(MapNum, index, x, y).Target = 0
        
        MapAttributeNpc(MapNum, index, x, y).HP = GetNpcMaxHP(NpcNum)
        MapAttributeNpc(MapNum, index, x, y).MP = GetNpcMaxMP(NpcNum)
        MapAttributeNpc(MapNum, index, x, y).SP = GetNpcMaxSP(NpcNum)
                
        MapAttributeNpc(MapNum, index, x, y).Dir = Int(Rnd * 4)
        
        ' Well try 100 times to randomly place the sprite
        If Map(MapNum).Tile(x, y).Data3 > 0 Then
        BX3 = x + Map(MapNum).Tile(x, y).Data3
        BX2 = x - Map(MapNum).Tile(x, y).Data3
        BY3 = y + Map(MapNum).Tile(x, y).Data3
        BY2 = y - Map(MapNum).Tile(x, y).Data3
        
        If BX2 < 0 Then BX2 = 1
        If BX3 > MAX_MAPX Then BX3 = MAX_MAPX
        If BY2 < 0 Then BY2 = 1
        If BY3 > MAX_MAPY Then BY3 = MAX_MAPY
        
            For i = 1 To 100
                BX = Int(Rand(BX3, BX2))
                BY = Int(Rand(BY3, BY2))
                
                BX = BX - 1
                BY = BY - 1
    
                ' Check if the tile is walkable
                If Map(MapNum).Tile(BX, BY).Type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(BX, BY).Type = TILE_TYPE_NPC_SPAWN Then
                    MapAttributeNpc(MapNum, index, x, y).x = BX
                    MapAttributeNpc(MapNum, index, x, y).y = BY
                    Spawned = True
                    Exit For
                End If
            Next i
            
            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then
                For BY = BY2 To BY3
                    For BX = BX2 To BX3
                        If Map(MapNum).Tile(BX, BY).Type = TILE_TYPE_WALKABLE Or Map(MapNum).Tile(BX, BY).Type = TILE_TYPE_NPC_SPAWN Then
                            MapAttributeNpc(MapNum, index, x, y).x = BX
                            MapAttributeNpc(MapNum, index, x, y).y = BY
                            Spawned = True
                        End If
                    Next BX
                Next BY
            End If
        Else
            MapAttributeNpc(MapNum, index, x, y).x = x
            MapAttributeNpc(MapNum, index, x, y).y = y
            Spawned = True
        End If
             
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = "spawnattributenpc" & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).num & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).x & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).y & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).Dir & SEP_CHAR & Npc(MapAttributeNpc(MapNum, index, x, y).num).Big & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        End If
    'End If
    
    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).HP & SEP_CHAR & GetNpcMaxHP(MapAttributeNpc(MapNum, index, x, y).num) & SEP_CHAR & END_CHAR)
End Sub

Sub SpawnMapAttributeNpcs(ByVal MapNum As Long)
Dim i As Long, x As Long, y As Long
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                For i = 1 To MAX_ATTRIBUTE_NPCS
                    Call SpawnAttributeNpc(i, x, y, MapNum)
                Next i
            End If
        Next x
    Next y
End Sub

Sub SpawnAllMapAttributeNpcs()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapAttributeNpcs(i)
    Next i
End Sub

Sub AttributeNPCGameAI(ByVal MapNum As Long)
Dim i As Long, x As Long, y As Long, n As Long, x1 As Long, y1 As Long, d As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long, Target As Long
Dim DidWalk As Boolean


For y = 0 To MAX_MAPY
    For x = 0 To MAX_MAPX
        If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
            For n = 1 To MAX_ATTRIBUTE_NPCS
                If n <= Map(MapNum).Tile(x, y).Data2 Then
                    NpcNum = MapAttributeNpc(MapNum, n, x, y).num

                    ' /////////////////////////////////////////
                    ' // This is used for ATTACKING ON SIGHT //
                    ' /////////////////////////////////////////
                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        For d = 1 To MAX_PLAYERS
                            If IsPlaying(d) Then
                                If GetPlayerMap(d) = y And MapAttributeNpc(MapNum, n, x, y).Target = 0 And GetPlayerAccess(d) <= ADMIN_MONITER Then
                                    n = Npc(NpcNum).Range
                                    
                                    DistanceX = MapAttributeNpc(MapNum, n, x, y).x - GetPlayerX(d)
                                    DistanceY = MapAttributeNpc(MapNum, n, x, y).y - GetPlayerY(d)
                                    
                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                                                Call PlayerMsg(d, "A " & Trim(Npc(NpcNum).Name) & " : " & Trim(Npc(NpcNum).AttackSay) & "", SayColor)
                                            End If
                                            
                                            MapAttributeNpc(MapNum, n, x, y).Target = d
                                        End If
                                    End If
                                End If
                            End If
                        Next d
                    End If
                                                                            
                    ' /////////////////////////////////////////////
                    ' // This is used for NPC walking/targetting //
                    ' /////////////////////////////////////////////
                    Target = MapAttributeNpc(MapNum, n, x, y).Target
                    
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
                                        If MapAttributeNpc(MapNum, n, x, y).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_UP) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapAttributeNpc(MapNum, n, x, y).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_DOWN) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapAttributeNpc(MapNum, n, x, y).x > GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_LEFT) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapAttributeNpc(MapNum, n, x, y).x < GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_RIGHT) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 1
                                        ' Right
                                        If MapAttributeNpc(MapNum, n, x, y).x < GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_RIGHT) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapAttributeNpc(MapNum, n, x, y).x > GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_LEFT) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapAttributeNpc(MapNum, n, x, y).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_DOWN) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapAttributeNpc(MapNum, n, x, y).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_UP) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        
                                    Case 2
                                        ' Down
                                        If MapAttributeNpc(MapNum, n, x, y).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_DOWN) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapAttributeNpc(MapNum, n, x, y).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_UP) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapAttributeNpc(MapNum, n, x, y).x < GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_RIGHT) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapAttributeNpc(MapNum, n, x, y).x > GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_LEFT) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 3
                                        ' Left
                                        If MapAttributeNpc(MapNum, n, x, y).x > GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_LEFT) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapAttributeNpc(MapNum, n, x, y).x < GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_RIGHT) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapAttributeNpc(MapNum, n, x, y).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_UP) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapAttributeNpc(MapNum, n, x, y).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, x, y, MapNum, DIR_DOWN) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                End Select
                                
                                
                            
                                ' Check if we can't move and if player is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapAttributeNpc(MapNum, n, x, y).x - 1 = GetPlayerX(Target) And MapAttributeNpc(MapNum, n, x, y).y = GetPlayerY(Target) Then
                                        If MapAttributeNpc(MapNum, n, x, y).Dir <> DIR_LEFT Then
                                            Call AttributeNpcDir(n, x, y, MapNum, DIR_LEFT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapAttributeNpc(MapNum, n, x, y).x + 1 = GetPlayerX(Target) And MapAttributeNpc(MapNum, n, x, y).y = GetPlayerY(Target) Then
                                        If MapAttributeNpc(MapNum, n, x, y).Dir <> DIR_RIGHT Then
                                            Call AttributeNpcDir(n, x, y, MapNum, DIR_RIGHT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapAttributeNpc(MapNum, n, x, y).x = GetPlayerX(Target) And MapAttributeNpc(MapNum, n, x, y).y - 1 = GetPlayerY(Target) Then
                                        If MapAttributeNpc(MapNum, n, x, y).Dir <> DIR_UP Then
                                            Call AttributeNpcDir(n, x, y, MapNum, DIR_UP)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapAttributeNpc(MapNum, n, x, y).x = GetPlayerX(Target) And MapAttributeNpc(MapNum, n, x, y).y + 1 = GetPlayerY(Target) Then
                                        If MapAttributeNpc(MapNum, n, x, y).Dir <> DIR_DOWN Then
                                            Call AttributeNpcDir(n, x, y, MapNum, DIR_DOWN)
                                        End If
                                        DidWalk = True
                                    End If
                                    
                                    ' We could not move so player must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
                                            If CanAttributeNPCMove(n, x, y, MapNum, i) Then
                                                Call AttributeNpcMove(n, x, y, MapNum, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                MapAttributeNpc(MapNum, n, x, y).Target = 0
                            End If
                        Else
                            i = Int(Rnd * 4)
                            If i = 1 Then
                                i = Int(Rnd * 4)
                                If CanAttributeNPCMove(n, x, y, MapNum, i) Then
                                    Call AttributeNpcMove(n, x, y, MapNum, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                    
                    ' /////////////////////////////////////////////
                    ' // This is used for npcs to attack players //
                    ' /////////////////////////////////////////////
                    Target = MapAttributeNpc(MapNum, n, x, y).Target
                    
                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                        ' Is the target playing and on the same map?
                        If IsPlaying(Target) And GetPlayerMap(Target) = y Then
                            ' Can the npc attack the player?
                            If CanAttributeNpcAttackPlayer(x, Target) Then
                                If Not CanPlayerBlockHit(Target) Then
                                    Damage = Npc(NpcNum).STR - GetPlayerProtection(Target)
                                    If Damage > 0 Then
                                        Call NpcAttackPlayer(x, Target, Damage)
                                    Else
                                        Call BattleMsg(Target, "The " & Trim(Npc(NpcNum).Name) & " couldn't hurt you!", BrightBlue, 1)
                                        
                                        'Call PlayerMsg(Target, "The " & Trim(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                    End If
                                Else
                                    Call BattleMsg(Target, "You blocked the " & Trim(Npc(NpcNum).Name) & "'s hit!", BrightCyan, 1)
                                    
                                    'Call PlayerMsg(Target, "Your " & Trim(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                End If
                            End If
                        Else
                            ' Player left map or game, set target to 0
                            MapAttributeNpc(MapNum, n, x, y).Target = 0
                        End If
                    End If
                    
                    ' ////////////////////////////////////////////
                    ' // This is used for regenerating NPC's HP //
                    ' ////////////////////////////////////////////
                    ' Check to see if we want to regen some of the npc's hp
                    If GetTickCount > GiveNPCHPTimer + 10000 Then
                        If MapAttributeNpc(MapNum, n, x, y).HP > 0 Then
                            MapAttributeNpc(MapNum, n, x, y).HP = MapAttributeNpc(MapNum, n, x, y).HP + GetNpcHPRegen(NpcNum)
                        
                            ' Check if they have more then they should and if so just set it to max
                            If MapAttributeNpc(MapNum, n, x, y).HP > GetNpcMaxHP(NpcNum) Then
                                MapAttributeNpc(MapNum, n, x, y).HP = GetNpcMaxHP(NpcNum)
                            End If
                        End If
                    End If
                        
                    ' ////////////////////////////////////////////////////////
                    ' // This is used for checking if an NPC is dead or not //
                    ' ////////////////////////////////////////////////////////
                    ' Check if the npc is dead or not
                    If NpcNum > 0 Then
                        If MapAttributeNpc(MapNum, n, x, y).HP <= 0 And GetNpcMaxHP(NpcNum) > 0 Then
                            MapAttributeNpc(MapNum, n, x, y).num = 0
                            MapAttributeNpc(MapNum, n, x, y).SpawnWait = GetTickCount
                        End If
                    End If
                                        
                    ' //////////////////////////////////////
                    ' // This is used for spawning an NPC //
                    ' //////////////////////////////////////
                    ' Check if we are supposed to spawn an npc or not
                    If NpcNum <= 0 Then
                        If GetTickCount > MapAttributeNpc(MapNum, n, x, y).SpawnWait + (Npc(Map(MapNum).Tile(x, y).Data1).SpawnSecs * 1000) Then
                            Call SpawnAttributeNpc(n, x, y, MapNum)
                        End If
                    End If
                    Call SendDataToMap(MapNum, "attributenpchp" & SEP_CHAR & n & SEP_CHAR & MapAttributeNpc(MapNum, n, x, y).HP & SEP_CHAR & GetNpcMaxHP(MapAttributeNpc(MapNum, n, x, y).num) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                End If
            Next n
        End If
    Next x
Next y
End Sub

Sub AttackAttributeNpcs(ByVal index As Long)
Dim i As Long, x As Long, y As Long, n As Long, NpcNum As Long, MapNum As Long, Damage As Long

MapNum = GetPlayerMap(index)

For y = 0 To MAX_MAPY
    For x = 0 To MAX_MAPX
        If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
            For i = 1 To MAX_ATTRIBUTE_NPCS
                If i <= Map(MapNum).Tile(x, y).Data2 Then
                
                    NpcNum = MapAttributeNpc(MapNum, i, x, y).num
                    
                    ' Can we attack the npc?
                    If CanAttackAttributeNpc(index, i, x, y) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(index) Then
                            Damage = GetPlayerDamage(index) - Int(Npc(NpcNum).DEF / 2)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            n = GetPlayerDamage(index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(NpcNum).DEF / 2)
                            Call BattleMsg(index, "You feel a surge of energy upon swinging!", BrightCyan, 0)
                            
                            'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackAttributeNpc(i, x, y, index, Damage)
                            'Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                        Else
                            Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)
                            
                            'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                            'Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                        Exit Sub
                    End If
                End If
            Next i
        End If
    Next x
Next y
End Sub

Function CanAttackAttributeNpc(ByVal Attacker As Long, ByVal index As Long, ByVal x As Long, ByVal y As Long) As Boolean
Dim AttackSpeed As Long
Dim n As Long, NpcNum As Long, MapNum As Long

If GetPlayerWeaponSlot(Attacker) > 0 Then
    AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
Else
    AttackSpeed = 1000
End If

CanAttackAttributeNpc = False
 
' Check for subscript out of range
If IsPlaying(Attacker) = False Or index <= 0 Or index > MAX_ATTRIBUTE_NPCS Then Exit Function
 
MapNum = GetPlayerMap(Attacker)
' Check for subscript out of range
'If MapAttributeNpc(MapNum, index, x, y).num <= 0 Then Exit Function

NpcNum = Map(MapNum).Tile(x, y).Data1
 
' Make sure the npc isn't already dead
'If MapAttributeNpc(MapNum, index, x, y).HP <= 0 Then Exit Function
 
' Make sure they are on the same map
'If IsPlaying(Attacker) Then
    If GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (MapAttributeNpc(MapNum, index, x, y).y + 1 = GetPlayerY(Attacker)) And (MapAttributeNpc(MapNum, index, x, y).x = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackAttributeNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                    Exit Function
                End If
 
            Case DIR_DOWN
                If (MapAttributeNpc(MapNum, index, x, y).y - 1 = GetPlayerY(Attacker)) And (MapAttributeNpc(MapNum, index, x, y).x = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackAttributeNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                    Exit Function
                End If
 
            Case DIR_LEFT
                If (MapAttributeNpc(MapNum, index, x, y).y = GetPlayerY(Attacker)) And (MapAttributeNpc(MapNum, index, x, y).x + 1 = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackAttributeNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                    Exit Function
                End If
 
            Case DIR_RIGHT
                If (MapAttributeNpc(MapNum, index, x, y).y = GetPlayerY(Attacker)) And (MapAttributeNpc(MapNum, index, x, y).x - 1 = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackAttributeNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                    Exit Function
                End If
        End Select
    End If
'End If
End Function

Function CanAttributeNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
    
    CanAttributeNpcAttackPlayer = False
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_ATTRIBUTE_NPCS Or IsPlaying(index) = False Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index), MapNpcNum).num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(index)
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If Player(index).GettingMap = YES Then
        Exit Function
    End If
    
    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NpcNum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).x) Then
                CanAttributeNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).x) Then
                    CanAttributeNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) + 1 = MapNpc(MapNum, MapNpcNum).x) Then
                        CanAttributeNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
                            CanAttributeNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If

'            Select Case MapNpc(MapNum, MapNpcNum).Dir
'                Case DIR_UP
'                    If (GetPlayerY(Index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanAttributeNpcAttackPlayer = True
'                    End If
'
'                Case DIR_DOWN
'                    If (GetPlayerY(Index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanAttributeNpcAttackPlayer = True
'                    End If
'
'                Case DIR_LEFT
'                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanAttributeNpcAttackPlayer = True
'                    End If
'
'                Case DIR_RIGHT
'                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanAttributeNpcAttackPlayer = True
'                    End If
'            End Select
        End If
    End If
End Function

Sub ClearMapAttributeNpc(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal MapNum As Long)
    MapAttributeNpc(MapNum, index, x, y).num = 0
    MapAttributeNpc(MapNum, index, x, y).Target = 0
    MapAttributeNpc(MapNum, index, x, y).HP = 0
    MapAttributeNpc(MapNum, index, x, y).MP = 0
    MapAttributeNpc(MapNum, index, x, y).SP = 0
    MapAttributeNpc(MapNum, index, x, y).x = 0
    MapAttributeNpc(MapNum, index, x, y).y = 0
    MapAttributeNpc(MapNum, index, x, y).Dir = 0
    
    ' Server use only
    MapAttributeNpc(MapNum, index, x, y).SpawnWait = 0
    MapAttributeNpc(MapNum, index, x, y).AttackTimer = 0
End Sub

Sub ClearMapAttributeNpcs()
Dim x As Long
Dim y As Long
Dim i As Long
Dim MapNum As Long

For MapNum = 1 To MAX_MAPS
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                For i = 1 To MAX_ATTRIBUTE_NPCS
                    If i <= Map(MapNum).Tile(x, y).Data2 Then
                        Call ClearMapAttributeNpc(i, x, y, MapNum)
                    End If
                Next i
            End If
        Next x
    Next y
Next MapNum
End Sub

Sub SendMapAttributeNpcsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long, n As Long, x As Long, y As Long

    If MapNum > 0 Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                    Packet = "MAPATTRIBUTENPCDATA" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR
                
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        'If i <= Map(MapNum).Tile(X, Y).Data2 Then
                            Packet = Packet & MapAttributeNpc(MapNum, i, x, y).num & SEP_CHAR & MapAttributeNpc(MapNum, i, x, y).x & SEP_CHAR & MapAttributeNpc(MapNum, i, x, y).y & SEP_CHAR & MapAttributeNpc(MapNum, i, x, y).Dir & SEP_CHAR
                        'End If
                    Next i
                    
                    Packet = Packet & END_CHAR
                    
                    Call SendDataTo(index, Packet)
                End If
            Next x
        Next y
    End If
End Sub

Sub SendMapAttributeNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long, n As Long, x As Long, y As Long

    If MapNum > 0 Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                    Packet = "MAPATTRIBUTENPCDATA" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR
                
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        If i <= Map(MapNum).Tile(x, y).Data2 Then
                            Packet = Packet & MapAttributeNpc(MapNum, i, x, y).num & SEP_CHAR & MapAttributeNpc(MapNum, i, x, y).x & SEP_CHAR & MapAttributeNpc(MapNum, i, x, y).y & SEP_CHAR & MapAttributeNpc(MapNum, i, x, y).Dir & SEP_CHAR
                        End If
                    Next i
                    
                    Packet = Packet & END_CHAR
                    
                    Call SendDataToMap(MapNum, Packet)
                End If
            Next x
        Next y
    End If
End Sub

Function CanAttributeNPCMove(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal MapNum As Long, ByVal Dir) As Boolean
Dim i As Long, n As Long
Dim BX As Long, BY As Long
Dim BX2 As Long, BY2 As Long

    CanAttributeNPCMove = False
        
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    
    If index > Map(MapNum).Tile(x, y).Data2 Then Exit Function
    
    BX = MapAttributeNpc(MapNum, index, x, y).x
    BY = MapAttributeNpc(MapNum, index, x, y).y
    
    CanAttributeNPCMove = True
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If BY > 0 Then
                n = Map(MapNum).Tile(BX, BY - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                If CanAttributeNPCMovePlayer(MapNum, index, x, y, DIR_UP) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                If CanAttributeNPCMoveAttributeNPC(MapNum, index, x, y, DIR_UP) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                If CanAttributeNPCMoveNPC(MapNum, index, x, y, DIR_UP) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
            Else
                CanAttributeNPCMove = False
            End If
                
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If BY < MAX_MAPY Then
                n = Map(MapNum).Tile(BX, BY + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                If CanAttributeNPCMovePlayer(MapNum, index, x, y, DIR_DOWN) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                If CanAttributeNPCMoveAttributeNPC(MapNum, index, x, y, DIR_DOWN) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                If CanAttributeNPCMoveNPC(MapNum, index, x, y, DIR_DOWN) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
            Else
                CanAttributeNPCMove = False
            End If
                
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If BX > 0 Then
                n = Map(MapNum).Tile(BX - 1, BY).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                If CanAttributeNPCMovePlayer(MapNum, index, x, y, DIR_LEFT) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                                
                If CanAttributeNPCMoveAttributeNPC(MapNum, index, x, y, DIR_LEFT) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                If CanAttributeNPCMoveNPC(MapNum, index, x, y, DIR_LEFT) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
            Else
                CanAttributeNPCMove = False
            End If
                
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If BX < MAX_MAPX Then
                n = Map(MapNum).Tile(BX + 1, BY).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                If CanAttributeNPCMovePlayer(MapNum, index, x, y, DIR_RIGHT) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                                
                If CanAttributeNPCMoveAttributeNPC(MapNum, index, x, y, DIR_RIGHT) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
                
                If CanAttributeNPCMoveNPC(MapNum, index, x, y, DIR_RIGHT) = False Then
                    CanAttributeNPCMove = False
                    Exit Function
                End If
            Else
                CanAttributeNPCMove = False
            End If
    End Select
End Function

Function CanAttributeNPCMovePlayer(ByVal MapNum As Long, ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Dir As Long) As Boolean
Dim i As Long

CanAttributeNPCMovePlayer = True

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) Then
        If GetPlayerMap(i) = MapNum Then
            Select Case Dir
                Case DIR_UP
                    If (MapAttributeNpc(MapNum, index, x, y).x = GetPlayerX(i)) And (MapAttributeNpc(MapNum, index, x, y).y - 1 = GetPlayerY(i)) Then
                        CanAttributeNPCMovePlayer = False
                        Exit Function
                    End If
                Case DIR_DOWN
                    If (MapAttributeNpc(MapNum, index, x, y).x = GetPlayerX(i)) And (MapAttributeNpc(MapNum, index, x, y).y + 1 = GetPlayerY(i)) Then
                        CanAttributeNPCMovePlayer = False
                        Exit Function
                    End If
                Case DIR_LEFT
                    If (MapAttributeNpc(MapNum, index, x, y).x - 1 = GetPlayerX(i)) And (MapAttributeNpc(MapNum, index, x, y).y = GetPlayerY(i)) Then
                        CanAttributeNPCMovePlayer = False
                        Exit Function
                    End If
                Case DIR_RIGHT
                    If (MapAttributeNpc(MapNum, index, x, y).x + 1 = GetPlayerX(i)) And (MapAttributeNpc(MapNum, index, x, y).y = GetPlayerY(i)) Then
                        CanAttributeNPCMovePlayer = False
                        Exit Function
                    End If
            End Select
        End If
    End If
Next i
End Function

Function CanAttributeNPCMoveNPC(ByVal MapNum As Long, ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Dir As Long) As Boolean
Dim i As Long

CanAttributeNPCMoveNPC = True

For i = 1 To MAX_MAP_NPCS
    If MapNpc(MapNum, i).num > 0 Then
        Select Case Dir
            Case DIR_UP
                If (MapAttributeNpc(MapNum, index, x, y).x = MapNpc(MapNum, i).x) And (MapAttributeNpc(MapNum, index, x, y).y - 1 = MapNpc(MapNum, i).y) Then
                    CanAttributeNPCMoveNPC = False
                    Exit Function
                End If
            Case DIR_DOWN
                If (MapAttributeNpc(MapNum, index, x, y).x = MapNpc(MapNum, i).x) And (MapAttributeNpc(MapNum, index, x, y).y + 1 = MapNpc(MapNum, i).y) Then
                    CanAttributeNPCMoveNPC = False
                    Exit Function
                End If
            Case DIR_LEFT
                If (MapAttributeNpc(MapNum, index, x, y).x - 1 = MapNpc(MapNum, i).x) And (MapAttributeNpc(MapNum, index, x, y).y = MapNpc(MapNum, i).y) Then
                    CanAttributeNPCMoveNPC = False
                    Exit Function
                End If
            Case DIR_RIGHT
                If (MapAttributeNpc(MapNum, index, x, y).x + 1 = MapNpc(MapNum, i).x) And (MapAttributeNpc(MapNum, index, x, y).y = MapNpc(MapNum, i).y) Then
                    CanAttributeNPCMoveNPC = False
                    Exit Function
                End If
        End Select
    End If
Next i
End Function

Function CanAttributeNPCMoveAttributeNPC(ByVal MapNum As Long, ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Dir As Long) As Boolean
Dim i As Long, BX As Long, BY As Long

CanAttributeNPCMoveAttributeNPC = True

For BX = 0 To MAX_MAPX
    For BY = 0 To MAX_MAPY
        If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
            For i = 1 To MAX_ATTRIBUTE_NPCS
                If i <> index Then
                    If MapAttributeNpc(MapNum, i, BX, BY).num > 0 Then
                        Select Case Dir
                            Case DIR_UP
                                If (MapAttributeNpc(MapNum, index, x, y).x = MapAttributeNpc(MapNum, i, BX, BY).x) And (MapAttributeNpc(MapNum, index, x, y).y - 1 = MapAttributeNpc(MapNum, i, BX, BY).y) Then
                                    CanAttributeNPCMoveAttributeNPC = False
                                    Exit Function
                                End If
                            Case DIR_DOWN
                                If (MapAttributeNpc(MapNum, index, x, y).x = MapAttributeNpc(MapNum, i, BX, BY).x) And (MapAttributeNpc(MapNum, index, x, y).y + 1 = MapAttributeNpc(MapNum, i, BX, BY).y) Then
                                    CanAttributeNPCMoveAttributeNPC = False
                                    Exit Function
                                End If
                            Case DIR_LEFT
                                If (MapAttributeNpc(MapNum, index, x, y).x - 1 = MapAttributeNpc(MapNum, i, BX, BY).x) And (MapAttributeNpc(MapNum, index, x, y).y = MapAttributeNpc(MapNum, i, BX, BY).y) Then
                                    CanAttributeNPCMoveAttributeNPC = False
                                    Exit Function
                                End If
                            Case DIR_RIGHT
                                If (MapAttributeNpc(MapNum, index, x, y).x + 1 = MapAttributeNpc(MapNum, i, BX, BY).x) And (MapAttributeNpc(MapNum, index, x, y).y = MapAttributeNpc(MapNum, i, BX, BY).y) Then
                                    CanAttributeNPCMoveAttributeNPC = False
                                    Exit Function
                                End If
                        End Select
                    End If
                End If
            Next i
        End If
    Next BY
Next BX
End Function

Sub AttributeNpcMove(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal MapNum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim i As Long

    If index > Map(MapNum).Tile(x, y).Data2 Then Exit Sub

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    MapAttributeNpc(MapNum, index, x, y).Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            MapAttributeNpc(MapNum, index, x, y).y = MapAttributeNpc(MapNum, index, x, y).y - 1
            Packet = "attributenpcmove" & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).x & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).y & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).Dir & SEP_CHAR & Movement & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)

        Case DIR_DOWN
            MapAttributeNpc(MapNum, index, x, y).y = MapAttributeNpc(MapNum, index, x, y).y + 1
            Packet = "attributenpcmove" & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).x & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).y & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).Dir & SEP_CHAR & Movement & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_LEFT
            MapAttributeNpc(MapNum, index, x, y).x = MapAttributeNpc(MapNum, index, x, y).x - 1
            Packet = "attributenpcmove" & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).x & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).y & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).Dir & SEP_CHAR & Movement & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_RIGHT
            MapAttributeNpc(MapNum, index, x, y).x = MapAttributeNpc(MapNum, index, x, y).x + 1
            Packet = "attributenpcmove" & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).x & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).y & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).Dir & SEP_CHAR & Movement & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    End Select
End Sub

Sub AttributeNpcDir(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal MapNum As Long, ByVal Dir As Long)
Dim Packet As String

    If index > Map(MapNum).Tile(x, y).Data2 Then Exit Sub

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    MapAttributeNpc(MapNum, index, x, y).Dir = Dir
    Packet = "ATTRIBUTENPCDIR" & SEP_CHAR & index & SEP_CHAR & Dir & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub AttributeNpcAttackPlayer(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim MapNum As Long

    ' Check for subscript out of range
    If index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim), index).num <= 0 Then
        Exit Sub
    End If
        
    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "ATTRIBUTENPCATTACK" & SEP_CHAR & index & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Victim)
    
    ':: AUTO TURN ::
    'If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "AutoTurn")) = 1 Then
        'If GetPlayerX(Victim) - 1 = MapNpc(MapNum, index).X Then
            'Call SetPlayerDir(Victim, DIR_LEFT)
        'End If
        'If GetPlayerX(Victim) + 1 = MapNpc(MapNum, index).X Then
            'Call SetPlayerDir(Victim, DIR_RIGHT)
        'End If
        'If GetPlayerY(Victim) - 1 = MapNpc(MapNum, index).Y Then
            'Call SetPlayerDir(Victim, DIR_UP)
        'End If
        'If GetPlayerY(Victim) + 1 = MapNpc(MapNum, index).Y Then
            'Call SetPlayerDir(Victim, DIR_DOWN)
        'End If
        'Call SendDataToMap(GetPlayerMap(Victim), "changedir" & SEP_CHAR & GetPlayerDir(Victim) & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    'End If
    ':: END AUTO TURN ::
    
    Name = Trim(Npc(MapNpc(MapNum, index).num).Name)
    
    If Damage >= GetPlayerHP(Victim) Then
        ' Say damage
        Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)
        
        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by a " & Name, BrightRed)
        
        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "DropItems " & Victim
            Else
                If GetPlayerWeaponSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
                End If
            
                If GetPlayerArmorSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
                End If
                
                If GetPlayerHelmetSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
                End If
            
                If GetPlayerShieldSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
                End If
            End If
            
            ' Calculate exp to give attacker
            Exp = Int(GetPlayerExp(Victim) / 3)
            
            ' Make sure we dont get less then 0
            If Exp < 0 Then
                Exp = 0
            End If
            
            If Exp = 0 Then
                Call BattleMsg(Victim, "You lost no experience.", BrightRed, 0)
            Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call BattleMsg(Victim, "You lost " & Exp & " experience.", BrightRed, 0)
            End If
        End If
                
        ' Warp player away
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
        Else
            Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        End If
        
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
        
        ' Set NPC target to 0
        MapNpc(MapNum, index).Target = 0
        
        ' If the player the attacker killed was a pk then take it away
        If GetPlayerPK(Victim) = YES Then
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Say damage
        Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)
        
        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
    End If
    
    Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & END_CHAR)
End Sub

Sub AttackAttributeNpc(ByVal MapNpcNum As Long, ByVal x As Long, ByVal y As Long, ByVal Attacker As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim n As Long, i As Long, q As Integer, d As Long
Dim STR As Long, DEF As Long, MapNum As Long, NpcNum As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_ATTRIBUTE_NPCS Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    Else
        n = 0
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapAttributeNpc(MapNum, MapNpcNum, x, y).num
    Name = Trim(Npc(NpcNum).Name)
        
    If Damage >= MapAttributeNpc(MapNum, MapNpcNum, x, y).HP Then
        ' Check for a weapon and say damage
        
        Call BattleMsg(Attacker, "You killed a " & Name, BrightRed, 0)

        Dim add As String

        add = 0
        If GetPlayerWeaponSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AddEXP
        End If
        If GetPlayerArmorSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerArmorSlot(Attacker))).AddEXP
        End If
        If GetPlayerShieldSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerShieldSlot(Attacker))).AddEXP
        End If
        If GetPlayerHelmetSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerHelmetSlot(Attacker))).AddEXP
        End If
        
        If add > 0 Then
            If add < 100 Then
                If add < 10 Then
                    add = 0 & ".0" & Right(add, 2)
                Else
                    add = 0 & "." & Right(add, 2)
                End If
            Else
                add = Mid(add, 1, 1) & "." & Right(add, 2)
            End If
        End If
                                
        ' Calculate exp to give attacker
        If add > 0 Then
            Exp = Npc(NpcNum).Exp + (Npc(NpcNum).Exp * Val(add))
        Else
            Exp = Npc(NpcNum).Exp
        End If
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If

        ' Check if in party, if so divide the exp up by 2
        If Player(Attacker).InParty = False Or Player(Attacker).Party.ShareExp = False Then
            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "You cant gain anymore experience!", BrightBlue, 0)
            Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "You have gained " & Exp & " experience.", BrightBlue, 0)
            End If
        Else
            q = 0
            'The following code will tell us how many party members we have active
            For d = 1 To MAX_PARTY_MEMBERS
            If Player(Attacker).Party.Member(d) > 0 Then q = q + 1
            Next d
            'MsgBox "in party" & q
            If q = 2 Then 'Remember, if they aren't in a party they'll only get one person, so this has to be at least two
                Exp = Exp * 0.75 ' 3/4 experience
                'MsgBox Exp
                For d = 1 To MAX_PARTY_MEMBERS
                    If Player(Attacker).Party.Member(d) > 0 Then
                        If Player(Player(Attacker).Party.Member(d)).Party.ShareExp = True Then
                            If GetPlayerLevel(Player(Attacker).Party.Member(d)) = MAX_LEVEL Then
                                Call SetPlayerExp(Player(Attacker).Party.Member(d), Experience(MAX_LEVEL))
                                Call BattleMsg(Player(Attacker).Party.Member(d), "You cant gain anymore experience!", BrightBlue, 0)
                            Else
                                Call SetPlayerExp(Player(Attacker).Party.Member(d), GetPlayerExp(Player(Attacker).Party.Member(d)) + Exp)
                                Call BattleMsg(Player(Attacker).Party.Member(d), "You have gained " & Exp & " party experience.", BrightBlue, 0)
                            End If
                        End If
                    End If
                Next d
            Else 'if there are 3 or more party members..
                Exp = Exp * 0.5  ' 1/2 experience
                    For d = 1 To MAX_PARTY_MEMBERS
                        If Player(Attacker).Party.Member(d) > 0 Then
                            If Player(Player(Attacker).Party.Member(d)).Party.ShareExp = True Then
                                If GetPlayerLevel(Player(Attacker).Party.Member(d)) = MAX_LEVEL Then
                                    Call SetPlayerExp(Player(Attacker).Party.Member(d), Experience(MAX_LEVEL))
                                    Call BattleMsg(Player(Attacker).Party.Member(d), "You cant gain anymore experience!", BrightBlue, 0)
                                Else
                                    Call SetPlayerExp(Player(Attacker).Party.Member(d), GetPlayerExp(n) + Exp)
                                    Call BattleMsg(Player(Attacker).Party.Member(d), "You have gained " & Exp & " party experience.", BrightBlue, 0)
                                End If
                            End If
                        End If
                    Next d
            End If
        End If
                                
        For i = 1 To MAX_NPC_DROPS
            ' Drop the goods if they get it
            n = Int(Rnd * Npc(NpcNum).ItemNPC(i).Chance) + 1
            If n = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapAttributeNpc(MapNum, MapNpcNum, x, y).x, MapAttributeNpc(MapNum, MapNpcNum, x, y).y)
            End If
        Next i
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        Call ClearMapAttributeNpc(MapNpcNum, x, y, MapNum)
        MapAttributeNpc(MapNum, MapNpcNum, x, y).num = 0
        MapAttributeNpc(MapNum, MapNpcNum, x, y).SpawnWait = GetTickCount
        MapAttributeNpc(MapNum, MapNpcNum, x, y).HP = 0
        Call SendDataToMap(MapNum, "ATTRIBUTENPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
       
        ' Check for level up party member
        If Player(Attacker).InParty = YES Then
            For d = 1 To MAX_PARTY_MEMBERS
                Call CheckPlayerLevelUp(Player(Attacker).Party.Member(d))
            Next d
        End If
        
        ' Check for level up party member
        If Player(Attacker).InParty = YES Then
            Call CheckPlayerLevelUp(Player(Attacker).PartyPlayer)
        End If
    
        ' Check if target is npc that died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_NPC And Player(Attacker).Target = MapNpcNum Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
    Else
        ' NPC not dead, just do the damage
        MapAttributeNpc(MapNum, MapNpcNum, x, y).HP = MapAttributeNpc(MapNum, MapNpcNum, x, y).HP - Damage
        
        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "You hit a " & Name & " for " & Damage & " damage.", White, 0)
        
        If n = 0 Then
            'Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        Else
            'Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If
        
        ' Check if we should send a message
        If MapAttributeNpc(MapNum, MapNpcNum, x, y).Target = 0 And MapAttributeNpc(MapNum, MapNpcNum, x, y).Target <> Attacker Then
            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " : " & Trim(Npc(NpcNum).AttackSay) & "", SayColor)
            End If
        End If
        
        ' Set the NPC target to the player
        MapAttributeNpc(MapNum, MapNpcNum, x, y).Target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapAttributeNpc(MapNum, MapNpcNum, x, y).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_ATTRIBUTE_NPCS
                If MapNpc(MapNum, i).num = MapAttributeNpc(MapNum, MapNpcNum, x, y).num Then
                    MapNpc(MapNum, i).Target = Attacker
                End If
            Next i
        End If
    End If
        
    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapAttributeNpc(MapNum, MapNpcNum, x, y).HP & SEP_CHAR & GetNpcMaxHP(MapAttributeNpc(MapNum, MapNpcNum, x, y).num) & SEP_CHAR & END_CHAR)
        
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Function CanNPCMoveAttributeNPC(ByVal MapNum As Long, ByVal index As Long, ByVal Dir As Long) As Boolean
Dim i As Long, BX As Long, BY As Long, x As Long, y As Long

CanNPCMoveAttributeNPC = True

For BX = 0 To MAX_MAPX
    For BY = 0 To MAX_MAPY
        If Map(MapNum).Tile(BX, BY).Type = TILE_TYPE_NPC_SPAWN Then
            For i = 1 To MAX_ATTRIBUTE_NPCS
                If MapAttributeNpc(MapNum, i, BX, BY).num > 0 Then
                    Select Case Dir
                        Case DIR_UP
                            If (MapNpc(MapNum, index).x = MapAttributeNpc(MapNum, i, BX, BY).x) And (MapNpc(MapNum, index).y - 1 = MapAttributeNpc(MapNum, i, BX, BY).y) Then
                                CanNPCMoveAttributeNPC = False
                                Exit Function
                            End If
                        Case DIR_DOWN
                            If (MapNpc(MapNum, index).x = MapAttributeNpc(MapNum, i, BX, BY).x) And (MapNpc(MapNum, index).y + 1 = MapAttributeNpc(MapNum, i, BX, BY).y) Then
                                CanNPCMoveAttributeNPC = False
                                Exit Function
                            End If
                        Case DIR_LEFT
                            If (MapNpc(MapNum, index).x - 1 = MapAttributeNpc(MapNum, i, BX, BY).x) And (MapNpc(MapNum, index).y = MapAttributeNpc(MapNum, i, BX, BY).y) Then
                                CanNPCMoveAttributeNPC = False
                                Exit Function
                            End If
                        Case DIR_RIGHT
                            If (MapNpc(MapNum, index).x + 1 = MapAttributeNpc(MapNum, i, BX, BY).x) And (MapNpc(MapNum, index).y = MapAttributeNpc(MapNum, i, BX, BY).y) Then
                                CanNPCMoveAttributeNPC = False
                                Exit Function
                            End If
                    End Select
                End If
            Next i
        End If
    Next BY
Next BX
End Function
