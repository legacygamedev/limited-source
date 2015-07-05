Attribute VB_Name = "modGameLogic"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Function FindOpenPlayerSlot() As Long
Dim i As Long

    FindOpenPlayerSlot = 0
    
    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
Dim i As Long

    FindOpenMapItemSlot = 0
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If
    
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next
End Function

Function TotalOnlinePlayers() As Long
Dim i As Long

    TotalOnlinePlayers = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next
End Function

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, x, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    i = MapItemSlot
    
    If i <> 0 Then
        If ItemNum >= 0 Then
            If ItemNum <= MAX_ITEMS Then
    
                MapItem(MapNum, i).Num = ItemNum
                MapItem(MapNum, i).Value = ItemVal
                
                If ItemNum <> 0 Then
                    If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                        MapItem(MapNum, i).Dur = ItemDur
                    Else
                        MapItem(MapNum, i).Dur = 0
                    End If
                Else
                    MapItem(MapNum, i).Dur = 0
                End If
                
                MapItem(MapNum, i).x = x
                MapItem(MapNum, i).y = y
                
                Set Buffer = New clsBuffer
                
                Buffer.WriteLong SSpawnItem
                Buffer.WriteLong i
                Buffer.WriteLong ItemNum
                Buffer.WriteLong ItemVal
                Buffer.WriteLong MapItem(MapNum, i).Dur
                Buffer.WriteLong x
                Buffer.WriteLong y
                    
                SendDataToMap MapNum, Buffer.ToArray()
                
                Set Buffer = Nothing
                
            End If
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
Dim x As Long
Dim y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Spawn what we have
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If
        Next
    Next
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Packet As String
Dim NpcNum As Long
Dim i As Long
Dim x As Long
Dim y As Long
Dim Spawned As Boolean
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    NpcNum = Map(MapNum).Npc(MapNpcNum)
    If NpcNum > 0 Then
        MapNpc(MapNum).Npc(MapNpcNum).Num = NpcNum
        MapNpc(MapNum).Npc(MapNpcNum).Target = 0
        
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(NpcNum, Vitals.MP)
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.SP) = GetNpcMaxVital(NpcNum, Vitals.SP)
                
        MapNpc(MapNum).Npc(MapNpcNum).Dir = Int(Rnd * 4)
        
        ' Well try 100 times to randomly place the sprite
        For i = 1 To 100
            x = Int(Rnd * Map(MapNum).MaxX)
            y = Int(Rnd * Map(MapNum).MaxY)
            
            ' Check if the tile is walkable
            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                MapNpc(MapNum).Npc(MapNpcNum).x = x
                MapNpc(MapNum).Npc(MapNpcNum).y = y
                Spawned = True
                Exit For
            End If
        Next
        
        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For x = 0 To Map(MapNum).MaxX
                For y = 0 To Map(MapNum).MaxY
                    If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                        MapNpc(MapNum).Npc(MapNpcNum).x = x
                        MapNpc(MapNum).Npc(MapNpcNum).y = y
                        Spawned = True
                    End If
                Next
            Next
        End If
             
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
        
            Set Buffer = New clsBuffer
            
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Num
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            
            SendDataToMap MapNum, Buffer.ToArray()
            
            Set Buffer = Nothing
        End If
    End If
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next
End Sub

Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    ' Check attack timer
    If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
    
    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function
       
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function
   
    ' Check if at same coordinates
    Select Case GetPlayerDir(Attacker)
        Case DIR_UP
            If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
        Case DIR_DOWN
            If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
        Case DIR_LEFT
            If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
        Case DIR_RIGHT
            If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
        Case Else
            Exit Function
    End Select
    
    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If
   
    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function
    
    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If
   
    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
    
    CanAttackPlayer = True

End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long
Dim NpcNum As Long
Dim NpcX As Long
Dim NpcY As Long
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + 1000 Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x + 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x - 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y
            End Select
            
            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, "You cannot attack a " & Trim$(Npc(NpcNum).Name) & "!", BrightBlue)
                    End If
                End If
            End If
        End If
    End If
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long
Dim NpcNum As Long
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Index)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum).Npc(MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If
    
    MapNpc(MapNum).Npc(MapNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NpcNum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If

'            Select Case mapnpc(mapnum).Npc(MapNpcNum).Dir
'                Case DIR_UP
'                    If (GetPlayerY(Index) + 1 = mapnpc(mapnum).Npc(MapNpcNum).y) And (GetPlayerX(Index) = mapnpc(mapnum).Npc(MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_DOWN
'                    If (GetPlayerY(Index) - 1 = mapnpc(mapnum).Npc(MapNpcNum).y) And (GetPlayerX(Index) = mapnpc(mapnum).Npc(MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_LEFT
'                    If (GetPlayerY(Index) = mapnpc(mapnum).Npc(MapNpcNum).y) And (GetPlayerX(Index) + 1 = mapnpc(mapnum).Npc(MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_RIGHT
'                    If (GetPlayerY(Index) = mapnpc(mapnum).Npc(MapNpcNum).y) And (GetPlayerX(Index) - 1 = mapnpc(mapnum).Npc(MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'            End Select

        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim MapNum As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Name)
    
    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong MapNpcNum
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' reduce dur. on victims equipment
    Call DamageEquipment(Victim, Armor)
    Call DamageEquipment(Victim, Helmet)
        
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by a " & Name, BrightRed)
                
        ' Calculate exp to give attacker
        Exp = GetPlayerExp(Victim) \ 3
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then Exp = 0
        
        If Exp = 0 Then
            Call PlayerMsg(Victim, "You lost no experience points.", BrightRed)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            Call PlayerMsg(Victim, "You lost " & Exp & " experience points.", BrightRed)
        End If
        
        ' Set NPC target to 0
        MapNpc(MapNum).Npc(MapNpcNum).Target = 0
        
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' Say damage
        Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
    End If
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
Dim i As Long
Dim n As Long
Dim x As Long
Dim y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    
    x = MapNpc(MapNum).Npc(MapNpcNum).x
    y = MapNpc(MapNum).Npc(MapNpcNum).y
    
    CanNpcMove = True
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(MapNum).Tile(x, y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(x, y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
            Else
                CanNpcMove = False
            End If
                
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(MapNum).Tile(x - 1, y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
            Else
                CanNpcMove = False
            End If
                
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If x < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    MapNpc(MapNum).Npc(MapNpcNum).Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1
            
            Set Buffer = New clsBuffer
            
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            Buffer.WriteLong Movement
            
            SendDataToMap MapNum, Buffer.ToArray()
            
            Set Buffer = Nothing
    
        Case DIR_DOWN
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1
            
            Set Buffer = New clsBuffer
            
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            Buffer.WriteLong Movement
            
            SendDataToMap MapNum, Buffer.ToArray()
            
            Set Buffer = Nothing
    
        Case DIR_LEFT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
            
            Set Buffer = New clsBuffer
            
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            Buffer.WriteLong Movement
            
            SendDataToMap MapNum, Buffer.ToArray()
            
            Set Buffer = Nothing
    
        Case DIR_RIGHT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
            
            Set Buffer = New clsBuffer
            
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            Buffer.WriteLong Movement
            
            SendDataToMap MapNum, Buffer.ToArray()
            
            Set Buffer = Nothing
            
    End Select
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    MapNpc(MapNum).Npc(MapNpcNum).Dir = Dir
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong MapNpcNum
    Buffer.WriteLong Dir
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
Dim i As Long
Dim n As Long

    n = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If
    Next
    
    GetTotalMapPlayers = n
End Function

Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
Dim x As Long
Dim y As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If
    
    Select Case Vital
        Case HP
            x = Npc(NpcNum).Stat(Stats.Strength)
            y = Npc(NpcNum).Stat(Stats.Defense)
            GetNpcMaxVital = x * y
        Case MP
            GetNpcMaxVital = Npc(NpcNum).Stat(Stats.Magic) * 2
        Case SP
            GetNpcMaxVital = Npc(NpcNum).Stat(Stats.Speed) * 2
    End Select
End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If
    
    Select Case Vital
        Case HP
            i = Npc(NpcNum).Stat(Stats.Defense) \ 3
            If i < 1 Then i = 1
                GetNpcVitalRegen = i
        'Case MP
        
        'Case SP
    
    End Select
End Function

Sub ClearTempTiles()
Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next
End Sub
Sub ClearTempTile(ByVal MapNum As Long)
Dim y As Long
Dim x As Long

    TempTile(MapNum).DoorTimer = 0
        
    ReDim TempTile(MapNum).DoorOpen(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
        
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            TempTile(MapNum).DoorOpen(x, y) = NO
        Next
    Next
End Sub

Public Sub UpdateHighIndex()
Dim i As Integer
Dim array_index As Integer
Dim Buffer As clsBuffer
    
    ' no players are logged in, allow one connection
    If TotalPlayersOnline < 1 Then
        High_Index = 1
        Exit Sub
    End If

    ' new size
    ReDim PlayersOnline(1 To TotalPlayersOnline)

    For i = 1 To MAX_PLAYERS
        If LenB((GetPlayerLogin(i))) > 0 Then
            High_Index = i
            array_index = array_index + 1
            PlayersOnline(array_index) = i
            
            ' early finish if all players are found
            If array_index >= TotalPlayersOnline Then
                Exit For
            End If
            
        End If
    Next
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong High_Index
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

