Attribute VB_Name = "modGameLogic"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Function FindOpenPlayerSlot() As Long
Dim i As Long

    FindOpenPlayerSlot = 0
    
    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next i
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
    Next i
End Function

Function TotalOnlinePlayers() As Long
Dim i As Long

    TotalOnlinePlayers = 0
    
    For i = 1 To HighIndex
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next i
End Function

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    For i = 1 To HighIndex
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim(Name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim(Name)))) = UCase(Trim(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next i
End Function

Function PushBlockBlocked(ByVal MapNum As Long, ByVal x As Integer, ByVal y As Integer) As Boolean
Dim i As Long
    
    PushBlockBlocked = False

    For i = 1 To MAX_PLAYERS
        If Player(i).Char(Player(i).CharNum).Map = MapNum Then
            If Player(i).Char(Player(i).CharNum).x = x And Player(i).Char(Player(i).CharNum).y = y Then
                PushBlockBlocked = True
            End If
        End If
    Next i
End Function

Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, i)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If i = GetPlayerWeaponSlot(Index) Then
                                Call SetPlayerWeaponSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                
                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(Index) > 0 Then
                            If i = GetPlayerArmorSlot(Index) Then
                                Call SetPlayerArmorSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(Index) > 0 Then
                            If i = GetPlayerHelmetSlot(Index) Then
                                Call SetPlayerHelmetSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(Index) > 0 Then
                            If i = GetPlayerShieldSlot(Index) Then
                                Call SetPlayerShieldSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                    Case ITEM_TYPE_TOOL
                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If i = GetPlayerWeaponSlot(Index) Then
                                Call SetPlayerWeaponSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                    Case ITEM_TYPE_AMULET
                        If GetPlayerAmuletSlot(Index) > 0 Then
                            If i = GetPlayerAmuletSlot(Index) Then
                                Call SetPlayerAmuletSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                    Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(Index) > 0 Then
                            If i = GetPlayerRingSlot(Index) Then
                                Call SetPlayerRingSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                    Case ITEM_TYPE_ARROW
                        If GetPlayerArrowSlot(Index) > 0 Then
                            If i = GetPlayerArrowSlot(Index) Then
                                Call SetPlayerArrowSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select

                
                n = Item(GetPlayerInvItemNum(Index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_TOOL) And (n <> ITEM_TYPE_AMULET) And (n <> ITEM_TYPE_RING) And (n <> ITEM_TYPE_ARROW) Then
                    TakeItem = True
                End If
            End If
                            
            If TakeItem = True Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Call SetPlayerInvItemDur(Index, i, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = FindOpenInvSlot(Index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_TOOL) Or (Item(ItemNum).Type = ITEM_TYPE_ARROW) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(Index, i)
    Else
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Inventory Full" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
    End If
End Sub

Function GiveQuestReward(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Byte
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    GiveQuestReward = 0
    
    i = FindOpenInvSlot(Index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_TOOL) Or (Item(ItemNum).Type = ITEM_TYPE_ARROW) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
        End If
        
        GiveQuestReward = 1
        Call SendInventoryUpdate(Index, i)
    Else
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Inventory Full" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
    End If
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, x, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim Packet As String
Dim i As Long
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    i = MapItemSlot
    
    If i <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
        MapItem(MapNum, i).Num = ItemNum
        MapItem(MapNum, i).Value = ItemVal
        
        If ItemNum <> 0 Then
            If ((Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD)) Or (Item(ItemNum).Type = ITEM_TYPE_TOOL) Or (Item(ItemNum).Type = ITEM_TYPE_ARROW) Then
                MapItem(MapNum, i).Dur = ItemDur
            Else
                MapItem(MapNum, i).Dur = 0
            End If
        Else
            MapItem(MapNum, i).Dur = 0
        End If
        
        MapItem(MapNum, i).x = x
        MapItem(MapNum, i).y = y
            
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
End Sub

Sub SpawnAllMapsItems()
Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next i
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
Dim x As Long
Dim y As Long
Dim i As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Spawn what we have
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If
        Next x
    Next y
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Packet As String
Dim NpcNum As Long
Dim i As Long, x As Long, y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    Spawned = False
    
    NpcNum = Map(MapNum).Npc(MapNpcNum)
    If NpcNum > 0 Then
        MapNpc(MapNum, MapNpcNum).Num = NpcNum
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxHP(NpcNum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)
                
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)
        
        ' Check for npc spawn location
        If (Map(MapNum).NSpawn(MapNpcNum).NSx > 0) Or (Map(MapNum).NSpawn(MapNpcNum).NSy > 0) Then
            x = Map(MapNum).NSpawn(MapNpcNum).NSx
            y = Map(MapNum).NSpawn(MapNpcNum).NSy
            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NSPAWN Then
                MapNpc(MapNum, MapNpcNum).x = x
                MapNpc(MapNum, MapNpcNum).y = y
                Spawned = True
                If Spawned Then
                    Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & END_CHAR
                    Call SendDataToMap(MapNum, Packet)
                    Exit Sub
                End If
            End If
        End If
        
        ' Well try 100 times to randomly place the sprite
        For i = 1 To 100
            x = Int(Rnd * MAX_MAPX)
            y = Int(Rnd * MAX_MAPY)
            
            ' Check if the tile is walkable
            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                MapNpc(MapNum, MapNpcNum).x = x
                MapNpc(MapNum, MapNpcNum).y = y
                Spawned = True
                Exit For
            End If
        Next i
        
        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                        MapNpc(MapNum, MapNpcNum).x = x
                        MapNpc(MapNum, MapNpcNum).y = y
                        Spawned = True
                    End If
                Next x
            Next y
        End If
             
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        End If
    End If
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next i
End Sub

Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next i
End Sub

Sub SpawnResource(ByVal MapResourceNum As Long, ByVal MapNum As Long)
Dim Packet As String
Dim ResourceNum As Long
Dim i As Long, x As Long, y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapResourceNum <= 0 Or MapResourceNum > MAX_MAP_RESOURCES Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    Spawned = False
    
    ResourceNum = Map(MapNum).Resource(MapResourceNum)
    If ResourceNum > 0 Then
        MapResource(MapNum, MapResourceNum).Num = ResourceNum
        
        MapResource(MapNum, MapResourceNum).HP = GetNpcMaxHP(ResourceNum)
        
        ' Check for resource spawn location
        If (Map(MapNum).RSpawn(MapResourceNum).RSx) Or (Map(MapNum).RSpawn(MapResourceNum).RSy) > 0 Then
            x = Map(MapNum).RSpawn(MapResourceNum).RSx
            y = Map(MapNum).RSpawn(MapResourceNum).RSy
            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_RSPAWN Then
                MapNpc(MapNum, MapResourceNum).x = x
                MapNpc(MapNum, MapResourceNum).y = y
                Spawned = True
                If Spawned Then
                    Packet = "SPAWNRESOURCE" & SEP_CHAR & MapResourceNum & SEP_CHAR & MapResource(MapNum, MapResourceNum).Num & SEP_CHAR & MapResource(MapNum, MapResourceNum).x & SEP_CHAR & MapResource(MapNum, MapResourceNum).y & SEP_CHAR & END_CHAR
                    Call SendDataToMap(MapNum, Packet)
                    Exit Sub
                End If
            End If
        End If
        
        ' Well try 100 times to randomly place the sprite
        For i = 1 To 100
            x = Int(Rnd * MAX_MAPX)
            y = Int(Rnd * MAX_MAPY)
            
            ' Check if the tile is walkable
            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                MapResource(MapNum, MapResourceNum).x = x
                MapResource(MapNum, MapResourceNum).y = y
                Spawned = True
                Exit For
            End If
        Next i
        
        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    If Map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                        MapResource(MapNum, MapResourceNum).x = x
                        MapResource(MapNum, MapResourceNum).y = y
                        Spawned = True
                    End If
                Next x
            Next y
        End If
             
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = "SPAWNRESOURCE" & SEP_CHAR & MapResourceNum & SEP_CHAR & MapResource(MapNum, MapResourceNum).Num & SEP_CHAR & MapResource(MapNum, MapResourceNum).x & SEP_CHAR & MapResource(MapNum, MapResourceNum).y & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        End If
    End If
End Sub

Sub SpawnMapResources(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To MAX_MAP_RESOURCES
        Call SpawnResource(i, MapNum)
    Next i
End Sub

Sub SpawnAllMapResources()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapResources(i)
    Next i
End Sub

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
    
    CanNpcAttackPlayer = False
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Index) = False Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index), MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    ' Make sure the Npc type can attack
    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_RESOURCE Then
        Exit Function
    End If
    
    ' Check if player is an admin
    If GetPlayerAccess(Index) >= ADMIN_MAPPER Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Index)
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If Player(Index).GettingMap = YES Then
        Exit Function
    End If
    
    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NpcNum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum, MapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If

'            Select Case MapNpc(MapNum, MapNpcNum).Dir
'                Case DIR_UP
'                    If (GetPlayerY(Index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_DOWN
'                    If (GetPlayerY(Index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_LEFT
'                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'
'                Case DIR_RIGHT
'                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
'                        CanNpcAttackPlayer = True
'                    End If
'            End Select
        End If
    End If
End Function

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim EXP As Long
Dim n As Long
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
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
        
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call SendDataTo(Attacker, "BLITWARNMSG" & SEP_CHAR & "Battle WON" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
            Call SendDataTo(Victim, "BLITWARNMSG" & SEP_CHAR & "Battle LOST" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
        Else
            Call SendDataTo(Attacker, "BLITWARNMSG" & SEP_CHAR & "Battle Won" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
            Call SendDataTo(Victim, "BLITWARNMSG" & SEP_CHAR & "Battle LOST" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
        End If
        
        '' Player is dead
        'Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        
        ' Drop all worn items by victim
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
        If GetPlayerAmuletSlot(Victim) > 0 Then
            Call PlayerMapDropItem(Victim, GetPlayerAmuletSlot(Victim), 0)
        End If
        If GetPlayerRingSlot(Victim) > 0 Then
            Call PlayerMapDropItem(Victim, GetPlayerRingSlot(Victim), 0)
        End If
        If GetPlayerArrowSlot(Victim) > 0 Then
            Call PlayerMapDropItem(Victim, GetPlayerArrowSlot(Victim), 0)
        End If

        ' Calculate exp to give attacker
        EXP = Int(GetPlayerExp(Victim) / 10)
        
        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 0
        End If
        
        If EXP = 0 Then
            Call SendDataTo(Victim, "BLITPLAYERMSG" & SEP_CHAR & "No EXP Lost" & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
            Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "No EXP Gained" & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - EXP)
            Call SendDataTo(Victim, "BLITPLAYERMSG" & SEP_CHAR & EXP & " EXP Lost" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
            Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & EXP & " EXP Gained" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        End If
                
        ' Warp player away
        Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
                
        ' Check for a level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check if target is player who died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
        
        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                'Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If
        Else
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
            'Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Say damage
        If n = 0 Then
            Call SendDataTo(Attacker, "BLITPKDMG" & SEP_CHAR & Damage & SEP_CHAR & Victim & SEP_CHAR & White & SEP_CHAR & END_CHAR)
            Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        Else
            Call SendDataTo(Attacker, "BLITPKDMG" & SEP_CHAR & Damage & SEP_CHAR & Victim & SEP_CHAR & White & SEP_CHAR & END_CHAR)
            Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        End If
    End If
    
    ' Reset timer for attacking
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Name As String
Dim EXP As Long
Dim MapNum As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim), MapNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Victim)
    Name = Trim(Npc(MapNpc(MapNum, MapNpcNum).Num).Name)
    
    If Damage >= GetPlayerHP(Victim) Then
        ' Say damage
        Call SendDataTo(Victim, "BLITNPCMSG" & SEP_CHAR & "Death Blow" & SEP_CHAR & MapNpcNum & SEP_CHAR & Brown & SEP_CHAR & END_CHAR)
        
        ' Player is dead
        'Call GlobalMsg(GetPlayerName(Victim) & " has been killed by a " & Name, BrightRed)
        
        ' Drop all worn items by victim
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
        If GetPlayerAmuletSlot(Victim) > 0 Then
            Call PlayerMapDropItem(Victim, GetPlayerAmuletSlot(Victim), 0)
        End If
        If GetPlayerRingSlot(Victim) > 0 Then
            Call PlayerMapDropItem(Victim, GetPlayerRingSlot(Victim), 0)
        End If
        If GetPlayerArrowSlot(Victim) > 0 Then
            Call PlayerMapDropItem(Victim, GetPlayerArrowSlot(Victim), 0)
        End If
        
        ' Calculate exp to give attacker
        EXP = Int(GetPlayerExp(Victim) / 3)
        
        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 0
        End If
        
        If EXP = 0 Then
            Call SendDataTo(Victim, "BLITPLAYERMSG" & SEP_CHAR & "No EXP Lost" & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - EXP)
            Call SendDataTo(Victim, "BLITPLAYERMSG" & SEP_CHAR & EXP & " EXP Lost" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        End If
                
        ' Warp player away
        Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
        
        ' Set NPC target to 0
        MapNpc(MapNum, MapNpcNum).Target = 0
        
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
        Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
    End If
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
Dim i As Long, n As Long
Dim x As Long, y As Long

    CanNpcMove = False
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    
    x = MapNpc(MapNum, MapNpcNum).x
    y = MapNpc(MapNum, MapNpcNum).y
    
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
                For i = 1 To HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
                
                ' Check to make sure that there is not a resource in the way
                For i = 1 To MAX_MAP_RESOURCES
                    If (MapResource(MapNum, i).Num > 0) And (MapResource(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x) And (MapResource(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If y < MAX_MAPY Then
                n = Map(MapNum).Tile(x, y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
                
                ' Check to make sure that there is not a resource in the way
                For i = 1 To MAX_MAP_RESOURCES
                    If (MapResource(MapNum, i).Num > 0) And (MapResource(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x) And (MapResource(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
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
                For i = 1 To HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x - 1) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
                
                ' Check to make sure that there is not a resource in the way
                For i = 1 To MAX_MAP_RESOURCES
                    If (MapResource(MapNum, i).Num > 0) And (MapResource(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x - 1) And (MapResource(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
            Else
                CanNpcMove = False
            End If
                
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If x < MAX_MAPX Then
                n = Map(MapNum).Tile(x + 1, y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x + 1) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
                
                ' Check to make sure that there is not a resource in the way
                For i = 1 To MAX_MAP_RESOURCES
                    If (MapResource(MapNum, i).Num > 0) And (MapResource(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x + 1) And (MapResource(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim x As Long
Dim y As Long
Dim i As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum, MapNpcNum).y = MapNpc(MapNum, MapNpcNum).y - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_DOWN
            MapNpc(MapNum, MapNpcNum).y = MapNpc(MapNum, MapNpcNum).y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_LEFT
            MapNpc(MapNum, MapNpcNum).x = MapNpc(MapNum, MapNpcNum).x - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).x = MapNpc(MapNum, MapNpcNum).x + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    End Select
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
Dim Packet As String

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    Packet = "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub JoinGame(ByVal Index As Long)
    ' Set the flag so we know the person is in the game
    Player(Index).InGame = True
        
    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) <= ADMIN_MONITER Then
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", White)
    End If
        
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(Index, "LOGINOK" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendNpcs(Index)
    
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendSkills(Index)
    Call SendQuests(Index)
    Call SendGUIS(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendStats(Index)
    Call SendPlayerSpells(Index)
    Call SendPlayerSkills(Index)
    Call SendPlayerQuests(Index)
    'Call SendWeatherTo(Index)
    'Call SendTimeTo(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            
    ' Send welcome messages
    'Call SendWelcome(Index)

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(Index, "INGAME" & SEP_CHAR & END_CHAR)
End Sub

Sub LeftGame(ByVal Index As Long)
Dim n As Long

    If Player(Index).InGame = True Then
        Player(Index).InGame = False
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) = 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
        ' Check for boot map
        If Map(GetPlayerMap(Index)).BootMap > 0 Then
            Call SetPlayerX(Index, Map(GetPlayerMap(Index)).BootX)
            Call SetPlayerY(Index, Map(GetPlayerMap(Index)).BootY)
            Call SetPlayerMap(Index, Map(GetPlayerMap(Index)).BootMap)
        End If
        
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If Player(Index).InParty = YES Then
            n = Player(Index).PartyPlayer
            
            Call PlayerMsg(n, GetPlayerName(Index) & " has left " & GAME_NAME & ", disbanning party.", Pink)
            Player(n).InParty = NO
            Player(n).PartyPlayer = 0
        End If
            
        Call SavePlayer(Index)
    
        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITER Then
            Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", White)
        End If
        Call TextAdd(frmCServer.txtText, GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".", True)
        Call AddLog(GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".", PLAYER_LOG)
        Call SendLeftGame(Index)
    End If
    
    Call ClearPlayer(Index)
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
Dim i As Long, n As Long

    n = 0
    
    For i = 1 To HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If
    Next i
    
    GetTotalMapPlayers = n
End Function

Function GetNpcMaxHP(ByVal NpcNum As Long)

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxHP = 0
        Exit Function
    End If
    
    GetNpcMaxHP = Npc(NpcNum).MaxHp
End Function

Function GetNpcMaxMP(ByVal NpcNum As Long)
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxMP = 0
        Exit Function
    End If
        
    GetNpcMaxMP = Npc(NpcNum).MAGI * 2
End Function

Function GetNpcMaxSP(ByVal NpcNum As Long)
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxSP = 0
        Exit Function
    End If
        
    GetNpcMaxSP = Npc(NpcNum).SPEED * 2
End Function

Function GetSpellReqLevel(ByVal Index As Long, ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).LevelReq
End Function

Function GetSkillReqLevel(ByVal Index As Long, ByVal SkillNum As Long)
    GetSkillReqLevel = Skill(SkillNum).LevelReq
End Function

Sub CheckEquippedItems(ByVal Index As Long)
Dim Slot As Long, ItemNum As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    Slot = GetPlayerWeaponSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
                If Item(ItemNum).Type <> ITEM_TYPE_TOOL Then
                    Call SetPlayerWeaponSlot(Index, 0)
                End If
            End If
        Else
            Call SetPlayerWeaponSlot(Index, 0)
        End If
    End If

    Slot = GetPlayerArmorSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then
                Call SetPlayerArmorSlot(Index, 0)
            End If
        Else
            Call SetPlayerArmorSlot(Index, 0)
        End If
    End If

    Slot = GetPlayerHelmetSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then
                Call SetPlayerHelmetSlot(Index, 0)
            End If
        Else
            Call SetPlayerHelmetSlot(Index, 0)
        End If
    End If

    Slot = GetPlayerShieldSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then
                Call SetPlayerShieldSlot(Index, 0)
            End If
        Else
            Call SetPlayerShieldSlot(Index, 0)
        End If
    End If
    
    Slot = GetPlayerAmuletSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_AMULET Then
                Call SetPlayerAmuletSlot(Index, 0)
            End If
        Else
            Call SetPlayerAmuletSlot(Index, 0)
        End If
    End If
    
    Slot = GetPlayerRingSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_RING Then
                Call SetPlayerRingSlot(Index, 0)
            End If
        Else
            Call SetPlayerRingSlot(Index, 0)
        End If
    End If
    
    Slot = GetPlayerArrowSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_ARROW Then
                Call SetPlayerArrowSlot(Index, 0)
            End If
        Else
            Call SetPlayerArrowSlot(Index, 0)
        End If
    End If
End Sub

Sub ClearTempTile()
Dim i As Long, y As Long, x As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                TempTile(i).DoorOpen(x, y) = NO
            Next x
        Next y
    Next i
End Sub

Sub ClearClasses()
Dim i As Long

    For i = 0 To Max_Classes
        Class(i).Name = ""
        Class(i).STR = 0
        Class(i).DEF = 0
        Class(i).SPEED = 0
        Class(i).MAGI = 0
        Class(i).DEX = 0
    Next i
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Player(Index).Login = ""
    Player(Index).Password = ""
    
    For i = 1 To MAX_CHARS
        Player(Index).Char(i).Name = ""
        Player(Index).Char(i).Sex = 0
        Player(Index).Char(i).Class = 0
        Player(Index).Char(i).Sprite = 0
        Player(Index).Char(i).Level = 0
        Player(Index).Char(i).Access = 0
        Player(Index).Char(i).EXP = 0
        Player(Index).Char(i).PK = NO
        Player(Index).Char(i).Guild = 0
        
        Player(Index).Char(i).HP = 0
        Player(Index).Char(i).MP = 0
        Player(Index).Char(i).SP = 0
        
        Player(Index).Char(i).STR = 0
        Player(Index).Char(i).DEF = 0
        Player(Index).Char(i).SPEED = 0
        Player(Index).Char(i).MAGI = 0
        Player(Index).Char(i).DEX = 0
        Player(Index).Char(i).POINTS = 0
        
        Player(Index).Char(i).WeaponSlot = 0
        Player(Index).Char(i).ArmorSlot = 0
        Player(Index).Char(i).HelmetSlot = 0
        Player(Index).Char(i).ShieldSlot = 0
        Player(Index).Char(i).AmuletSlot = 0
        Player(Index).Char(i).RingSlot = 0
        Player(Index).Char(i).ArrowSlot = 0
        
        Player(Index).Char(i).Map = 0
        Player(Index).Char(i).x = 0
        Player(Index).Char(i).y = 0
        Player(Index).Char(i).Dir = 0
        
        For n = 1 To MAX_INV
            Player(Index).Char(i).Inv(n).Num = 0
            Player(Index).Char(i).Inv(n).Value = 0
            Player(Index).Char(i).Inv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SKILLS
            Player(Index).Char(i).Skills(n).Num = 0
            Player(Index).Char(i).Skills(n).Level = 0
            Player(Index).Char(i).Skills(n).EXP = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(i).Spells(n).Num = 0
            Player(Index).Char(i).Spells(n).Level = 0
            Player(Index).Char(i).Spells(n).EXP = 0
        Next n
        
        For n = 1 To MAX_PLAYER_QUESTS
            Player(Index).Char(i).Quests(n).Num = 0
        Next n
        
        For n = 1 To MAX_PLAYER_MAPS
            Player(Index).Char(i).Maps(n).Num = 0
        Next n
    Next i
        
        ' Temporary vars
        Player(Index).Buffer = ""
        'Player(Index).IncBuffer = ""
        Player(Index).CharNum = 0
        Player(Index).InGame = False
        Player(Index).AttackTimer = 0
        Player(Index).DataTimer = 0
        Player(Index).DataBytes = 0
        Player(Index).DataPackets = 0
        Player(Index).PartyPlayer = 0
        Player(Index).InParty = 0
        Player(Index).Target = 0
        Player(Index).TargetType = 0
        Player(Index).CastedSpell = NO
        Player(Index).PartyStarter = NO
        Player(Index).GettingMap = NO
End Sub

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
Dim n As Long
    
    Player(Index).Char(CharNum).Name = ""
    Player(Index).Char(CharNum).Sex = 0
    Player(Index).Char(CharNum).Class = 0
    Player(Index).Char(CharNum).Sprite = 0
    Player(Index).Char(CharNum).Level = 0
    Player(Index).Char(CharNum).Access = 0
    Player(Index).Char(CharNum).EXP = 0
    Player(Index).Char(CharNum).PK = NO
    Player(Index).Char(CharNum).Guild = 0
    
    Player(Index).Char(CharNum).HP = 0
    Player(Index).Char(CharNum).MP = 0
    Player(Index).Char(CharNum).SP = 0
    
    Player(Index).Char(CharNum).STR = 0
    Player(Index).Char(CharNum).DEF = 0
    Player(Index).Char(CharNum).SPEED = 0
    Player(Index).Char(CharNum).MAGI = 0
    Player(Index).Char(CharNum).DEX = 0
    Player(Index).Char(CharNum).POINTS = 0
    
    Player(Index).Char(CharNum).ArmorSlot = 0
    Player(Index).Char(CharNum).WeaponSlot = 0
    Player(Index).Char(CharNum).HelmetSlot = 0
    Player(Index).Char(CharNum).ShieldSlot = 0
    Player(Index).Char(CharNum).AmuletSlot = 0
    Player(Index).Char(CharNum).RingSlot = 0
    Player(Index).Char(CharNum).ArrowSlot = 0
    
    Player(Index).Char(CharNum).Map = 0
    Player(Index).Char(CharNum).x = 0
    Player(Index).Char(CharNum).y = 0
    Player(Index).Char(CharNum).Dir = 0
    
    For n = 1 To MAX_INV
        Player(Index).Char(CharNum).Inv(n).Num = 0
        Player(Index).Char(CharNum).Inv(n).Value = 0
        Player(Index).Char(CharNum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SKILLS
        Player(Index).Char(CharNum).Skills(n).Num = 0
        Player(Index).Char(CharNum).Skills(n).Level = 0
        Player(Index).Char(CharNum).Skills(n).EXP = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SPELLS
        Player(Index).Char(CharNum).Spells(n).Num = 0
        Player(Index).Char(CharNum).Spells(n).Level = 0
        Player(Index).Char(CharNum).Spells(n).EXP = 0
    Next n
    
    For n = 1 To MAX_PLAYER_QUESTS
        Player(Index).Char(CharNum).Quests(n).Num = 0
        Player(Index).Char(CharNum).Quests(n).SetMap = 0
        Player(Index).Char(CharNum).Quests(n).SetBy = 0
        Player(Index).Char(CharNum).Quests(n).Value = 0
        Player(Index).Char(CharNum).Quests(n).Count = 0
    Next n
    
    For n = 1 To MAX_PLAYER_MAPS
        Player(Index).Char(CharNum).Maps(n).Num = 0
    Next n
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = ""
    
    Item(Index).Pic = 0
    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearNpc(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Npc(Index).Name = ""
    Npc(Index).Sprite = 0
    Npc(Index).SpawnSecs = 0
    Npc(Index).Behavior = 0
    Npc(Index).Range = 0
    Npc(Index).STR = 0
    Npc(Index).DEF = 0
    Npc(Index).SPEED = 0
    Npc(Index).MAGI = 0
    Npc(Index).Big = 0
    Npc(Index).MaxHp = 0
    Npc(Index).Respawn = 0
    Npc(Index).HitOnlyWith = 0
    Npc(Index).ShopLink = 0
    Npc(Index).ExpType = 0
    Npc(Index).EXP = 0
    For i = 1 To MAX_NPC_QUESTS
        Npc(Index).QuestNPC(i) = 0
    Next i
    For i = 1 To MAX_NPC_DROPS
        Npc(Index).ItemNPC(i).Chance = 0
        Npc(Index).ItemNPC(i).ItemNum = 0
        Npc(Index).ItemNPC(i).ItemValue = 0
    Next i
End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    MapItem(MapNum, Index).Num = 0
    MapItem(MapNum, Index).Value = 0
    MapItem(MapNum, Index).Dur = 0
    MapItem(MapNum, Index).x = 0
    MapItem(MapNum, Index).y = 0
End Sub

Sub ClearMapItems()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next x
    Next y
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    MapNpc(MapNum, Index).Num = 0
    MapNpc(MapNum, Index).Target = 0
    MapNpc(MapNum, Index).HP = 0
    MapNpc(MapNum, Index).MP = 0
    MapNpc(MapNum, Index).SP = 0
    MapNpc(MapNum, Index).x = 0
    MapNpc(MapNum, Index).y = 0
    MapNpc(MapNum, Index).Dir = 0
    
    ' Server use only
    MapNpc(MapNum, Index).SpawnWait = 0
    MapNpc(MapNum, Index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next x
    Next y
End Sub

Sub ClearMapResource(ByVal Index As Long, ByVal MapNum As Long)
    MapResource(MapNum, Index).Num = 0
    
    MapResource(MapNum, Index).HP = 0
    MapResource(MapNum, Index).x = 0
    MapResource(MapNum, Index).y = 0
    
    ' Server use only
    MapResource(MapNum, Index).SpawnWait = 0
End Sub

Sub ClearMapResources()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_RESOURCES
            Call ClearMapResource(x, y)
        Next x
    Next y
End Sub

Sub ClearMap(ByVal MapNum As Long)
Dim i As Long
Dim x As Long
Dim y As Long

    Map(MapNum).Name = ""
    Map(MapNum).Owner = ""
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).Left = 0
    Map(MapNum).Right = 0
    Map(MapNum).Music = 0
    Map(MapNum).BootMap = 0
    Map(MapNum).BootX = 0
    Map(MapNum).BootY = 0
    Map(MapNum).Indoors = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map(MapNum).Tile(x, y).Tileset = 0
            Map(MapNum).Tile(x, y).Ground = 0
            Map(MapNum).Tile(x, y).Mask = 0
            Map(MapNum).Tile(x, y).Mask2 = 0
            Map(MapNum).Tile(x, y).Anim = 0
            Map(MapNum).Tile(x, y).Fringe = 0
            Map(MapNum).Tile(x, y).Fringe2 = 0
            Map(MapNum).Tile(x, y).FAnim = 0
            Map(MapNum).Tile(x, y).Light = 0
            Map(MapNum).Tile(x, y).Type = 0
            Map(MapNum).Tile(x, y).Data1 = 0
            Map(MapNum).Tile(x, y).Data2 = 0
            Map(MapNum).Tile(x, y).Data3 = 0
            Map(MapNum).Tile(x, y).WalkUp = 0
            Map(MapNum).Tile(x, y).WalkDown = 0
            Map(MapNum).Tile(x, y).WalkLeft = 0
            Map(MapNum).Tile(x, y).WalkRight = 0
        Next x
    Next y
    
    For i = 1 To MAX_MAP_NPCS
        Map(MapNum).NSpawn(i).NSx = 0
        Map(MapNum).NSpawn(i).NSy = 0
    Next i
    
    For i = 1 To MAX_MAP_RESOURCES
        Map(MapNum).RSpawn(i).RSx = 0
        Map(MapNum).RSpawn(i).RSy = 0
    Next i
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
End Sub

Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next i
End Sub

Sub ClearShop(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Shop(Index).Name = ""
    Shop(Index).FixesItems = 0
    
    For i = 1 To MAX_TRADES
        For n = 1 To MAX_GIVE_ITEMS
            Shop(Index).TradeItem(i).GiveItem(n) = 0
        Next n
        For n = 1 To MAX_GIVE_VALUE
            Shop(Index).TradeItem(i).GiveValue(n) = 0
        Next n
        For n = 1 To MAX_GET_ITEMS
            Shop(Index).TradeItem(i).GetItem(n) = 0
        Next n
        For n = 1 To MAX_GET_VALUE
            Shop(Index).TradeItem(i).GetValue(n) = 0
        Next n
        Shop(Index).ItemStock(i) = 0
    Next i
End Sub

Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i
End Sub

Sub ClearSpell(ByVal Index As Long)
    Spell(Index).Name = ""
    Spell(Index).SpellSprite = 0
    Spell(Index).ClassReq = 0
    Spell(Index).LevelReq = 0
    Spell(Index).Type = 0
    Spell(Index).Data1 = 0
    Spell(Index).Data2 = 0
    Spell(Index).Data3 = 0
End Sub

Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next i
End Sub

Sub ClearSkill(ByVal Index As Long)
    Skill(Index).Name = ""
    Skill(Index).SkillSprite = 0
    Skill(Index).ClassReq = 0
    Skill(Index).LevelReq = 0
    Skill(Index).Type = 0
    Skill(Index).Data1 = 0
    Skill(Index).Data2 = 0
    Skill(Index).Data3 = 0
End Sub

Sub ClearSkills()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSkill(i)
    Next i
End Sub

Sub ClearQuest(ByVal Index As Long)
    Quest(Index).Name = ""
    Quest(Index).Description = ""
    Quest(Index).SetBy = 0
    Quest(Index).ClassReq = 0
    Quest(Index).LevelMin = 0
    Quest(Index).LevelMax = 0
    Quest(Index).Type = 0
    Quest(Index).Reward = 0
    Quest(Index).RewardValue = 0
    Quest(Index).Data1 = 0
    Quest(Index).Data2 = 0
    Quest(Index).Data3 = 0
End Sub

Sub ClearQuests()
Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next i
End Sub

Sub ClearGUI(ByVal Index As Long)
Dim i As Long

    GUI(Index).Name = ""
    GUI(Index).Designer = ""
    GUI(Index).Revision = 0
    For i = 1 To 7
        If i = 1 Then
            GUI(Index).Background(i).Data1 = 16
            GUI(Index).Background(i).Data2 = 16
            GUI(Index).Background(i).Data3 = 480
            GUI(Index).Background(i).Data4 = 640
            GUI(Index).Background(i).Data5 = 0
        ElseIf i = 7 Then
            GUI(Index).Background(i).Data1 = 160
            GUI(Index).Background(i).Data2 = 16
            GUI(Index).Background(i).Data3 = 121
            GUI(Index).Background(i).Data4 = 409
            GUI(Index).Background(i).Data5 = 0
        Else
            GUI(Index).Background(i).Data1 = 272
            GUI(Index).Background(i).Data2 = 16
            GUI(Index).Background(i).Data3 = 105
            GUI(Index).Background(i).Data4 = 297
            GUI(Index).Background(i).Data5 = 0
        End If
    Next i
    For i = 1 To 5
        GUI(Index).Menu(i).Data1 = 32
        GUI(Index).Menu(i).Data2 = 24 + ((i * 64) - 64)
        GUI(Index).Menu(i).Data3 = 41
        GUI(Index).Menu(i).Data4 = 169
    Next i
    For i = 1 To 4
        GUI(Index).Login(i).Data1 = 24
        GUI(Index).Login(i).Data2 = 8 + ((i * 24) - 24)
        GUI(Index).Login(i).Data3 = 17
        GUI(Index).Login(i).Data4 = 185
    Next i
    For i = 1 To 4
        GUI(Index).NewAcc(i).Data1 = 24
        GUI(Index).NewAcc(i).Data2 = 8 + ((i * 24) - 24)
        GUI(Index).NewAcc(i).Data3 = 17
        GUI(Index).NewAcc(i).Data4 = 185
    Next i
    For i = 1 To 4
        GUI(Index).DelAcc(i).Data1 = 24
        GUI(Index).DelAcc(i).Data2 = 8 + ((i * 24) - 24)
        GUI(Index).DelAcc(i).Data3 = 17
        GUI(Index).DelAcc(i).Data4 = 185
    Next i
    GUI(Index).Credits(1).Data1 = 16
    GUI(Index).Credits(1).Data2 = 8
    GUI(Index).Credits(1).Data3 = 65
    GUI(Index).Credits(1).Data4 = 265
    GUI(Index).Credits(2).Data1 = 192
    GUI(Index).Credits(2).Data2 = 80
    GUI(Index).Credits(2).Data3 = 17
    GUI(Index).Credits(2).Data4 = 89
    For i = 1 To 5
        If i = 1 Then
            GUI(Index).Chars(i).Data1 = 16
            GUI(Index).Chars(i).Data2 = 16
            GUI(Index).Chars(i).Data3 = 73
            GUI(Index).Chars(i).Data4 = 113
        Else
            GUI(Index).Chars(i).Data1 = 144
            GUI(Index).Chars(i).Data2 = 8 + ((i * 24) - 48)
            GUI(Index).Chars(i).Data3 = 17
            GUI(Index).Chars(i).Data4 = 121
        End If
    Next i
    GUI(Index).NewChar(1).Data1 = 16
    GUI(Index).NewChar(1).Data2 = 8
    GUI(Index).NewChar(1).Data3 = 15
    GUI(Index).NewChar(1).Data4 = 70
    GUI(Index).NewChar(2).Data1 = 16
    GUI(Index).NewChar(2).Data2 = 24
    GUI(Index).NewChar(2).Data3 = 15
    GUI(Index).NewChar(2).Data4 = 70
    GUI(Index).NewChar(3).Data1 = 16
    GUI(Index).NewChar(3).Data2 = 40
    GUI(Index).NewChar(3).Data3 = 15
    GUI(Index).NewChar(3).Data4 = 70
    GUI(Index).NewChar(4).Data1 = 96
    GUI(Index).NewChar(4).Data2 = 8
    GUI(Index).NewChar(4).Data3 = 15
    GUI(Index).NewChar(4).Data4 = 100
    GUI(Index).NewChar(5).Data1 = 96
    GUI(Index).NewChar(5).Data2 = 24
    GUI(Index).NewChar(5).Data3 = 15
    GUI(Index).NewChar(5).Data4 = 100
    GUI(Index).NewChar(6).Data1 = 96
    GUI(Index).NewChar(6).Data2 = 40
    GUI(Index).NewChar(6).Data3 = 15
    GUI(Index).NewChar(6).Data4 = 100
    GUI(Index).NewChar(7).Data1 = 96
    GUI(Index).NewChar(7).Data2 = 56
    GUI(Index).NewChar(7).Data3 = 15
    GUI(Index).NewChar(7).Data4 = 100
    GUI(Index).NewChar(8).Data1 = 96
    GUI(Index).NewChar(8).Data2 = 72
    GUI(Index).NewChar(8).Data3 = 15
    GUI(Index).NewChar(8).Data4 = 100
    GUI(Index).NewChar(9).Data1 = 208
    GUI(Index).NewChar(9).Data2 = 8
    GUI(Index).NewChar(9).Data3 = 17
    GUI(Index).NewChar(9).Data4 = 185
    GUI(Index).NewChar(10).Data1 = 208
    GUI(Index).NewChar(10).Data2 = 32
    GUI(Index).NewChar(10).Data3 = 17
    GUI(Index).NewChar(10).Data4 = 185
    GUI(Index).NewChar(11).Data1 = 208
    GUI(Index).NewChar(11).Data2 = 56
    GUI(Index).NewChar(11).Data3 = 17
    GUI(Index).NewChar(11).Data4 = 90
    GUI(Index).NewChar(12).Data1 = 304
    GUI(Index).NewChar(12).Data2 = 56
    GUI(Index).NewChar(12).Data3 = 17
    GUI(Index).NewChar(12).Data4 = 90
    GUI(Index).NewChar(13).Data1 = 64
    GUI(Index).NewChar(13).Data2 = 96
    GUI(Index).NewChar(13).Data3 = 17
    GUI(Index).NewChar(13).Data4 = 120
    GUI(Index).NewChar(14).Data1 = 248
    GUI(Index).NewChar(14).Data2 = 96
    GUI(Index).NewChar(14).Data3 = 17
    GUI(Index).NewChar(14).Data4 = 120
End Sub

Sub ClearGUIS()
Dim i As Long

    For i = 1 To MAX_GUIS
        Call ClearGUI(i)
    Next i
End Sub

Sub SetHighIndex()
Dim i As Integer
Dim x As Integer

    For i = 0 To MAX_PLAYERS
        x = MAX_PLAYERS - i
        
        If x = 0 Then
            x = 1
        End If
       
        If IsConnected(x) = True Then
             HighIndex = x
             Exit Sub
        End If
    Next i
   
    HighIndex = 0
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim(Player(Index).Login)
End Function

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Char(Player(Index).CharNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = (GetPlayerLevel(Index) + 1) * (GetPlayerSTR(Index) + GetPlayerDEF(Index) + GetPlayerMAGI(Index) + GetPlayerSPEED(Index) + GetPlayerPOINTS(Index)) * 25
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    Player(Index).Char(Player(Index).CharNum).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If
    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If
    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If
    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim AmuletSlot As Long
Dim RingSlot As Long
Dim i As Long

    CharNum = Player(Index).CharNum
    AmuletSlot = GetPlayerAmuletSlot(Index)
    RingSlot = GetPlayerRingSlot(Index)
    GetPlayerMaxHP = (Player(Index).Char(CharNum).Level + Int(GetPlayerSTR(Index) / 2) + Class(Player(Index).Char(CharNum).Class).STR) * 2
    If AmuletSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDHP Then
            GetPlayerMaxHP = GetPlayerMaxHP + Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
        End If
    End If
    If RingSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDHP Then
            GetPlayerMaxHP = GetPlayerMaxHP + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
        End If
    End If
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim AmuletSlot As Long
Dim RingSlot As Long

    CharNum = Player(Index).CharNum
    AmuletSlot = GetPlayerAmuletSlot(Index)
    RingSlot = GetPlayerRingSlot(Index)
    GetPlayerMaxMP = (Player(Index).Char(CharNum).Level + Int(GetPlayerMAGI(Index) / 2) + Class(Player(Index).Char(CharNum).Class).MAGI) * 2
    If AmuletSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDMP Then
            GetPlayerMaxMP = GetPlayerMaxMP + Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
        End If
    End If
    If RingSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDMP Then
            GetPlayerMaxMP = GetPlayerMaxMP + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
        End If
    End If
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim AmuletSlot As Long
Dim RingSlot As Long

    CharNum = Player(Index).CharNum
    AmuletSlot = GetPlayerAmuletSlot(Index)
    RingSlot = GetPlayerRingSlot(Index)
    GetPlayerMaxSP = (Player(Index).Char(CharNum).Level + Int(GetPlayerSPEED(Index) / 2) + Class(Player(Index).Char(CharNum).Class).SPEED) * 2
    If AmuletSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDSP Then
            GetPlayerMaxSP = GetPlayerMaxSP + Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
        End If
    End If
    If RingSlot > 0 Then
        If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDSP Then
            GetPlayerMaxSP = GetPlayerMaxSP + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
        End If
    End If
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim(Class(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(Class(ClassNum).MAGI / 2) + Class(ClassNum).MAGI) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Class(ClassNum).SPEED / 2) + Class(ClassNum).SPEED) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = Class(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = Class(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = Class(ClassNum).SPEED
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = Class(ClassNum).MAGI
End Function

Function GetClassDEX(ByVal ClassNum As Long) As Long
    GetClassDEX = Class(ClassNum).DEX
End Function

Function GetPlayerSTR(ByVal Index As Long, Optional RAW As Boolean = False) As Long
Dim AmuletSlot As Long
Dim RingSlot As Long
Dim i As Long

    If RAW = False Then
        GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR
        AmuletSlot = GetPlayerAmuletSlot(Index)
        RingSlot = GetPlayerRingSlot(Index)
        For i = 1 To MAX_PLAYER_SKILLS
            If GetPlayerSkill(Index, i) > 0 Then
                If (Skill(GetPlayerSkill(Index, i)).Type = SKILL_TYPE_ATTRIBUTE) And (Skill(GetPlayerSkill(Index, i)).Data1 = SKILL_ATTRIBUTE_STR) Then
                    GetPlayerSTR = GetPlayerSTR + (Skill(GetPlayerSkill(Index, i)).Data2 * GetPlayerSkillLevel(Index, i))
                End If
            End If
        Next i
        If AmuletSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDSTR Then
                GetPlayerSTR = GetPlayerSTR + Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
            End If
        End If
        If RingSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDSTR Then
                GetPlayerSTR = GetPlayerSTR + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
            End If
        End If
    Else
        GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR
    End If
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long, Optional RAW As Boolean = False) As Long
Dim AmuletSlot As Long
Dim RingSlot As Long
Dim i As Long

    If RAW = False Then
        GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF
        AmuletSlot = GetPlayerAmuletSlot(Index)
        RingSlot = GetPlayerRingSlot(Index)
        For i = 1 To MAX_PLAYER_SKILLS
            If GetPlayerSkill(Index, i) > 0 Then
                If (Skill(GetPlayerSkill(Index, i)).Type = SKILL_TYPE_ATTRIBUTE) And (Skill(GetPlayerSkill(Index, i)).Data1 = SKILL_ATTRIBUTE_DEF) Then
                    GetPlayerDEF = GetPlayerDEF + (Skill(GetPlayerSkill(Index, i)).Data2 * GetPlayerSkillLevel(Index, i))
                End If
            End If
        Next i
        If AmuletSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDDEF Then
                GetPlayerDEF = GetPlayerDEF + Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
            End If
        End If
        If RingSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDDEF Then
                GetPlayerDEF = GetPlayerDEF + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
            End If
        End If
    Else
        GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF
    End If
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long, Optional RAW As Boolean = False) As Long
Dim AmuletSlot As Long
Dim RingSlot As Long
Dim i As Long

    If RAW = False Then
        GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).SPEED
        AmuletSlot = GetPlayerAmuletSlot(Index)
        RingSlot = GetPlayerRingSlot(Index)
        For i = 1 To MAX_PLAYER_SKILLS
            If GetPlayerSkill(Index, i) > 0 Then
                If (Skill(GetPlayerSkill(Index, i)).Type = SKILL_TYPE_ATTRIBUTE) And (Skill(GetPlayerSkill(Index, i)).Data1 = SKILL_ATTRIBUTE_SPEED) Then
                    GetPlayerSPEED = GetPlayerSPEED + (Skill(GetPlayerSkill(Index, i)).Data2 * GetPlayerSkillLevel(Index, i))
                End If
            End If
        Next i
        If AmuletSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDSPEED Then
                GetPlayerSPEED = GetPlayerSPEED + Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
            End If
        End If
        If RingSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDSPEED Then
                GetPlayerSPEED = GetPlayerSPEED + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
            End If
        End If
    Else
        GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).SPEED
    End If
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal SPEED As Long)
    Player(Index).Char(Player(Index).CharNum).SPEED = SPEED
End Sub

Function GetPlayerMAGI(ByVal Index As Long, Optional RAW As Boolean = False) As Long
Dim AmuletSlot As Long
Dim RingSlot As Long
Dim i As Long

    If RAW = False Then
        GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).MAGI
        AmuletSlot = GetPlayerAmuletSlot(Index)
        RingSlot = GetPlayerRingSlot(Index)
        For i = 1 To MAX_PLAYER_SKILLS
            If GetPlayerSkill(Index, i) > 0 Then
                If (Skill(GetPlayerSkill(Index, i)).Type = SKILL_TYPE_ATTRIBUTE) And (Skill(GetPlayerSkill(Index, i)).Data1 = SKILL_ATTRIBUTE_MAGI) Then
                    GetPlayerMAGI = GetPlayerMAGI + (Skill(GetPlayerSkill(Index, i)).Data2 * GetPlayerSkillLevel(Index, i))
                End If
            End If
        Next i
        If AmuletSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDMAGI Then
                GetPlayerMAGI = GetPlayerMAGI + Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
            End If
        End If
        If RingSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDMAGI Then
                GetPlayerMAGI = GetPlayerMAGI + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
            End If
        End If
    Else
        GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).MAGI
    End If
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)
    Player(Index).Char(Player(Index).CharNum).MAGI = MAGI
End Sub

Function GetPlayerDEX(ByVal Index As Long, Optional RAW As Boolean = False) As Long
Dim AmuletSlot As Long
Dim RingSlot As Long
Dim i As Long

    If RAW = False Then
        GetPlayerDEX = Player(Index).Char(Player(Index).CharNum).DEX
        AmuletSlot = GetPlayerAmuletSlot(Index)
        RingSlot = GetPlayerRingSlot(Index)
        For i = 1 To MAX_PLAYER_SKILLS
            If GetPlayerSkill(Index, i) > 0 Then
                If (Skill(GetPlayerSkill(Index, i)).Type = SKILL_TYPE_ATTRIBUTE) And (Skill(GetPlayerSkill(Index, i)).Data1 = SKILL_ATTRIBUTE_DEX) Then
                    GetPlayerDEX = GetPlayerDEX + (Skill(GetPlayerSkill(Index, i)).Data2 * GetPlayerSkillLevel(Index, i))
                End If
            End If
        Next i
        If AmuletSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDDEX Then
                GetPlayerDEX = GetPlayerDEX + Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
            End If
        End If
        If RingSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDDEX Then
                GetPlayerDEX = GetPlayerDEX + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
            End If
        End If
    Else
        GetPlayerDEX = Player(Index).Char(Player(Index).CharNum).DEX
    End If
End Function

Sub SetPlayerDEX(ByVal Index As Long, ByVal DEX As Long)
    Player(Index).Char(Player(Index).CharNum).DEX = DEX
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).Char(Player(Index).CharNum).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).Char(Player(Index).CharNum).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = GameServer.Sockets.Item(Index).RemoteAddress
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spells(SpellSlot).Num
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(Player(Index).CharNum).Spells(SpellSlot).Num = SpellNum
End Sub

Function GetPlayerSpellLevel(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpellLevel = Player(Index).Char(Player(Index).CharNum).Spells(SpellSlot).Level
End Function

Sub SetPlayerSpellLevel(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellLevel As Long)
    Player(Index).Char(Player(Index).CharNum).Spells(SpellSlot).Level = SpellLevel
End Sub

Function GetSpellNextLevel(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetSpellNextLevel = ((GetPlayerSpellLevel(Index, SpellSlot) + 1) * GetPlayerMAGI(Index, True)) * 10
End Function

Function GetPlayerSpellExp(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpellExp = Player(Index).Char(Player(Index).CharNum).Spells(SpellSlot).EXP
End Function

Sub SetPlayerSpellExp(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellExp As Long)
    Player(Index).Char(Player(Index).CharNum).Spells(SpellSlot).EXP = SpellExp
End Sub

Function GetPlayerSkill(ByVal Index As Long, ByVal SkillSlot As Long) As Long
    GetPlayerSkill = Player(Index).Char(Player(Index).CharNum).Skills(SkillSlot).Num
End Function

Sub SetPlayerSkill(ByVal Index As Long, ByVal SkillSlot As Long, ByVal SkillNum As Long)
    Player(Index).Char(Player(Index).CharNum).Skills(SkillSlot).Num = SkillNum
End Sub

Function GetPlayerSkillLevel(ByVal Index As Long, ByVal SkillSlot As Long) As Long
    GetPlayerSkillLevel = Player(Index).Char(Player(Index).CharNum).Skills(SkillSlot).Level
End Function

Sub SetPlayerSkillLevel(ByVal Index As Long, ByVal SkillSlot As Long, ByVal SkillLevel As Long)
    Player(Index).Char(Player(Index).CharNum).Skills(SkillSlot).Level = SkillLevel
End Sub

Function GetSkillNextLevel(ByVal Index As Long, ByVal SkillSlot As Long) As Long
    GetSkillNextLevel = (GetPlayerSkillLevel(Index, SkillSlot) + 1) * (GetPlayerSTR(Index, True) + GetPlayerDEF(Index, True) + GetPlayerMAGI(Index, True) + GetPlayerSPEED(Index, True) + GetPlayerPOINTS(Index)) * 10
End Function

Function GetPlayerSkillExp(ByVal Index As Long, ByVal SkillSlot As Long) As Long
    GetPlayerSkillExp = Player(Index).Char(Player(Index).CharNum).Skills(SkillSlot).EXP
End Function

Sub SetPlayerSkillExp(ByVal Index As Long, ByVal SkillSlot As Long, ByVal SkillExp As Long)
    Player(Index).Char(Player(Index).CharNum).Skills(SkillSlot).EXP = SkillExp
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
End Sub

Function GetPlayerAmuletSlot(ByVal Index As Long) As Long
    GetPlayerAmuletSlot = Player(Index).Char(Player(Index).CharNum).AmuletSlot
End Function

Sub SetPlayerAmuletSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).AmuletSlot = InvNum
End Sub

Function GetPlayerRingSlot(ByVal Index As Long) As Long
    GetPlayerRingSlot = Player(Index).Char(Player(Index).CharNum).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).RingSlot = InvNum
End Sub

Function GetPlayerArrowSlot(ByVal Index As Long) As Long
    GetPlayerArrowSlot = Player(Index).Char(Player(Index).CharNum).ArrowSlot
End Function

Sub SetPlayerArrowSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ArrowSlot = InvNum
End Sub

Function GetPlayerQuest(ByVal Index As Long, ByVal QuestSlot As Long) As Long
    GetPlayerQuest = Player(Index).Char(Player(Index).CharNum).Quests(QuestSlot).Num
End Function

Function GetPlayerQuestMap(ByVal Index As Long, ByVal QuestSlot As Long) As Long
    GetPlayerQuestMap = Player(Index).Char(Player(Index).CharNum).Quests(QuestSlot).SetMap
End Function

Function GetPlayerQuestBy(ByVal Index As Long, ByVal QuestSlot As Long) As Long
    GetPlayerQuestBy = Player(Index).Char(Player(Index).CharNum).Quests(QuestSlot).SetBy
End Function

Function GetPlayerQuestValue(ByVal Index As Long, ByVal QuestSlot As Long) As Long
    GetPlayerQuestValue = Player(Index).Char(Player(Index).CharNum).Quests(QuestSlot).Value
End Function

Function GetPlayerQuestCount(ByVal Index As Long, ByVal QuestSlot As Long) As Long
    GetPlayerQuestCount = Player(Index).Char(Player(Index).CharNum).Quests(QuestSlot).Count
End Function

Sub SetPlayerQuest(ByVal Index As Long, ByVal QuestSlot As Long, ByVal NpcNum As Long, ByVal QuestNum As Long)
    With Player(Index).Char(Player(Index).CharNum).Quests(QuestSlot)
        .Num = QuestNum
        .SetMap = GetPlayerMap(Index)
        .SetBy = NpcNum
        .Value = Quest(QuestNum).Data2
        .Count = 0
    End With
End Sub

Sub ClearPlayerQuest(ByVal Index As Long, ByVal QuestSlot As Long)
    With Player(Index).Char(Player(Index).CharNum).Quests(QuestSlot)
        .Num = 0
        .SetMap = 0
        .SetBy = 0
        .Value = 0
        .Count = 0
    End With
End Sub
