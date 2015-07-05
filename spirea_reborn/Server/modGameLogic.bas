Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal index As Long) As Long
Dim WeaponSlot As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    GetPlayerDamage = Int(GetPlayerSTR(index) / 2)
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
    
    If GetPlayerWeaponSlot(index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(index)
        
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(index, WeaponSlot)).Data2
        
        Call SetPlayerInvItemDur(index, WeaponSlot, GetPlayerInvItemDur(index, WeaponSlot) - 1)
        
        If GetPlayerInvItemDur(index, WeaponSlot) <= 0 Then
            Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " has broken.", Yellow)
            Call TakeItem(index, GetPlayerInvItemNum(index, WeaponSlot), 0)
        Else
            If GetPlayerInvItemDur(index, WeaponSlot) <= 5 Then
                Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " is about to break!", Yellow)
            End If
        End If
    End If
End Function

Function GetPlayerProtection(ByVal index As Long) As Long
Dim ArmorSlot As Long, HelmSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ArmorSlot = GetPlayerArmorSlot(index)
    HelmSlot = GetPlayerHelmetSlot(index)
    GetPlayerProtection = Int(GetPlayerDEF(index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, ArmorSlot)).Data2
        Call SetPlayerInvItemDur(index, ArmorSlot, GetPlayerInvItemDur(index, ArmorSlot) - 1)
        
        If GetPlayerInvItemDur(index, ArmorSlot) <= 0 Then
            Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " has broken.", Yellow)
            Call TakeItem(index, GetPlayerInvItemNum(index, ArmorSlot), 0)
        Else
            If GetPlayerInvItemDur(index, ArmorSlot) <= 5 Then
                Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " is about to break!", Yellow)
            End If
        End If
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, HelmSlot)).Data2
        Call SetPlayerInvItemDur(index, HelmSlot, GetPlayerInvItemDur(index, HelmSlot) - 1)
        
        If GetPlayerInvItemDur(index, HelmSlot) <= 0 Then
            Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " has broken.", Yellow)
            Call TakeItem(index, GetPlayerInvItemNum(index, HelmSlot), 0)
        Else
            If GetPlayerInvItemDur(index, HelmSlot) <= 5 Then
                Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " is about to break!", Yellow)
            End If
        End If
    End If
End Function

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

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
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

Function FindOpenSpellSlot(ByVal index As Long) As Long
Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next i
End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next i
End Function

Function TotalOnlinePlayers() As Long
Dim i As Long

    TotalOnlinePlayers = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next i
End Function

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
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

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next i
End Function

Sub TakeItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(index, i)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(index) > 0 Then
                            If i = GetPlayerWeaponSlot(index) Then
                                Call SetPlayerWeaponSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                
                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(index) > 0 Then
                            If i = GetPlayerArmorSlot(index) Then
                                Call SetPlayerArmorSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(index) > 0 Then
                            If i = GetPlayerHelmetSlot(index) Then
                                Call SetPlayerHelmetSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(index) > 0 Then
                            If i = GetPlayerShieldSlot(index) Then
                                Call SetPlayerShieldSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select

                
                n = Item(GetPlayerInvItemNum(index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) Then
                    TakeItem = True
                End If
            End If
                            
            If TakeItem = True Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Call SetPlayerInvItemDur(index, i, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = FindOpenInvSlot(index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, ItemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(index, i, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(index, i)
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
    End If
End Sub

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
            
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & x & SEP_CHAR & y & END_CHAR
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

Sub PlayerMapGetItem(ByVal index As Long)
Dim i As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String

    If IsPlaying(index) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(index)
    
    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).x = GetPlayerX(index)) And (MapItem(MapNum, i).y = GetPlayerY(index)) Then
                ' Find open slot
                n = FindOpenInvSlot(index, MapItem(MapNum, i).Num)
                
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(index, n, MapItem(MapNum, i).Num)
                    If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
                        Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).Value)
                        Msg = "You picked up " & MapItem(MapNum, i).Value & " " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(index, n, 0)
                        Msg = "You picked up a " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    End If
                    Call SetPlayerInvItemDur(index, n, MapItem(MapNum, i).Dur)
                        
                    ' Erase item from the map
                    MapItem(MapNum, i).Num = 0
                    MapItem(MapNum, i).Value = 0
                    MapItem(MapNum, i).Dur = 0
                    MapItem(MapNum, i).x = 0
                    MapItem(MapNum, i).y = 0
                        
                    Call SendInventoryUpdate(index, n)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    Call PlayerMsg(index, Msg, Yellow)
                    Exit Sub
                Else
                    Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                    Exit Sub
                End If
            End If
        End If
    Next i
End Sub

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Ammount As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(index))
        
        If i <> 0 Then
            MapItem(GetPlayerMap(index), i).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(index) Then
                        Call SetPlayerArmorSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(index) Then
                        Call SetPlayerWeaponSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(index) Then
                        Call SetPlayerHelmetSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(index) Then
                        Call SetPlayerShieldSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
            MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                        
            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Ammount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, InvNum)
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), i).Value = Ammount
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Ammount & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Ammount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(index), i).Value = 0
                If Item(GetPlayerInvItemNum(index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & " " & GetPlayerInvItemDur(index, InvNum) & "/" & Item(GetPlayerInvItemNum(index, InvNum)).Data1 & ".", Yellow)
                Else
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, Ammount, MapItem(GetPlayerMap(index), i).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsg(index, "To many items already on the ground.", BrightRed)
        End If
    End If
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Packet As String
Dim npcnum As Long
Dim i As Long, x As Long, y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    Spawned = False
    
    npcnum = Map(MapNum).Npc(MapNpcNum)
    If npcnum > 0 Then
        MapNpc(MapNum, MapNpcNum).Num = npcnum
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxHP(npcnum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(npcnum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(npcnum)
                
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)
        
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
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & END_CHAR
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

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    CanAttackPlayer = False
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Then
        Exit Function
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerHP(Victim) <= 0 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If Player(Victim).GettingMap = YES Then
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If (GetPlayerMap(Attacker) = GetPlayerMap(Victim)) And (GetTickCount > Player(Attacker).AttackTimer + 950) Then
        
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                        Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                    Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
                            End If
                        End If
                    End If
                End If
            
            Case DIR_DOWN
                If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                        Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                    Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
                            End If
                        End If
                    End If
                End If
        
            Case DIR_LEFT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                        Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                    Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
                            End If
                        End If
                    End If
                End If
            
            Case DIR_RIGHT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                        Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                    Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
                            End If
                        End If
                    End If
                End If
        End Select
    End If
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, npcnum As Long
Dim i As Integer
    CanAttackNpc = False
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker), MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Attacker)
    npcnum = MapNpc(MapNum, MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If npcnum > 0 And GetTickCount > Player(Attacker).AttackTimer + 950 Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If (MapNpc(MapNum, MapNpcNum).y + 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x = GetPlayerX(Attacker)) Then
                      If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            If Npc(npcnum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                                For i = 1 To MAX_SHOPS
                                    If Trim(Shop(i).Name) = Trim(Npc(npcnum).Name) Then
                                        Call SendTrade(Attacker, i)
                                    End If
                                Next i
                            Else
                                Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(npcnum).Name) & "!", BrightBlue)
                            End If
                        End If
                    End If

                
                Case DIR_DOWN
                    If (MapNpc(MapNum, MapNpcNum).y - 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x = GetPlayerX(Attacker)) Then
                        If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            If Npc(npcnum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                                For i = 1 To MAX_SHOPS
                                    If Trim(Shop(i).Name) = Trim(Npc(npcnum).Name) Then
                                        Call SendTrade(Attacker, i)
                                    End If
                                Next i
                            Else
                                Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(npcnum).Name) & "!", BrightBlue)
                            End If
                        End If
                    End If


                
                Case DIR_LEFT
                    If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x + 1 = GetPlayerX(Attacker)) Then
                         If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            If Npc(npcnum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                                For i = 1 To MAX_SHOPS
                                    If Trim(Shop(i).Name) = Trim(Npc(npcnum).Name) Then
                                        Call SendTrade(Attacker, i)
                                    End If
                                Next i
                            Else
                                Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(npcnum).Name) & "!", BrightBlue)
                            End If
                        End If
                    End If

                
                Case DIR_RIGHT
                    If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x - 1 = GetPlayerX(Attacker)) Then
                       If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            If Npc(npcnum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                                For i = 1 To MAX_SHOPS
                                    If Trim(Shop(i).Name) = Trim(Npc(npcnum).Name) Then
                                        Call SendTrade(Attacker, i)
                                    End If
                                Next i

                                Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(npcnum).Name) & "!", BrightBlue)
                            End If
                        End If
                    End If

            End Select
        End If
    End If
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
Dim MapNum As Long, npcnum As Long
    
    CanNpcAttackPlayer = False
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(index) = False Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index), MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(index)
    npcnum = MapNpc(MapNum, MapNpcNum).Num
    
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
        If npcnum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) + 1 = MapNpc(MapNum, MapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
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
Dim Exp As Long
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
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & END_CHAR)
        
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        
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

        ' Calculate exp to give attacker
        Exp = Int(GetPlayerExp(Victim) / 10)
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If
        
        If Exp = 0 Then
            Call PlayerMsg(Victim, "You lost no experience points.", BrightRed)
            Call PlayerMsg(Attacker, "You received no experience points from that weak insignificant player.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            Call PlayerMsg(Victim, "You lost " & Exp & " experience points.", BrightRed)
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You got " & Exp & " experience points for killing " & GetPlayerName(Victim) & ".", BrightBlue)
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
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If
        Else
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
    End If
    
    ' Reset timer for attacking
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
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
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & END_CHAR)
    
    MapNum = GetPlayerMap(Victim)
    Name = Trim(Npc(MapNpc(MapNum, MapNpcNum).Num).Name)
    
    If Damage >= GetPlayerHP(Victim) Then
        ' Say damage
        Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by a " & Name, BrightRed)
        
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
        
        ' Calculate exp to give attacker
        Exp = Int(GetPlayerExp(Victim) / 3)
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If
        
        If Exp = 0 Then
            Call PlayerMsg(Victim, "You lost no experience points.", BrightRed)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            Call PlayerMsg(Victim, "You lost " & Exp & " experience points.", BrightRed)
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
        Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
    End If
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim n As Long, i As Long
Dim STR As Long, DEF As Long, MapNum As Long, npcnum As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    Else
        n = 0
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    npcnum = MapNpc(MapNum, MapNpcNum).Num
    Name = Trim(Npc(npcnum).Name)
        
    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points, killing it.", BrightRed)
        Else
            Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points, killing it.", BrightRed)
        End If
                        
        ' Calculate exp to give attacker
        STR = Npc(npcnum).STR
        DEF = Npc(npcnum).DEF
        Exp = STR * DEF * 2
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If Player(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You have gained " & Exp & " experience points.", BrightBlue)
        Else
            Exp = Exp / 2
            
            If Exp < 0 Then
                Exp = 1
            End If
            
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You have gained " & Exp & " party experience points.", BrightBlue)
            
            n = Player(Attacker).PartyPlayer
            If n > 0 Then
                Call SetPlayerExp(n, GetPlayerExp(n) + Exp)
                Call PlayerMsg(n, "You have gained " & Exp & " party experience points.", BrightBlue)
            End If
        End If
                                
        ' Drop the goods if they get it
        n = Int(Rnd * Npc(npcnum).DropChance) + 1
        If n = 1 Then
            Call SpawnItem(Npc(npcnum).DropItem, Npc(npcnum).DropItemValue, MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).Num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
        
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
        MapNpc(MapNum, MapNpcNum).HP = MapNpc(MapNum, MapNpcNum).HP - Damage
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        Else
            Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If
        
        ' Check if we should send a message
        If MapNpc(MapNum, MapNpcNum).Target = 0 And MapNpc(MapNum, MapNpcNum).Target <> Attacker Then
            If Trim(Npc(npcnum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, "A " & Trim(Npc(npcnum).Name) & " says, '" & Trim(Npc(npcnum).AttackSay) & "' to you.", SayColor)
            End If
        End If
        
        ' Set the NPC target to the player
        MapNpc(MapNum, MapNpcNum).Target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum, MapNpcNum).Num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum, i).Num = MapNpc(MapNum, MapNpcNum).Num Then
                    MapNpc(MapNum, i).Target = Attacker
                End If
            Next i
        End If
    End If
    
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim Packet As String
Dim ShopNum As Long, OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    ShopNum = Map(GetPlayerMap(index)).Shop
    If ShopNum > 0 Then
        If Trim(Shop(ShopNum).LeaveSay) <> "" Then
            Call PlayerMsg(index, Trim(Shop(ShopNum).Name) & " says, '" & Trim(Shop(ShopNum).LeaveSay) & "'", SayColor)
        End If
    End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    Call SendLeaveMap(index, OldMap)
    
    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    
    ' Check if there is an npc on the map and say hello if so
    ShopNum = Map(GetPlayerMap(index)).Shop
    If ShopNum > 0 Then
        If Trim(Shop(ShopNum).JoinSay) <> "" Then
            Call PlayerMsg(index, Trim(Shop(ShopNum).Name) & " says, '" & Trim(Shop(ShopNum).JoinSay) & "'", SayColor)
        End If
    End If
            
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    
    Player(index).GettingMap = YES
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim MapNum As Long
Dim x As Long
Dim y As Long
Dim i As Long
Dim Moved As Byte

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    Call SetPlayerDir(index, Dir)
    
    Moved = NO
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), MAX_MAPY)
                    Moved = YES
                End If
            End If
                    
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < MAX_MAPY Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                End If
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, MAX_MAPX, GetPlayerY(index))
                    Moved = YES
                End If
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < MAX_MAPX Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                End If
            End If
    End Select
        
    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        x = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        y = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3
                        
        Call PlayerWarp(index, MapNum, x, y)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KEYOPEN Then
        x = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        y = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        
        If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
            TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
            Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
        End If
    End If
    
    ' They tried to hack
    If Moved = NO Then
        Call HackingAttempt(index, "Position Modification")
    End If
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
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
                For i = 1 To MAX_PLAYERS
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
                For i = 1 To MAX_PLAYERS
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
                For i = 1 To MAX_PLAYERS
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
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_DOWN
            MapNpc(MapNum, MapNpcNum).y = MapNpc(MapNum, MapNpcNum).y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_LEFT
            MapNpc(MapNum, MapNpcNum).x = MapNpc(MapNum, MapNpcNum).x - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).x = MapNpc(MapNum, MapNpcNum).x + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
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
    Packet = "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub JoinGame(ByVal index As Long)
    ' Set the flag so we know the person is in the game
    Player(index).InGame = True
        
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITER Then
        Call GlobalMsg(GetPlayerName(index) & " has joined " & GAME_NAME & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has joined " & GAME_NAME & "!", White)
    End If
        
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(index, "LOGINOK" & SEP_CHAR & index & END_CHAR)
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)

    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendStats(index)
    Call SendWeatherTo(index)
    Call SendTimeTo(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(index, "INGAME" & END_CHAR)
    frmServer.lvwInfo.Refresh
    Call UsersOnline_Start(index)
End Sub

Sub LeftGame(ByVal index As Long)
Dim n As Long

    If Player(index).InGame = True Then
        Player(index).InGame = False
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) = 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' Check for boot map
        If Map(GetPlayerMap(index)).BootMap > 0 Then
            Call SetPlayerX(index, Map(GetPlayerMap(index)).BootX)
            Call SetPlayerY(index, Map(GetPlayerMap(index)).BootY)
            Call SetPlayerMap(index, Map(GetPlayerMap(index)).BootMap)
        End If
        
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If Player(index).InParty = YES Then
            n = Player(index).PartyPlayer
            
            Call PlayerMsg(n, GetPlayerName(index) & " has left " & GAME_NAME & ", disbanning party.", Pink)
            Player(n).InParty = NO
            Player(n).PartyPlayer = 0
        End If
            
        Call SavePlayer(index)
    
        ' Send a global message that he/she left
        If GetPlayerAccess(index) <= ADMIN_MONITER Then
            Call GlobalMsg(GetPlayerName(index) & " has left " & GAME_NAME & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has left " & GAME_NAME & "!", White)
        End If
        Call TextAdd(frmServer.txtText, GetPlayerName(index) & " has disconnected from " & GAME_NAME & ".", True)
        Call SendLeftGame(index)
    End If
    
    Call ClearPlayer(index)
    Call UsersOnline_Start(index)
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
Dim i As Long, n As Long

    n = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If
    Next i
    
    GetTotalMapPlayers = n
End Function

Function GetNpcMaxHP(ByVal npcnum As Long)
Dim x As Long, y As Long

    ' Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcMaxHP = 0
        Exit Function
    End If
    
    x = Npc(npcnum).STR
    y = Npc(npcnum).DEF
    GetNpcMaxHP = x * y
End Function

Function GetNpcMaxMP(ByVal npcnum As Long)
    ' Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcMaxMP = 0
        Exit Function
    End If
        
    GetNpcMaxMP = Npc(npcnum).MAGI * 2
End Function

Function GetNpcMaxSP(ByVal npcnum As Long)
    ' Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcMaxSP = 0
        Exit Function
    End If
        
    GetNpcMaxSP = Npc(npcnum).SPEED * 2
End Function

Function GetPlayerHPRegen(ByVal index As Long)
Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerHPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerDEF(index) / 2)
    If i < 2 Then i = 2
    
    GetPlayerHPRegen = i
End Function

Function GetPlayerMPRegen(ByVal index As Long)
Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerMPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerMAGI(index) / 2)
    If i < 2 Then i = 2
    
    GetPlayerMPRegen = i
End Function

Function GetPlayerSPRegen(ByVal index As Long)
Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerSPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerSPEED(index) / 2)
    If i < 2 Then i = 2
    
    GetPlayerSPRegen = i
End Function

Function GetNpcHPRegen(ByVal npcnum As Long)
Dim i As Long

    'Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcHPRegen = 0
        Exit Function
    End If
    
    i = Int(Npc(npcnum).DEF / 3)
    If i < 1 Then i = 1
    
    GetNpcHPRegen = i
End Function

Sub CheckPlayerLevelUp(ByVal index As Long)
Dim i As Long

    ' Check if attacker got a level up
    If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
        Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
                    
        ' Get the ammount of skill points to add
        i = Int(GetPlayerSPEED(index) / 10)
        If i < 1 Then i = 1
        If i > 3 Then i = 3
            
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)
        Call SetPlayerExp(index, 0)
        Call GlobalMsg(GetPlayerName(index) & " has gained a level!", Brown)
        Call PlayerMsg(index, "You have gained a level!  You now have " & GetPlayerPOINTS(index) & " stat points to distribute.", BrightBlue)
    End If
End Sub

Sub CastSpell(ByVal index As Long, ByVal SpellSlot As Long)
Dim SpellNum As Long, MPReq As Long, i As Long, n As Long, Damage As Long
Dim Casted As Boolean

    Casted = False
    
    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    SpellNum = GetPlayerSpell(index, SpellSlot)
    
    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then
        Call PlayerMsg(index, "You do not have this spell!", BrightRed)
        Exit Sub
    End If
    
    i = GetSpellReqLevel(index, SpellNum)
    MPReq = (i + Spell(SpellNum).Data1 + Spell(SpellNum).Data2 + Spell(SpellNum).Data3)
    
    ' Check if they have enough MP
    If GetPlayerMP(index) < MPReq Then
        Call PlayerMsg(index, "Not enough mana points!", BrightRed)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If i > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & i & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < Player(index).AttackTimer + 1000 Then
        Exit Sub
    End If
    
    ' Check if the spell is a give item and do that instead of a stat modification
    If Spell(SpellNum).Type = SPELL_TYPE_GIVEITEM Then
        n = FindOpenInvSlot(index, Spell(SpellNum).Data1)
        
        If n > 0 Then
            Call GiveItem(index, Spell(SpellNum).Data1, Spell(SpellNum).Data2)
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & ".", BrightBlue)
            
            ' Take away the mana points
            Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
            Call SendMP(index)
            Casted = True
        Else
            Call PlayerMsg(index, "Your inventory is full!", BrightRed)
        End If
        
        Exit Sub
    End If
        
    n = Player(index).Target
    
    If Player(index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
            If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
'                If GetPlayerLevel(n) + 5 >= GetPlayerLevel(Index) Then
'                    If GetPlayerLevel(n) - 5 <= GetPlayerLevel(Index) Then
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                
                        Select Case Spell(SpellNum).Type
                            Case SPELL_TYPE_SUBHP
                        
                                Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(n)
                                If Damage > 0 Then
                                    Call AttackPlayer(index, n, Damage)
                                Else
                                    Call PlayerMsg(index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BrightRed)
                                End If
                    
                            Case SPELL_TYPE_SUBMP
                                Call SetPlayerMP(n, GetPlayerMP(n) - Spell(SpellNum).Data1)
                                Call SendMP(n)
                
                            Case SPELL_TYPE_SUBSP
                                Call SetPlayerSP(n, GetPlayerSP(n) - Spell(SpellNum).Data1)
                                Call SendSP(n)
                        End Select
'                    Else
'                        Call PlayerMsg(Index, GetPlayerName(n) & " is far to powerful to even consider attacking.", BrightBlue)
'                    End If
'                Else
'                    Call PlayerMsg(Index, GetPlayerName(n) & " is to weak to even bother with.", BrightBlue)
'                End If
            
                ' Take away the mana points
                Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
                Call SendMP(index)
                Casted = True
            Else
                If GetPlayerMap(index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then
                    Select Case Spell(SpellNum).Type
                    
                        Case SPELL_TYPE_ADDHP
                            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                            Call SendHP(n)
                                    
                        Case SPELL_TYPE_ADDMP
                            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                            Call SendMP(n)
                    
                        Case SPELL_TYPE_ADDSP
                            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                            Call SendMP(n)
                    End Select
                    
                    ' Take away the mana points
                    Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
                    Call SendMP(index)
                    Casted = True
                Else
                    Call PlayerMsg(index, "Could not cast spell!", BrightRed)
                End If
            End If
        Else
            Call PlayerMsg(index, "Could not cast spell!", BrightRed)
        End If
    Else
        If Npc(MapNpc(GetPlayerMap(index), n).Num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(index), n).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on a " & Trim(Npc(MapNpc(GetPlayerMap(index), n).Num).Name) & ".", BrightBlue)
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_ADDHP
                    MapNpc(GetPlayerMap(index), n).HP = MapNpc(GetPlayerMap(index), n).HP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBHP
                    
                    Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(index), n).Num).DEF / 2)
                    If Damage > 0 Then
                        Call AttackNpc(index, n, Damage)
                    Else
                        Call PlayerMsg(index, "The spell was to weak to hurt " & Trim(Npc(MapNpc(GetPlayerMap(index), n).Num).Name) & "!", BrightRed)
                    End If
                    
                Case SPELL_TYPE_ADDMP
                    MapNpc(GetPlayerMap(index), n).MP = MapNpc(GetPlayerMap(index), n).MP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBMP
                    MapNpc(GetPlayerMap(index), n).MP = MapNpc(GetPlayerMap(index), n).MP - Spell(SpellNum).Data1
            
                Case SPELL_TYPE_ADDSP
                    MapNpc(GetPlayerMap(index), n).SP = MapNpc(GetPlayerMap(index), n).SP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBSP
                    MapNpc(GetPlayerMap(index), n).SP = MapNpc(GetPlayerMap(index), n).SP - Spell(SpellNum).Data1
            End Select
        
            ' Take away the mana points
            Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
            Call SendMP(index)
            Casted = True
        Else
            Call PlayerMsg(index, "Could not cast spell!", BrightRed)
        End If
    End If

    If Casted = True Then
        Player(index).AttackTimer = GetTickCount
        Player(index).CastedSpell = YES
    End If
End Sub

Function GetSpellReqLevel(ByVal index As Long, ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).Data1 - Int(GetClassMAGI(GetPlayerClass(index)) / 4)
End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
Dim i As Long, n As Long

    CanPlayerCriticalHit = False
    
    If GetPlayerWeaponSlot(index) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerSTR(index) / 2) + Int(GetPlayerLevel(index) / 2)
    
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
Dim i As Long, n As Long, ShieldSlot As Long

    CanPlayerBlockHit = False
    
    ShieldSlot = GetPlayerShieldSlot(index)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerDEF(index) / 2) + Int(GetPlayerLevel(index) / 2)
        
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Sub CheckEquippedItems(ByVal index As Long)
Dim Slot As Long, ItemNum As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    Slot = GetPlayerWeaponSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
                Call SetPlayerWeaponSlot(index, 0)
            End If
        Else
            Call SetPlayerWeaponSlot(index, 0)
        End If
    End If

    Slot = GetPlayerArmorSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then
                Call SetPlayerArmorSlot(index, 0)
            End If
        Else
            Call SetPlayerArmorSlot(index, 0)
        End If
    End If

    Slot = GetPlayerHelmetSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then
                Call SetPlayerHelmetSlot(index, 0)
            End If
        Else
            Call SetPlayerHelmetSlot(index, 0)
        End If
    End If

    Slot = GetPlayerShieldSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then
                Call SetPlayerShieldSlot(index, 0)
            End If
        Else
            Call SetPlayerShieldSlot(index, 0)
        End If
    End If
End Sub

Sub UsersOnline_Start(ByVal index As Long)
Dim ls As ListItem

    On Error Resume Next

    If frmServer.lvwInfo.ListItems.Count > 0 And IsPlaying(index) = True Then
        frmServer.lvwInfo.ListItems.Remove index
    End If
    Set ls = frmServer.lvwInfo.ListItems.Add(index, , index)

    If IsPlaying(index) = False Then
        ls.SubItems(1) = ""
        ls.SubItems(2) = ""
        ls.SubItems(3) = ""
        ls.SubItems(4) = ""
        ls.SubItems(5) = ""
    Else
        ls.SubItems(1) = GetPlayerLogin(index)
        ls.SubItems(2) = GetPlayerName(index)
        ls.SubItems(3) = GetPlayerLevel(index)
        ls.SubItems(4) = GetPlayerSprite(index)
        ls.SubItems(5) = GetPlayerAccess(index)
    End If
End Sub
