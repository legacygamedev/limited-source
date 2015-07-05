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
    GetPlayerDamage = GetPlayerDamage + Int(Rnd * 2) - 1
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
    
    If GetPlayerWeaponSlot(index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(index)
        
        GetPlayerDamage = GetPlayerDamage + item(GetPlayerInvItemNum(index, WeaponSlot)).Data2
        
        If GetPlayerInvItemDur(index, WeaponSlot) > -1 Then
            Call SetPlayerInvItemDur(index, WeaponSlot, GetPlayerInvItemDur(index, WeaponSlot) - 1)
        
            If GetPlayerInvItemDur(index, WeaponSlot) = 0 Then
                Call BattleMsg(index, "Your " & Trim(item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, WeaponSlot), 0)
            Else
                If GetPlayerInvItemDur(index, WeaponSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, WeaponSlot) & "/" & Trim(item(GetPlayerInvItemNum(index, WeaponSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If GetPlayerDamage < 0 Then
        GetPlayerDamage = 0
    End If
End Function

Function GetPlayerProtection(ByVal index As Long) As Long
Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ArmorSlot = GetPlayerArmorSlot(index)
    HelmSlot = GetPlayerHelmetSlot(index)
    ShieldSlot = GetPlayerShieldSlot(index)
    GetPlayerProtection = Int(GetPlayerDEF(index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(GetPlayerInvItemNum(index, ArmorSlot)).Data2
        If GetPlayerInvItemDur(index, ArmorSlot) > -1 Then
            Call SetPlayerInvItemDur(index, ArmorSlot, GetPlayerInvItemDur(index, ArmorSlot) - 1)
        
            If GetPlayerInvItemDur(index, ArmorSlot) = 0 Then
                Call BattleMsg(index, "Your " & Trim(item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, ArmorSlot), 0)
            Else
                If GetPlayerInvItemDur(index, ArmorSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, ArmorSlot) & "/" & Trim(item(GetPlayerInvItemNum(index, ArmorSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(GetPlayerInvItemNum(index, HelmSlot)).Data2
        If GetPlayerInvItemDur(index, HelmSlot) > -1 Then
            Call SetPlayerInvItemDur(index, HelmSlot, GetPlayerInvItemDur(index, HelmSlot) - 1)

            If GetPlayerInvItemDur(index, HelmSlot) <= 0 Then
                Call BattleMsg(index, "Your " & Trim(item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, HelmSlot), 0)
            Else
                If GetPlayerInvItemDur(index, HelmSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, HelmSlot) & "/" & Trim(item(GetPlayerInvItemNum(index, HelmSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(GetPlayerInvItemNum(index, ShieldSlot)).Data2
        If GetPlayerInvItemDur(index, ShieldSlot) > -1 Then
            Call SetPlayerInvItemDur(index, ShieldSlot, GetPlayerInvItemDur(index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(index, ShieldSlot) <= 0 Then
                Call BattleMsg(index, "Your " & Trim(item(GetPlayerInvItemNum(index, ShieldSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, ShieldSlot), 0)
            Else
                If GetPlayerInvItemDur(index, ShieldSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(item(GetPlayerInvItemNum(index, ShieldSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, ShieldSlot) & "/" & Trim(item(GetPlayerInvItemNum(index, ShieldSlot)).Data1), Yellow, 0)
                End If
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
    
    If item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
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
        If MapItem(MapNum, i).num = 0 Then
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
            If item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
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
            If item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case item(GetPlayerInvItemNum(index, i)).Type
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

                
                n = item(GetPlayerInvItemNum(index, i)).Type
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
        
        If (item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (item(ItemNum).Type = ITEM_TYPE_HELMET) Or (item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(index, i, item(ItemNum).Data1)
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
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, item(ItemNum).Data1, MapNum, x, y)
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
        MapItem(MapNum, i).num = ItemNum
        MapItem(MapNum, i).value = ItemVal
        
        If ItemNum <> 0 Then
            If (item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
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
                If item(Map(MapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).Tile(x, y).Data2 <= 0 Then
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
        If (MapItem(MapNum, i).num > 0) And (MapItem(MapNum, i).num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).x = GetPlayerX(index)) And (MapItem(MapNum, i).y = GetPlayerY(index)) Then
                ' Find open slot
                n = FindOpenInvSlot(index, MapItem(MapNum, i).num)
                
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(index, n, MapItem(MapNum, i).num)
                    If item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
                        Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).value)
                        Msg = "You picked up " & MapItem(MapNum, i).value & " " & Trim(item(GetPlayerInvItemNum(index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(index, n, 0)
                        Msg = "You picked up a " & Trim(item(GetPlayerInvItemNum(index, n)).Name) & "."
                    End If
                    Call SetPlayerInvItemDur(index, n, MapItem(MapNum, i).Dur)
                        
                    ' Erase item from the map
                    MapItem(MapNum, i).num = 0
                    MapItem(MapNum, i).value = 0
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

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal amount As Long)
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
            Select Case item(GetPlayerInvItemNum(index, InvNum)).Type
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
                                
            MapItem(GetPlayerMap(index), i).num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
            MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                        
            If item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If amount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).value = GetPlayerInvItemValue(index, InvNum)
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim(item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), i).value = amount
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & amount & " " & Trim(item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - amount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(index), i).value = 0
                If item(GetPlayerInvItemNum(index, InvNum)).Type >= ITEM_TYPE_WEAPON And item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                    If item(GetPlayerInvItemNum(index, InvNum)).Data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(item(GetPlayerInvItemNum(index, InvNum)).Name) & " - Ind.", Yellow)
                    Else
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(item(GetPlayerInvItemNum(index, InvNum)).Name) & " - " & GetPlayerInvItemDur(index, InvNum) & "/" & item(GetPlayerInvItemNum(index, InvNum)).Data1 & ".", Yellow)
                    End If
                Else
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).num, amount, MapItem(GetPlayerMap(index), i).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsg(index, "To many items already on the ground.", BrightRed)
        End If
    End If
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
        If GameTime = TIME_NIGHT Then
            If Npc(NpcNum).SpawnTime = 1 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        Else
            If Npc(NpcNum).SpawnTime = 2 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
    
        MapNpc(MapNum, MapNpcNum).num = NpcNum
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxHP(NpcNum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)
                
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
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        End If
    End If
    
    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)
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

Function CanAttackPlayer(ByVal attacker As Long, ByVal Victim As Long) As Boolean
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(attacker) > 0 Then
        AttackSpeed = item(GetPlayerInvItemNum(attacker, GetPlayerWeaponSlot(attacker))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If
    CanAttackPlayer = False
    
    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(Victim) = False Then
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
    If (GetPlayerMap(attacker) = GetPlayerMap(Victim)) And (GetTickCount > Player(attacker).AttackTimer + AttackSpeed) Then
        
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
                If (GetPlayerY(Victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(Victim) = GetPlayerX(attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure that they dont have access
                        If GetPlayerAccess(attacker) > ADMIN_MONITER Then
                            Call PlayerMsg(attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                        Else
                            ' Check to make sure the victim isn't an admin
                            If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                                Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                            Else
                                ' Check if map is attackable
                                If Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                    ' Make sure they are high enough level
                                    If GetPlayerLevel(attacker) < GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") Then
                                        Call PlayerMsg(attacker, "You are below level" & GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") & ", you cannot attack another player yet!", BrightRed)
                                    Else
                                        If GetPlayerLevel(Victim) < GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") Then
                                            Call PlayerMsg(attacker, GetPlayerName(Victim) & " is below level" & GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") & ", you cannot attack this player yet!", BrightRed)
                                        Else
                                            If Trim(GetPlayerGuild(attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                                If Trim(GetPlayerGuild(attacker)) <> Trim(GetPlayerGuild(Victim)) Then
                                                    CanAttackPlayer = True
                                                Else
                                                    Call PlayerMsg(attacker, "You cant attack a guild member!", BrightRed)
                                                End If
                                            Else
                                                CanAttackPlayer = True
                                            End If
                                        End If
                                    End If
                                Else
                                    Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
                                End If
                            End If
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If

            Case DIR_DOWN
                If (GetPlayerY(Victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(Victim) = GetPlayerX(attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure that they dont have access
                        If GetPlayerAccess(attacker) > ADMIN_MONITER Then
                            Call PlayerMsg(attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                        Else
                            ' Check to make sure the victim isn't an admin
                            If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                                Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                            Else
                                ' Check if map is attackable
                                If Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                    ' Make sure they are high enough level
                                    If GetPlayerLevel(attacker) < GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") Then
                                        Call PlayerMsg(attacker, "You are below level" & GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") & ", you cannot attack another player yet!", BrightRed)
                                    Else
                                        If GetPlayerLevel(Victim) < GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") Then
                                            Call PlayerMsg(attacker, GetPlayerName(Victim) & " is below level" & GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") & ", you cannot attack this player yet!", BrightRed)
                                        Else
                                            If Trim(GetPlayerGuild(attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                                If Trim(GetPlayerGuild(attacker)) <> Trim(GetPlayerGuild(Victim)) Then
                                                    CanAttackPlayer = True
                                                Else
                                                    Call PlayerMsg(attacker, "You cant attack a guild member!", BrightRed)
                                                End If
                                            Else
                                                CanAttackPlayer = True
                                            End If
                                        End If
                                    End If
                                Else
                                    Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
                                End If
                            End If
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
        
            Case DIR_LEFT
                If (GetPlayerY(Victim) = GetPlayerY(attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure that they dont have access
                        If GetPlayerAccess(attacker) > ADMIN_MONITER Then
                            Call PlayerMsg(attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                        Else
                            ' Check to make sure the victim isn't an admin
                            If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                                Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                            Else
                                ' Check if map is attackable
                                If Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                    ' Make sure they are high enough level
                                    If GetPlayerLevel(attacker) < GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") Then
                                        Call PlayerMsg(attacker, "You are below level" & GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") & ", you cannot attack another player yet!", BrightRed)
                                    Else
                                        If GetPlayerLevel(Victim) < GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") Then
                                            Call PlayerMsg(attacker, GetPlayerName(Victim) & " is below level" & GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") & ", you cannot attack this player yet!", BrightRed)
                                        Else
                                            If Trim(GetPlayerGuild(attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                                If Trim(GetPlayerGuild(attacker)) <> Trim(GetPlayerGuild(Victim)) Then
                                                    CanAttackPlayer = True
                                                Else
                                                    Call PlayerMsg(attacker, "You cant attack a guild member!", BrightRed)
                                                End If
                                            Else
                                                CanAttackPlayer = True
                                            End If
                                        End If
                                    End If
                                Else
                                    Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
                                End If
                            End If
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
            
            Case DIR_RIGHT
                If (GetPlayerY(Victim) = GetPlayerY(attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure that they dont have access
                        If GetPlayerAccess(attacker) > ADMIN_MONITER Then
                            Call PlayerMsg(attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                        Else
                            ' Check to make sure the victim isn't an admin
                            If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                                Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                            Else
                                ' Check if map is attackable
                                If Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                    ' Make sure they are high enough level
                                    If GetPlayerLevel(attacker) < GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") Then
                                        Call PlayerMsg(attacker, "You are below level" & GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") & ", you cannot attack another player yet!", BrightRed)
                                    Else
                                        If GetPlayerLevel(Victim) < GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") Then
                                            Call PlayerMsg(attacker, GetPlayerName(Victim) & " is below level" & GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") & ", you cannot attack this player yet!", BrightRed)
                                        Else
                                            If Trim(GetPlayerGuild(attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                                If Trim(GetPlayerGuild(attacker)) <> Trim(GetPlayerGuild(Victim)) Then
                                                    CanAttackPlayer = True
                                                Else
                                                    Call PlayerMsg(attacker, "You cant attack a guild member!", BrightRed)
                                                End If
                                            Else
                                                CanAttackPlayer = True
                                            End If
                                        End If
                                    End If
                                Else
                                    Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
                                End If
                            End If
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
        End Select
    End If
End Function

Function CanAttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(attacker) > 0 Then
        AttackSpeed = item(GetPlayerInvItemNum(attacker, GetPlayerWeaponSlot(attacker))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

CanAttackNpc = False
 
' Check for subscript out of range
If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
Exit Function
End If
 
' Check for subscript out of range
If MapNpc(GetPlayerMap(attacker), MapNpcNum).num <= 0 Then
Exit Function
End If
 
MapNum = GetPlayerMap(attacker)
NpcNum = MapNpc(MapNum, MapNpcNum).num
 
' Make sure the npc isn't already dead
If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
Exit Function
End If
 
' Make sure they are on the same map
If IsPlaying(attacker) Then
    If NpcNum > 0 And GetTickCount > Player(attacker).AttackTimer + AttackSpeed Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
                If (MapNpc(MapNum, MapNpcNum).y + 1 = GetPlayerY(attacker)) And (MapNpc(MapNum, MapNpcNum).x = GetPlayerX(attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
 
            Case DIR_DOWN
                If (MapNpc(MapNum, MapNpcNum).y - 1 = GetPlayerY(attacker)) And (MapNpc(MapNum, MapNpcNum).x = GetPlayerX(attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
 
            Case DIR_LEFT
                If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(attacker)) And (MapNpc(MapNum, MapNpcNum).x + 1 = GetPlayerX(attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
 
            Case DIR_RIGHT
                If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(attacker)) And (MapNpc(MapNum, MapNpcNum).x - 1 = GetPlayerX(attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
        End Select
    End If
End If
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
    
    CanNpcAttackPlayer = False
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(index) = False Then
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

Sub AttackPlayer(ByVal attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Exp As Long
Dim n As Long
Dim i As Long
Dim J As Byte

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for weapon
    If GetPlayerWeaponSlot(attacker) > 0 Then
        n = GetPlayerInvItemNum(attacker, GetPlayerWeaponSlot(attacker))
    Else
        n = 0
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(attacker, GetPlayerMap(attacker), "ATTACK" & SEP_CHAR & attacker & SEP_CHAR & END_CHAR)

If Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA Then
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        Call BattleMsg(attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
        Call BattleMsg(Victim, GetPlayerName(attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)
    
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(attacker), BrightRed)
        
        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
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
            attacker = GetVar(App.Path & "\Data.ini", "ADDED", "ExpToAttacker")
            Exp = Int(GetPlayerExp(Victim) / attacker)
            
            ' Make sure we dont get less then 0
            If Exp < 0 Then
                Exp = 0
            End If
            
            If GetPlayerLevel(Victim) = MAX_LEVEL Then
                Call BattleMsg(Victim, "You cant lose any experience!", BrightRed, 1)
                Call BattleMsg(attacker, GetPlayerName(Victim) & " is the max level!", BrightBlue, 0)
            Else
                If Exp = 0 Then
                    Call BattleMsg(Victim, "You lost no experience.", BrightRed, 1)
                    Call BattleMsg(attacker, "You received no experience.", BrightBlue, 0)
                Else
                    Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                    Call BattleMsg(Victim, "You lost " & Exp & " experience.", BrightRed, 1)
                    Call SetPlayerExp(attacker, GetPlayerExp(attacker) + Exp)
                    Call BattleMsg(attacker, "You got " & Exp & " experience for killing " & GetPlayerName(Victim) & ".", BrightBlue, 0)
                End If
            End If
        End If
        
        ' Warp player away
    Call PlayerWarp(Victim, GetVar("Classes\Class" & GetPlayerClass(Victim) & ".ini", "DEATHLOCATION", "Map"), GetVar("Classes\Class" & GetPlayerClass(Victim) & ".ini", "DEATHLOCATION", "X"), GetVar("Classes\Class" & GetPlayerClass(Victim) & ".ini", "DEATHLOCATION", "Y"))
        
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
                
        ' Check for a level up
        Call CheckPlayerLevelUp(attacker)
        
        ' Check if target is player who died and if so set target to 0
        If Player(attacker).TargetType = TARGET_TYPE_PLAYER And Player(attacker).Target = Victim Then
            Player(attacker).Target = 0
            Player(attacker).TargetType = 0
        End If
        
        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(attacker) = NO Then
                Call SetPlayerPK(attacker, YES)
                Call SendPlayerData(attacker)
                Call GlobalMsg(GetPlayerName(attacker) & " has been deemed a Player Killer!", BrightRed)
            End If
        Else
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!", BrightRed)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Check for a weapon and say damage
        Call BattleMsg(attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
        Call BattleMsg(Victim, GetPlayerName(attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)
        If n = 0 Then
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
    End If
ElseIf Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker), GetPlayerY(attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        Call BattleMsg(attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
        Call BattleMsg(Victim, GetPlayerName(attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)
        If n = 0 Then
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
    
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed in the arena by " & GetPlayerName(attacker), BrightRed)
            
        ' Warp player away
        Call PlayerWarp(Victim, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data1, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data2, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data3)
        
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
                        
        ' Check if target is player who died and if so set target to 0
        If Player(attacker).TargetType = TARGET_TYPE_PLAYER And Player(attacker).Target = Victim Then
            Player(attacker).Target = 0
            Player(attacker).TargetType = 0
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Check for a weapon and say damage
        Call BattleMsg(attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
        Call BattleMsg(Victim, GetPlayerName(attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)
        If n = 0 Then
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
    End If
End If
    
    ' Reset timer for attacking
    Player(attacker).AttackTimer = GetTickCount
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & END_CHAR)
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
    If MapNpc(GetPlayerMap(Victim), MapNpcNum).num <= 0 Then
        Exit Sub
    End If
        
    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Victim)
    
    ':: AUTO TURN ::
    'If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "AutoTurn")) = 1 Then
        'If GetPlayerX(Victim) - 1 = MapNpc(MapNum, MapNpcNum).X Then
            'Call SetPlayerDir(Victim, DIR_LEFT)
        'End If
        'If GetPlayerX(Victim) + 1 = MapNpc(MapNum, MapNpcNum).X Then
            'Call SetPlayerDir(Victim, DIR_RIGHT)
        'End If
        'If GetPlayerY(Victim) - 1 = MapNpc(MapNum, MapNpcNum).Y Then
            'Call SetPlayerDir(Victim, DIR_UP)
        'End If
        'If GetPlayerY(Victim) + 1 = MapNpc(MapNum, MapNpcNum).Y Then
            'Call SetPlayerDir(Victim, DIR_DOWN)
        'End If
        'Call SendDataToMap(GetPlayerMap(Victim), "changedir" & SEP_CHAR & GetPlayerDir(Victim) & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    'End If
    ':: END AUTO TURN ::
    
    Name = Trim(Npc(MapNpc(MapNum, MapNpcNum).num).Name)
    
    If Damage >= GetPlayerHP(Victim) Then
        ' Say damage
        Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)
        
        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by a " & Name, BrightRed)
        
        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
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
            
            ' Calculate exp to give attacke
            Exp = Int(GetPlayerExp(Victim) / GetVar(App.Path & "\Data.ini", "ADDED", "ExpLostToEnemy"))
            
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
Call PlayerWarp(Victim, GetVar("Classes\Class" & GetPlayerClass(Victim) & ".ini", "DEATHLOCATION", "Map"), GetVar("Classes\Class" & GetPlayerClass(Victim) & ".ini", "DEATHLOCATION", "X"), GetVar("Classes\Class" & GetPlayerClass(Victim) & ".ini", "DEATHLOCATION", "Y"))
        
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
        Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)
        
        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
    End If
    
    Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & END_CHAR)
End Sub

Sub AttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim n As Long, i As Long, q As Integer, x As Long
Dim STR As Long, DEF As Long, MapNum As Long, NpcNum As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for weapon
    If GetPlayerWeaponSlot(attacker) > 0 Then
        n = GetPlayerInvItemNum(attacker, GetPlayerWeaponSlot(attacker))
    Else
        n = 0
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(attacker, GetPlayerMap(attacker), "ATTACK" & SEP_CHAR & attacker & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim(Npc(NpcNum).Name)
        
    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        
        Call BattleMsg(attacker, "You killed a " & Name, BrightRed, 0)

        Dim add As String

        add = 0
        If GetPlayerWeaponSlot(attacker) > 0 Then
            add = add + item(GetPlayerInvItemNum(attacker, GetPlayerWeaponSlot(attacker))).AddEXP
        End If
        If GetPlayerArmorSlot(attacker) > 0 Then
            add = add + item(GetPlayerInvItemNum(attacker, GetPlayerArmorSlot(attacker))).AddEXP
        End If
        If GetPlayerShieldSlot(attacker) > 0 Then
            add = add + item(GetPlayerInvItemNum(attacker, GetPlayerShieldSlot(attacker))).AddEXP
        End If
        If GetPlayerHelmetSlot(attacker) > 0 Then
            add = add + item(GetPlayerInvItemNum(attacker, GetPlayerHelmetSlot(attacker))).AddEXP
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
        If Player(attacker).InParty = False Or Player(attacker).Party.ShareExp = False Then
            If GetPlayerLevel(attacker) = MAX_LEVEL Then
                Call SetPlayerExp(attacker, Experience(MAX_LEVEL))
                Call BattleMsg(attacker, "You cant gain anymore experience!", BrightBlue, 0)
            Else
                Call SetPlayerExp(attacker, GetPlayerExp(attacker) + Exp)
                Call BattleMsg(attacker, "You have gained " & Exp & " experience.", BrightBlue, 0)
            End If
        Else
            q = 0
            'The following code will tell us how many party members we have active
            For x = 1 To MAX_PARTY_MEMBERS
            If Player(attacker).Party.Member(x) > 0 Then q = q + 1
            Next x
            'MsgBox "in party" & q
            If q = 2 Then 'Remember, if they aren't in a party they'll only get one person, so this has to be at least two
                Exp = Exp * 0.75 ' 3/4 experience
                'MsgBox Exp
                For x = 1 To MAX_PARTY_MEMBERS
                    If Player(attacker).Party.Member(x) > 0 Then
                        If Player(Player(attacker).Party.Member(x)).Party.ShareExp = True Then
                            If GetPlayerLevel(Player(attacker).Party.Member(x)) = MAX_LEVEL Then
                                Call SetPlayerExp(Player(attacker).Party.Member(x), Experience(MAX_LEVEL))
                                Call BattleMsg(Player(attacker).Party.Member(x), "You cant gain anymore experience!", BrightBlue, 0)
                            Else
                                Call SetPlayerExp(Player(attacker).Party.Member(x), GetPlayerExp(Player(attacker).Party.Member(x)) + Exp)
                                Call BattleMsg(Player(attacker).Party.Member(x), "You have gained " & Exp & " party experience.", BrightBlue, 0)
                            End If
                        End If
                    End If
                Next x
            Else 'if there are 3 or more party members..
                Exp = Exp * 0.5  ' 1/2 experience
                    For x = 1 To MAX_PARTY_MEMBERS
                        If Player(attacker).Party.Member(x) > 0 Then
                            If Player(Player(attacker).Party.Member(x)).Party.ShareExp = True Then
                                If GetPlayerLevel(Player(attacker).Party.Member(x)) = MAX_LEVEL Then
                                    Call SetPlayerExp(Player(attacker).Party.Member(x), Experience(MAX_LEVEL))
                                    Call BattleMsg(Player(attacker).Party.Member(x), "You cant gain anymore experience!", BrightBlue, 0)
                                Else
                                    Call SetPlayerExp(Player(attacker).Party.Member(x), GetPlayerExp(n) + Exp)
                                    Call BattleMsg(Player(attacker).Party.Member(x), "You have gained " & Exp & " party experience.", BrightBlue, 0)
                                End If
                            End If
                        End If
                    Next x
            End If
        End If
                                
        For i = 1 To MAX_NPC_DROPS
            ' Drop the goods if they get it
            n = Int(Rnd * Npc(NpcNum).ItemNPC(i).Chance) + 1
            If n = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
            End If
        Next i
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
        
        ' Check for level up
        Call CheckPlayerLevelUp(attacker)
       
        ' Check for level up party member
        If Player(attacker).InParty = YES Then
            For x = 1 To MAX_PARTY_MEMBERS
                Call CheckPlayerLevelUp(Player(attacker).Party.Member(x))
            Next x
        End If
        
        ' Check for level up party member
        If Player(attacker).InParty = YES Then
            Call CheckPlayerLevelUp(Player(attacker).PartyPlayer)
        End If
    
        ' Check if target is npc that died and if so set target to 0
        If Player(attacker).TargetType = TARGET_TYPE_NPC And Player(attacker).Target = MapNpcNum Then
            Player(attacker).Target = 0
            Player(attacker).TargetType = 0
        End If
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum, MapNpcNum).HP = MapNpc(MapNum, MapNpcNum).HP - Damage
        
        ' Check for a weapon and say damage
        Call BattleMsg(attacker, "You hit a " & Name & " for " & Damage & " damage.", White, 0)
        
        If n = 0 Then
            'Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        Else
            'Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If
        
        ' Check if we should send a message
        If MapNpc(MapNum, MapNpcNum).Target = 0 And MapNpc(MapNum, MapNpcNum).Target <> attacker Then
            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                Call PlayerMsg(attacker, "A " & Trim(Npc(NpcNum).Name) & " : " & Trim(Npc(NpcNum).AttackSay) & "", SayColor)
            End If
        End If
        
        ' Set the NPC target to the player
        MapNpc(MapNum, MapNpcNum).Target = attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum, i).num = MapNpc(MapNum, MapNpcNum).num Then
                    MapNpc(MapNum, i).Target = attacker
                End If
            Next i
        End If
    End If
        
    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)
        
    ' Reset attack timer
    Player(attacker).AttackTimer = GetTickCount
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim Packet As String
Dim OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    'If Trim(Shop(ShopNum).LeaveSay) <> "" Then
        'Call PlayerMsg(Index, Trim(Shop(ShopNum).Name) & " : " & Trim(Shop(ShopNum).LeaveSay) & "", SayColor)
    'End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    Call SendLeaveMap(index, OldMap)
    
    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
                
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Player(index).GettingMap = YES
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "warp" & SEP_CHAR & END_CHAR)
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)
    
    Call SendInventory(index)
    Call SendWornEquipment(index)
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim MapNum As Long
Dim x As Long
Dim y As Long
Dim i As Long
Dim Moved As Byte

    ' They tried to hack
    'If Moved = NO Then
        'Call HackingAttempt(index, "Position Modification")
        'Exit Sub
    'End If
    
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
                    If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
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
                    If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
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
                    If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
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
                    If (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
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
    
    If GetPlayerX(index) < 0 Or GetPlayerY(index) < 0 Or GetPlayerX(index) > MAX_MAPX Or GetPlayerY(index) > MAX_MAPY Or GetPlayerMap(index) <= 0 Then
        Call HackingAttempt(index, "")
        Exit Sub
    End If
    
    'healing tiles code
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SendHP(index)
        Call PlayerMsg(index, "You feel a sudden rush through your body as you regain strength!", BrightGreen)
    End If
    
    'Check for kill tile, and if so kill them
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KILL Then
        Call SetPlayerHP(index, 0)
        Call PlayerMsg(index, "You embrace the cold finger of death; and feel your life extinguished", BrightRed)
        ' Warp player away
Call PlayerWarp(index, GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "DEATHLOCATION", "Map"), GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "DEATHLOCATION", "X"), GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "DEATHLOCATION", "Y"))
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SetPlayerSP(index, GetPlayerMaxSP(index))
        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
        Moved = YES
    End If

    If GetPlayerX(index) + 1 <= MAX_MAPX Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
            
            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerX(index) - 1 >= 0 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
            
            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(index) - 1 >= 0 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
            
            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(index) + 1 <= MAX_MAPY Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
            
            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
            
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
                            
            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            If Trim(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) = "" Then
                Call MapMsg(GetPlayerMap(index), "A door has been unlocked!", White)
            Else
                Call MapMsg(GetPlayerMap(index), Trim(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), White)
            End If
            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
        End If
    End If
        
    ' Check for shop
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SHOP Then
       If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 > 0 Then
            Call SendTrade(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
        Else
            Call PlayerMsg(index, "There is no shop here.", BrightRed)
        End If
    End If
        
    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(index) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            Call PlayerMsg(index, "You already have this sprite!", BrightRed)
            Exit Sub
        Else
            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
                Call SendDataTo(index, "spritechange" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Else
                If item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(index, "This sprite will cost you " & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3 & " " & Trim(item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", Yellow)
                Else
                    Call PlayerMsg(index, "This sprite will cost you a " & Trim(item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", Yellow)
                End If
                Call SendDataTo(index, "spritechange" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    
    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 > -1 Then
            If GetPlayerClass(index) <> Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
                Call PlayerMsg(index, "You arent the required class!", BrightRed)
                Exit Sub
            End If
        End If
        
        If GetPlayerClass(index) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            Call PlayerMsg(index, "You are already this class!", BrightRed)
        Else
            If Player(index).Char(Player(index).CharNum).Sex = 0 Then
                If GetPlayerSprite(index) = Class(GetPlayerClass(index)).MaleSprite Then
                    Call SetPlayerSprite(index, Class(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).MaleSprite)
                End If
            Else
                If GetPlayerSprite(index) = Class(GetPlayerClass(index)).FemaleSprite Then
                    Call SetPlayerSprite(index, Class(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).FemaleSprite)
                End If
            End If
            
            Call SetPlayerSTR(index, (Player(index).Char(Player(index).CharNum).STR - Class(GetPlayerClass(index)).STR))
            Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).DEF - Class(GetPlayerClass(index)).DEF))
            Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).Magi - Class(GetPlayerClass(index)).Magi))
            Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed - Class(GetPlayerClass(index)).Speed))
            
            Call SetPlayerClass(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)

            Call SetPlayerSTR(index, (Player(index).Char(Player(index).CharNum).STR + Class(GetPlayerClass(index)).STR))
            Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).DEF + Class(GetPlayerClass(index)).DEF))
            Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).Magi + Class(GetPlayerClass(index)).Magi))
            Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed + Class(GetPlayerClass(index)).Speed))
            
            
            Call PlayerMsg(index, "Your new class is a " & Trim(Class(GetPlayerClass(index)).Name) & "!", BrightGreen)
            
            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
            Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
        End If
    End If
    
    ' Check if player stepped on notice tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_NOTICE Then
        If Trim(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) <> "" Then
            Call PlayerMsg(index, Trim(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), Black)
        End If
        If Trim(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2) <> "" Then
            Call PlayerMsg(index, Trim(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2), Grey)
        End If
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String3 & SEP_CHAR & END_CHAR)
    End If
    
    ' Check if player stepped on sound tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1 & SEP_CHAR & END_CHAR)
    End If
    
    ' Check for Chest tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_CHEST Then
    Call GiveItem(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2)
    Call PlayerMsg(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1, Black)
    End If
    
    ' Check for Trap tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_TRAP Then
    Call SetPlayerHP(index, GetPlayerHP(index) - Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
    Call SetPlayerMP(index, GetPlayerMP(index) - Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2)
    Call SetPlayerSP(index, GetPlayerSP(index) - Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3)
    'Call SetPlayerExp(index, GetPlayerExp(index) - Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data4)
    'Call SetPlayerLevel(index, GetPlayerLevel(index) - Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data5)
    Call BattleMsg(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1, Red, 0)
    End If

    'Check for Quest tile, and if so run Quest Script
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_QUEST Then
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 1 Then
    Call Questone(index)
    End If
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 2 Then
    Call QuestTwo(index)
    End If
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 3 Then
    Call QuestThree(index)
    End If
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 4 Then
    Call QuestFour(index)
    End If
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 5 Then
    Call QuestFive(index)
    End If
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 6 Then
    Call QuestSix(index)
    End If
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 7 Then
    Call QuestSeven(index)
    End If
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 8 Then
    Call QuestEight(index)
    End If
    End If
    
    ' Check for Scripted Tile
    If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & index & "," & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        End If
    End If
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
Dim i As Long, n As Long
Dim x As Long, y As Long
Dim BX As Long, BY As Long

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
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
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
    
                If CanNPCMoveAttributeNPC(MapNum, MapNpcNum, DIR_UP) = False Then
                    CanNpcMove = False
                    Exit Function
                End If
                                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y - 1) Then
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
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
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
                
                If CanNPCMoveAttributeNPC(MapNum, MapNpcNum, DIR_DOWN) = False Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y + 1) Then
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
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
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
                
                If CanNPCMoveAttributeNPC(MapNum, MapNpcNum, DIR_LEFT) = False Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x - 1) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
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
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
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
                
                If CanNPCMoveAttributeNPC(MapNum, MapNpcNum, DIR_RIGHT) = False Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).num > 0) And (MapNpc(MapNum, i).x = MapNpc(MapNum, MapNpcNum).x + 1) And (MapNpc(MapNum, i).y = MapNpc(MapNum, MapNpcNum).y) Then
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

Sub JoinGame(ByVal index As Long)
Dim MOTD As String
Dim f As Long

    ' Set the flag so we know the person is in the game
    Player(index).InGame = True
    
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(index, "LOGINOK" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
    sssss = 1
    ReDim Player(index).Party.Member(1 To MAX_PARTY_MEMBERS)
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendEmoticons(index)
    Call SendArrows(index)
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
    Call SendOnlineList
    Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", 1)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    Call SendPlayerData(index)
    
    If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "JoinGame " & index
    Else
        MOTD = GetVar("motd.ini", "MOTD", "Msg")
        
        ' Send a global message that he/she joined
        If GetPlayerAccess(index) <= ADMIN_MONITER Then
            Call GlobalMsg(GetPlayerName(index) & " has joined " & GAME_NAME & "!", 7)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has joined " & GAME_NAME & "!", 15)
        End If
    
        ' Send them welcome
        Call PlayerMsg(index, "Welcome to " & GAME_NAME & "!", 15)
        
        ' Send motd
        If Trim(MOTD) <> "" Then
            Call PlayerMsg(index, "MOTD: " & MOTD, 11)
        End If
    End If
    
    ' Send whos online
    Call SendWhosOnline(index)
    Call ShowPLR(index)

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(index, "INGAME" & SEP_CHAR & END_CHAR)
End Sub

Sub LeftGame(ByVal index As Long)
Dim n As Long

    If Player(index).InGame = True Then
        Player(index).InGame = False
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) = 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If

        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        Call RemovePMember(index)
        
        If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "LeftGame " & index
        Else
            ' Check for boot map
            If Map(GetPlayerMap(index)).BootMap > 0 Then
                Call SetPlayerX(index, Map(GetPlayerMap(index)).BootX)
                Call SetPlayerY(index, Map(GetPlayerMap(index)).BootY)
                Call SetPlayerMap(index, Map(GetPlayerMap(index)).BootMap)
            End If
                  
            ' Send a global message that he/she left
            If GetPlayerAccess(index) <= 1 Then
                Call GlobalMsg(GetPlayerName(index) & " has left " & GAME_NAME & "!", 7)
            Else
                Call GlobalMsg(GetPlayerName(index) & " has left " & GAME_NAME & "!", 15)
            End If
        End If
        
        Call SavePlayer(index)
        
        Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & " has disconnected from " & GAME_NAME & ".", True)
        Call SendLeftGame(index)
        Call RemovePLR
        For n = 1 To MAX_PLAYERS
            Call ShowPLR(n)
        Next n
    End If
    
    Call ClearPlayer(index)
    Call SendOnlineList
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
        
    GetNpcMaxMP = Npc(NpcNum).Magi * 2
End Function

Function GetNpcMaxSP(ByVal NpcNum As Long)
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxSP = 0
        Exit Function
    End If
        
    GetNpcMaxSP = Npc(NpcNum).Speed * 2
End Function

Function GetPlayerHPRegen(ByVal index As Long)
Dim i As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
            GetPlayerHPRegen = 0
            Exit Function
        End If
        
        i = GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegenSpeed")
        If i < 1 Then i = 1
        
        'I = Int(GetPlayerDEF(index) / 2)
        'If I < 2 Then I = 2
        
        GetPlayerHPRegen = i
    End If
End Function

Function GetPlayerMPRegen(ByVal index As Long)
Dim i As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
            GetPlayerMPRegen = 0
            Exit Function
        End If
        
        i = GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegenSpeed")
        If i < 1 Then i = 1
        
        'I = Int(GetPlayerMAGI(index) / 2)
        'If I < 2 Then I = 2
        
        GetPlayerMPRegen = i
    End If
End Function

Function GetPlayerSPRegen(ByVal index As Long)
Dim i As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
            GetPlayerSPRegen = 0
            Exit Function
        End If
        
        i = GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegenSpeed")
        If i < 1 Then i = 1
        
        'I = Int(GetPlayerSPEED(index) / 2)
        'If I < 2 Then I = 2
        
        GetPlayerSPRegen = i
    End If
End Function

Function GetNpcHPRegen(ByVal NpcNum As Long)
Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcHPRegen = 0
        Exit Function
    End If
    
    i = GetVar(App.Path & "\Data.ini", "CONFIG", "NPCHPRegen")
    If i < 1 Then i = 1
    
    'I = Int(Npc(NpcNum).DEF / 3)
    'If I < 1 Then I = 1
    
    GetNpcHPRegen = i
End Function

Sub CheckPlayerLevelUp(ByVal index As Long)
Dim i As Long
Dim d As Long
Dim C As Long
    C = 0
    
    If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
        If GetPlayerLevel(index) < MAX_LEVEL Then
            If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerLevelUp " & index
            Else
                Do Until GetPlayerExp(index) < GetPlayerNextLevel(index)
                    DoEvents
                    If GetPlayerLevel(index) < MAX_LEVEL Then
                        If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
                            d = GetPlayerExp(index) - GetPlayerNextLevel(index)
                            Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
                                
                            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 1)
                            Call SetPlayerExp(index, d)
                            C = C + 1
                        End If
                    End If
                Loop
                If C > 1 Then
                    Call GlobalMsg(GetPlayerName(index) & " has gained " & C & " levels!", 6)
                Else
                    Call GlobalMsg(GetPlayerName(index) & " has gained a level!", 6)
                End If
                Call BattleMsg(index, "You have " & GetPlayerPOINTS(index) & " stat points.", 9, 0)
            End If
            Call SendDataToMap(GetPlayerMap(index), "levelup" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        End If
        
        If GetPlayerLevel(index) = MAX_LEVEL Then
            Call SetPlayerExp(index, Experience(MAX_LEVEL))
        End If
    End If
    
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendStats(index)
End Sub

Sub CastSpell(ByVal index As Long, ByVal SpellSlot As Long)
Dim SpellNum As Long, i As Long, n As Long, Damage As Long
Dim Casted As Boolean

    Casted = False
    
    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    SpellNum = GetPlayerSpell(index, SpellSlot)
    
    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then
        Call BattleMsg(index, "You do not have this spell!", BrightRed, 0)
        Exit Sub
    End If
    
    i = GetSpellReqLevel(index, SpellNum)

    ' Check if they have enough MP
    If GetPlayerMP(index) < Spell(SpellNum).MPCost Then
        Call BattleMsg(index, "Not enough mana!", BrightRed, 0)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If i > GetPlayerLevel(index) Then
        Call BattleMsg(index, "You must be level " & i & " to cast this spell.", BrightRed, 0)
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
            Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
            Call SendMP(index)
            Casted = True
        Else
            Call PlayerMsg(index, "Your inventory is full!", BrightRed)
        End If
        
        Exit Sub
    End If
    If Spell(SpellNum).Type = SPELL_TYPE_WARP Then
Dim Moved As Byte
n = Player(index).Target
If Player(index).TargetType = TARGET_TYPE_PLAYER Then
If IsPlaying(n) Then
Call PlayerWarp(n, Spell(SpellNum).Data1, Spell(SpellNum).Data2, Spell(SpellNum).Data3)
Call PlayerMsg(index, "You have teleported " & GetPlayerName(n), BrightBlue)
Call PlayerMsg(n, "You have been teleported by " & GetPlayerName(index), BrightBlue)
Moved = YES
Else
Call PlayerMsg(index, GetPlayerName(n) & " is not online", BrightRed)
End If
Else
Call PlayerMsg(index, "You can't teleport NPC'S", BrightRed)
End If
Casted = True
Exit Sub
End If
Dim x As Long, y As Long

If Spell(SpellNum).AE = 1 Then
    For y = GetPlayerY(index) - Spell(SpellNum).Range To GetPlayerY(index) + Spell(SpellNum).Range
        For x = GetPlayerX(index) - Spell(SpellNum).Range To GetPlayerX(index) + Spell(SpellNum).Range
            n = -1
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) = True Then
                    If GetPlayerMap(index) = GetPlayerMap(i) Then
                        If GetPlayerX(i) = x And GetPlayerY(i) = y Then
                            If i = index Then
                                If Spell(SpellNum).Type = SPELL_TYPE_ADDHP Or Spell(SpellNum).Type = SPELL_TYPE_ADDMP Or Spell(SpellNum).Type = SPELL_TYPE_ADDSP Then
                                    Player(index).Target = i
                                    Player(index).TargetType = TARGET_TYPE_PLAYER
                                    n = Player(index).Target
                                End If
                            Else
                                Player(index).Target = i
                                Player(index).TargetType = TARGET_TYPE_PLAYER
                                n = Player(index).Target
                            End If
                        End If
                    End If
                End If
            Next i
            
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(GetPlayerMap(index), i).num > 0 Then
                    If MapNpc(GetPlayerMap(index), i).x = x And MapNpc(GetPlayerMap(index), i).y = y Then
                        Player(index).Target = i
                        Player(index).TargetType = TARGET_TYPE_NPC
                        n = Player(index).Target
                    End If
                End If
            Next i
                
        Casted = False
        If n > 0 Then
            If Player(index).TargetType = TARGET_TYPE_PLAYER Then
                If IsPlaying(n) Then
                    If n <> index Then
                        Player(index).TargetType = TARGET_TYPE_PLAYER
                        If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") And GetPlayerLevel(n) >= GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
                            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                    
                            Select Case Spell(SpellNum).Type
                                Case SPELL_TYPE_SUBHP
                            
                                    Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(n)
                                    If Damage > 0 Then
                                        Call AttackPlayer(index, n, Damage)
                                    Else
                                        Call BattleMsg(index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BrightRed, 0)
                                    End If
                        
                                Case SPELL_TYPE_SUBMP
                                    Call SetPlayerMP(n, GetPlayerMP(n) - Spell(SpellNum).Data1)
                                    Call SendMP(n)
                    
                                Case SPELL_TYPE_SUBSP
                                    Call SetPlayerSP(n, GetPlayerSP(n) - Spell(SpellNum).Data1)
                                    Call SendSP(n)
                            End Select

                            Casted = True
                        Else
                            If GetPlayerMap(index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then
                                Select Case Spell(SpellNum).Type
                                
                                    Case SPELL_TYPE_ADDHP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                                        Call SendHP(n)
                                                
                                    Case SPELL_TYPE_ADDMP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                                        Call SendMP(n)
                                
                                    Case SPELL_TYPE_ADDSP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                                        Call SendMP(n)
                                End Select
                                
                                Casted = True
                            Else
                                Call PlayerMsg(index, "Could not cast spell!", BrightRed)
                            End If
                        End If
                    Else
                        Player(index).TargetType = TARGET_TYPE_PLAYER
                        If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= GetVar(App.Path & "\Data.ini", "ADDED", "CUREYOURSELF") And GetPlayerLevel(n) >= GetVar(App.Path & "\Data.ini", "ADDED", "CUREOTHER") And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
                        Else
                            If GetPlayerMap(index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then
                                Select Case Spell(SpellNum).Type
                                
                                    Case SPELL_TYPE_ADDHP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                                        Call SendHP(n)
                                                
                                    Case SPELL_TYPE_ADDMP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                                        Call SendMP(n)
                                
                                    Case SPELL_TYPE_ADDSP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                                        Call SendMP(n)
                                End Select

                                Casted = True
                            Else
                                Call BattleMsg(index, "Could not cast spell!", BrightRed, 0)
                            End If
                        End If
                    End If
                Else
                    Call BattleMsg(index, "Could not cast spell!", BrightRed, 0)
                End If
            Else
                Player(index).TargetType = TARGET_TYPE_NPC
                If Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    If Spell(SpellNum).Type >= SPELL_TYPE_SUBHP And Spell(SpellNum).Type <= SPELL_TYPE_SUBSP Then
                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on a " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                        Select Case Spell(SpellNum).Type
                            
                            Case SPELL_TYPE_SUBHP
                                Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(index), n).num).DEF / 2)
                                If Damage > 0 Then
                                    Call AttackNpc(index, n, Damage)
                                Else
                                    Call BattleMsg(index, "The spell was to weak to hurt " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & "!", BrightRed, 0)
                                End If
                            
                            Case SPELL_TYPE_SUBMP
                                MapNpc(GetPlayerMap(index), n).MP = MapNpc(GetPlayerMap(index), n).MP - Spell(SpellNum).Data1

                            Case SPELL_TYPE_SUBSP
                                MapNpc(GetPlayerMap(index), n).SP = MapNpc(GetPlayerMap(index), n).SP - Spell(SpellNum).Data1
                        End Select
                        
                        Casted = True
                    Else
                        Select Case Spell(SpellNum).Type
                            Case SPELL_TYPE_ADDHP
                                'MapNpc(GetPlayerMap(Index), n).HP = MapNpc(GetPlayerMap(Index), n).HP + Spell(SpellNum).Data1
                            
                            Case SPELL_TYPE_ADDMP
                                'MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP + Spell(SpellNum).Data1
                            
                            Case SPELL_TYPE_ADDSP
                                'MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP + Spell(SpellNum).Data1
                        End Select
                        Casted = False
                    End If
                Else
                    Call BattleMsg(index, "Could not cast spell!", BrightRed, 0)
                End If
            End If
        End If
        If Casted = True Then
            Call SendDataToMap(GetPlayerMap(index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & Player(index).Target & SEP_CHAR & END_CHAR)
            'Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
        End If
        Next x
    Next y
    
    Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
    Call SendMP(index)
Else
    n = Player(index).Target
    If Player(index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
            If GetPlayerName(n) <> GetPlayerName(index) Then
                If CInt(Sqr((GetPlayerX(index) - GetPlayerX(n)) ^ 2 + ((GetPlayerY(index) - GetPlayerY(n)) ^ 2))) > Spell(SpellNum).Range Then
                    Call BattleMsg(index, "You are too far away to hit the target.", BrightRed, 0)
                    Exit Sub
                End If
            End If
            Player(index).TargetType = TARGET_TYPE_PLAYER
            If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") And GetPlayerLevel(n) >= GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
        
                Select Case Spell(SpellNum).Type
                    Case SPELL_TYPE_SUBHP
                
                        Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(n)
                        If Damage > 0 Then
                            Call AttackPlayer(index, n, Damage)
                        Else
                            Call BattleMsg(index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BrightRed, 0)
                        End If
            
                    Case SPELL_TYPE_SUBMP
                        Call SetPlayerMP(n, GetPlayerMP(n) - Spell(SpellNum).Data1)
                        Call SendMP(n)
        
                    Case SPELL_TYPE_SUBSP
                        Call SetPlayerSP(n, GetPlayerSP(n) - Spell(SpellNum).Data1)
                        Call SendSP(n)
                End Select
            
                ' Take away the mana points
                Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
                Call SendMP(index)
                Casted = True
            Else
                If GetPlayerMap(index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then
                    Select Case Spell(SpellNum).Type
                    
                        Case SPELL_TYPE_ADDHP
                            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                            Call SendHP(n)
                                    
                        Case SPELL_TYPE_ADDMP
                            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                            Call SendMP(n)
                    
                        Case SPELL_TYPE_ADDSP
                            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                            Call SendMP(n)
                    End Select
                    
                    ' Take away the mana points
                    Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
                    Call SendMP(index)
                    Casted = True
                Else
                    Call BattleMsg(index, "Could not cast spell!", BrightRed, 0)
                End If
            End If
        Else
            Call PlayerMsg(index, "Could not cast spell!", BrightRed)
        End If
    Else
        If CInt(Sqr((GetPlayerX(index) - MapNpc(GetPlayerMap(index), n).x) ^ 2 + ((GetPlayerY(index) - MapNpc(GetPlayerMap(index), n).y) ^ 2))) > Spell(SpellNum).Range Then
            Call BattleMsg(index, "You are too far away to hit the target.", BrightRed, 0)
            Exit Sub
        End If
        
        Player(index).TargetType = TARGET_TYPE_NPC
        
        If Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on a " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_ADDHP
                    MapNpc(GetPlayerMap(index), n).HP = MapNpc(GetPlayerMap(index), n).HP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBHP
                    
                    Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(index), n).num).DEF / 2)
                    If Damage > 0 Then
                        Call AttackNpc(index, n, Damage)
                    Else
                        Call BattleMsg(index, "The spell was to weak to hurt " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & "!", BrightRed, 0)
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
            Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
            Call SendMP(index)
            Casted = True
        Else
            Call BattleMsg(index, "Could not cast spell!", BrightRed, 0)
        End If
    End If
End If

    If Casted = True Then
        Player(index).AttackTimer = GetTickCount
        Player(index).CastedSpell = YES
        Call SendDataToMap(GetPlayerMap(index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & Player(index).Target & SEP_CHAR & Player(index).CastedSpell & SEP_CHAR & END_CHAR)
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
    End If
End Sub
Function GetSpellReqLevel(ByVal index As Long, ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).LevelReq ' - Int(GetClassMAGI(GetPlayerClass(index)) / 4)
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
            If item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
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
            If item(ItemNum).Type <> ITEM_TYPE_ARMOR Then
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
            If item(ItemNum).Type <> ITEM_TYPE_HELMET Then
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
            If item(ItemNum).Type <> ITEM_TYPE_SHIELD Then
                Call SetPlayerShieldSlot(index, 0)
            End If
        Else
            Call SetPlayerShieldSlot(index, 0)
        End If
    End If
End Sub

Public Sub SetPMember(ByVal Leader As Byte, ByVal MemberIndex As Byte)
Dim i As Integer
For i = 1 To MAX_PARTY_MEMBERS
    If Player(Leader).Party.Member(i) = 0 Then
        Player(Leader).Party.Member(i) = MemberIndex
        Exit For
    End If
Next i
Player(MemberIndex).Party.Leader = Leader

For i = 1 To MAX_PARTY_MEMBERS
    If Player(Leader).Party.Member(i) > 0 Then UpdateParty Player(Leader).Party.Member(i)
Next i
End Sub

Public Sub RemovePMember(ByVal index As Byte)
Dim i, b, q As Integer
i = 0

b = Player(index).Party.Leader

If Player(index).Party.Leader = index Then 'Change the party leader!
    For i = 1 To MAX_PARTY_MEMBERS
        If Player(index).Party.Member(i) > 0 And Player(index).Party.Member(i) <> index Then
        Call ChangePLeader(Player(index).Party.Member(i))
        Exit For
        End If
    Next i
End If

i = 0

For q = 1 To MAX_PARTY_MEMBERS ' find which member the player is
    If Player(index).Party.Member(q) = index Then
        Exit For
    End If
Next q

For i = 1 To MAX_PARTY_MEMBERS ' removes player from other members party
    If Player(index).Party.Member(i) > 0 Then Player(Player(index).Party.Member(i)).Party.Member(q) = 0
Next i

Player(index).Party.Leader = 0 'no leader
Player(index).InvitedBy = 0

For i = 1 To MAX_PARTY_MEMBERS ' clears player's party
    Player(index).Party.Member(i) = 0
Next i

Player(index).InParty = False 'not in party

q = 0

If b > 0 Then
    For i = 1 To MAX_PARTY_MEMBERS 'check to see if we need to clear out the party leader
        If Player(b).Party.Member(i) > 0 Then q = q + 1
    Next i

    If q < 1 Then
    Call PlayerMsg(b, "Your party has been disbanded!", White)
    Player(b).InParty = False

        For i = 1 To MAX_PARTY_MEMBERS ' clears player's party
            Player(b).Party.Member(i) = 0
        Next i
    End If
End If

End Sub

Public Sub ChangePLeader(ByVal index As Byte)
Dim i As Integer

Player(index).Party.Leader = index

For i = 1 To MAX_PARTY_MEMBERS
    If Player(index).Party.Member(i) > 0 Then Player(Player(index).Party.Member(i)).Party.Leader = index
Next i

For i = 1 To MAX_PARTY_MEMBERS
    If Player(index).Party.Member(i) > 0 Then Call PlayerMsg(Player(index).Party.Member(i), "Leadership has been passed to " & GetPlayerName(index) & "!", Pink)
Next i

End Sub

Public Sub UpdateParty(ByVal index As Byte)
Player(index).Party = Player(Player(index).Party.Leader).Party
End Sub

Public Sub SetPShare(ByVal index As Byte, ByVal share As Boolean)
Player(index).Party.ShareExp = share
End Sub

Function GetPLeader(ByVal index As Byte) As Byte
    GetPLeader = Player(index).Party.Leader
End Function

Function GetPMember(ByVal index As Byte, ByVal Member As Byte) As Byte
    GetPMember = Player(index).Party.Member(Member)
End Function

Function GetPShare(ByVal index As Byte) As Boolean
    GetPShare = Player(index).Party.ShareExp
End Function

Public Sub ShowPLR(ByVal index As Long)
Dim ls As ListItem

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(index) = True Then
        frmServer.lvUsers.ListItems.Remove index
    End If
    Set ls = frmServer.lvUsers.ListItems.add(index, , index)
    
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

Public Sub RemovePLR()
    frmServer.lvUsers.ListItems.Clear
End Sub

Function CanAttackPlayerWithArrow(ByVal attacker As Long, ByVal Victim As Long) As Boolean

CanAttackPlayerWithArrow = False

' Check To make sure that they dont have access
If GetPlayerAccess(attacker) > ADMIN_MONITER Then
    Call PlayerMsg(attacker, "You cannot attack Any player For thou art an admin!", BrightBlue)
Else
' Check To make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITER Then
    Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
Else
' Check If map Is attackable
If Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
' Make sure they are high enough level
If GetPlayerLevel(attacker) < GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") Then
Call PlayerMsg(attacker, "You are below level" & GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL") & ", you cannot attack another player yet!", BrightRed)
Else
If GetPlayerLevel(Victim) < GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") Then
Call PlayerMsg(attacker, GetPlayerName(Victim) & " is below level" & GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED") & ", you cannot attack this player yet!", BrightRed)
Else
If Trim(GetPlayerGuild(attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
If Trim(GetPlayerGuild(attacker)) <> Trim(GetPlayerGuild(Victim)) Then
    CanAttackPlayerWithArrow = True
Else
    Call PlayerMsg(attacker, "You cant attack a guild member!", BrightRed)
End If
Else
    CanAttackPlayerWithArrow = True
End If
End If
End If
Else
    Call PlayerMsg(attacker, "This Is a safe zone!", BrightRed)
End If
End If
End If

End Function

Function CanAttackNpcWithArrow(ByVal attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(attacker) > 0 Then
        AttackSpeed = item(GetPlayerInvItemNum(attacker, GetPlayerWeaponSlot(attacker))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

CanAttackNpcWithArrow = False

' Check For subscript out of range
If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
Exit Function
End If

' Check For subscript out of range
If MapNpc(GetPlayerMap(attacker), MapNpcNum).num <= 0 Then
Exit Function
End If

MapNum = GetPlayerMap(attacker)
NpcNum = MapNpc(MapNum, MapNpcNum).num

' Make sure the npc isn't already dead
If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
Exit Function
End If

' Make sure they are On the same map
If IsPlaying(attacker) Then
    If NpcNum > 0 And GetTickCount > Player(attacker).AttackTimer + AttackSpeed Then
        ' Check If at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpcWithArrow = True
                    Else
                        Call PlayerMsg(attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If

            Case DIR_DOWN
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpcWithArrow = True
                    Else
                        Call PlayerMsg(attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If

            Case DIR_LEFT
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpcWithArrow = True
                    Else
                        Call PlayerMsg(attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If

            Case DIR_RIGHT
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpcWithArrow = True
                    Else
                        Call PlayerMsg(attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                End If
        End Select
    End If
End If
End Function
