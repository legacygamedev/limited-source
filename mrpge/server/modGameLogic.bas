Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal index As Long) As Long
Dim WeaponSlot As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    'OLD OLD ODL GetPlayerDamage = Int(GetPlayerSTR(index) / 2)
    GetPlayerDamage = Int(GetPlayerSTR(index) * 3.14159265358979)
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
    
    If GetPlayerWeaponSlot(index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(index)
        
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(index, WeaponSlot)).BaseDamage
        
        Call SetPlayerInvItemDur(index, WeaponSlot, GetPlayerInvItemDur(index, WeaponSlot) - 1)
        
        If GetPlayerInvItemDur(index, WeaponSlot) <= 0 Then
            Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " has broken.", RGB_AlertColor)
            Call TakeItem(index, GetPlayerInvItemNum(index, WeaponSlot), 0)
        Else
            If GetPlayerInvItemDur(index, WeaponSlot) <= 5 Then
                Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " is about to break!", RGB_HelpColor)
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
    GetPlayerProtection = Int(GetPlayerDEX(index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, ArmorSlot)).Data2
        Call SetPlayerInvItemDur(index, ArmorSlot, GetPlayerInvItemDur(index, ArmorSlot) - 1)
        
        If GetPlayerInvItemDur(index, ArmorSlot) <= 0 Then
            Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " has broken.", RGB_AlertColor)
            Call TakeItem(index, GetPlayerInvItemNum(index, ArmorSlot), 0)
        Else
            If GetPlayerInvItemDur(index, ArmorSlot) <= 5 Then
                Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " is about to break!", RGB_AlertColor)
            End If
        End If
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, HelmSlot)).Data2
        Call SetPlayerInvItemDur(index, HelmSlot, GetPlayerInvItemDur(index, HelmSlot) - 1)
        
        If GetPlayerInvItemDur(index, HelmSlot) <= 0 Then
            Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " has broken.", RGB_AlertColor)
            Call TakeItem(index, GetPlayerInvItemNum(index, HelmSlot), 0)
        Else
            If GetPlayerInvItemDur(index, HelmSlot) <= 5 Then
                Call PlayerMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " is about to break!", RGB_AlertColor)
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

Function FindOpenInvSlot(ByVal index As Long, ByVal itemnum As Long) As Long
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(itemnum).type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = itemnum Then
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

Function FindOpenMapItemSlot(ByVal mapNum As Long) As Long
Dim i As Long

    FindOpenMapItemSlot = 0
    
    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Function
    End If
    
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(mapNum, i).num = 0 Then
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
Function FindOpenPrayerSlot(ByVal index As Long) As Long
Dim i As Long

    FindOpenPrayerSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerPrayer(index, i) = 0 Then
            FindOpenPrayerSlot = i
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

Function HasPrayer(ByVal index As Long, ByVal PrayerNum As Long) As Boolean
Dim i As Long

    HasPrayer = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerPrayer(index, i) = PrayerNum Then
            HasPrayer = True
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

Function HasItem(ByVal index As Long, ByVal itemnum As Long) As Long
Dim i As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If Item(itemnum).type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next i
End Function

Function BankHasItem(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long
    
    BankHasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_BANK
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(index, i) = itemnum Then
            If Item(itemnum).type = ITEM_TYPE_CURRENCY Then
                BankHasItem = GetPlayerBankItemValue(index, i)
            Else
                BankHasItem = 1
            End If
            Exit Function
        End If
    Next i

End Function

Sub TakeItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long)
Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If Item(itemnum).type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(index, i)).type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(index) > 0 Then
                            If i = GetPlayerWeaponSlot(index) Then
                                Call SetPlayerWeaponSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If itemnum <> GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) Then
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
                                If itemnum <> GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) Then
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
                                If itemnum <> GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) Then
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
                                If itemnum <> GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select

                
                n = Item(GetPlayerInvItemNum(index, i)).type
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

Sub GiveItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = FindOpenInvSlot(index, itemnum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, itemnum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        
        If (Item(itemnum).type = ITEM_TYPE_ARMOR) Or (Item(itemnum).type = ITEM_TYPE_WEAPON) Or (Item(itemnum).type = ITEM_TYPE_HELMET) Or (Item(itemnum).type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(index, i, Item(itemnum).Data1)
        End If
        
        Call SendInventoryUpdate(index, i)
    Else
        Call PlayerMsg(index, "Your inventory is full.", RGB_AlertColor)
    End If
End Sub

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
Dim i As Long

    ' Check for subscript out of range
    If itemnum < 0 Or itemnum > MAX_ITEMS Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    i = FindOpenMapItemSlot(mapNum)
    
    Call SpawnItemSlot(i, itemnum, ItemVal, Item(itemnum).Data1, mapNum, x, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
Dim packet As String
Dim i As Long
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemnum < 0 Or itemnum > MAX_ITEMS Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    i = MapItemSlot
    
    If i <> 0 And itemnum >= 0 And itemnum <= MAX_ITEMS Then
        MapItem(mapNum, i).num = itemnum
        MapItem(mapNum, i).value = ItemVal
        
        If itemnum <> 0 Then
            If (Item(itemnum).type >= ITEM_TYPE_WEAPON) And (Item(itemnum).type <= ITEM_TYPE_SHIELD) Then
                MapItem(mapNum, i).Dur = ItemDur
            Else
                MapItem(mapNum, i).Dur = 0
            End If
        Else
            MapItem(mapNum, i).Dur = 0
        End If
        
        MapItem(mapNum, i).x = x
        MapItem(mapNum, i).y = y
            
        packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & itemnum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(mapNum, i).Dur & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
        Call SendDataToMap(mapNum, packet)
    End If
End Sub

Sub SpawnAllMapsItems()
Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next i
End Sub

Sub SpawnMapItems(ByVal mapNum As Long)
Dim x As Long
Dim y As Long
Dim i As Long
On Error Resume Next
    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Spawn what we have
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (map(mapNum).Tile(x, y).type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(map(mapNum).Tile(x, y).Data1).type = ITEM_TYPE_CURRENCY And map(mapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(map(mapNum).Tile(x, y).Data1, 1, mapNum, x, y)
                Else
                    Call SpawnItem(map(mapNum).Tile(x, y).Data1, map(mapNum).Tile(x, y).Data2, mapNum, x, y)
                End If
            End If
        Next x
    Next y
End Sub

Sub PlayerMapGetItem(ByVal index As Long)
Dim i As Long
Dim n As Long
Dim mapNum As Long
Dim msg As String

    If IsPlaying(index) = False Then
        Exit Sub
    End If
    
    mapNum = GetPlayerMap(index)
    
    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapNum, i).num > 0) And (MapItem(mapNum, i).num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(mapNum, i).x = GetPlayerX(index)) And (MapItem(mapNum, i).y = GetPlayerY(index)) Then
                ' Find open slot
                n = FindOpenInvSlot(index, MapItem(mapNum, i).num)
                
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(index, n, MapItem(mapNum, i).num)
                    If Item(GetPlayerInvItemNum(index, n)).type = ITEM_TYPE_CURRENCY Then
                        Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(mapNum, i).value)
                        msg = "You picked up " & MapItem(mapNum, i).value & " " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(index, n, 0)
                        msg = "You picked up a " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    End If
                    Call checkQuestProgression(index)
                    Call SetPlayerInvItemDur(index, n, MapItem(mapNum, i).Dur)
                        
                    ' Erase item from the map
                    MapItem(mapNum, i).num = 0
                    MapItem(mapNum, i).value = 0
                    MapItem(mapNum, i).Dur = 0
                    MapItem(mapNum, i).x = 0
                    MapItem(mapNum, i).y = 0
                        
                    Call SendInventoryUpdate(index, n)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    Call PlayerMsg(index, msg, RGB_NpcColor)
                    Exit Sub
                Else
                    Call PlayerMsg(index, "Your inventory is full.", RGB_AlertColor)
                    Exit Sub
                End If
            End If
        End If
    Next i
End Sub
'HERE1281
Sub PlayerMapGetSign(ByVal index As Long)
Dim i As Long
Dim n As Long
Dim mapNum As Long
Dim msg As String

    If IsPlaying(index) = False Then
        Exit Sub
    End If
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_SIGN Then
        Dim signToLoad As Long
        
        signToLoad = Val(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
        
        Call PlayerSign(index, Trim(Signs(signToLoad).header), Trim(Signs(signToLoad).msg))
    End If
End Sub

Sub PlayerMapGetLevel(ByVal index As Long)
Dim i As Long
Dim n As Long
Dim mapNum As Long
Dim msg As String

    If IsPlaying(index) = False Then
        Exit Sub
    End If
    'Debug.Print map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_LEVEL Then
        Call CheckPlayerLevelUp(index)
    End If
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
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).type
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
                        
            If Item(GetPlayerInvItemNum(index, InvNum)).type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Ammount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).value = GetPlayerInvItemValue(index, InvNum)
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", RGB_HelpColor)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), i).value = Ammount
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Ammount & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", RGB_HelpColor)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Ammount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(index), i).value = 0
                If Item(GetPlayerInvItemNum(index, InvNum)).type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(index, InvNum)).type <= ITEM_TYPE_SHIELD Then
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & " " & GetPlayerInvItemDur(index, InvNum) & "/" & Item(GetPlayerInvItemNum(index, InvNum)).Data1 & ".", RGB_HelpColor)
                Else
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", RGB_HelpColor)
                End If
                
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).num, Ammount, MapItem(GetPlayerMap(index), i).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsg(index, "To many items already on the ground.", RGB_AlertColor)
        End If
    End If
End Sub

Sub SpawnNpc(ByVal mapnpcnum As Long, ByVal mapNum As Long)
Dim packet As String
Dim NpcNum As Long
Dim i As Long, x As Long, y As Long, n As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    Spawned = False
    
    NpcNum = map(mapNum).Npc(mapnpcnum)
    If NpcNum > 0 Then
        MapNpc(mapNum, mapnpcnum).num = NpcNum
        MapNpc(mapNum, mapnpcnum).target = 0
        
        MapNpc(mapNum, mapnpcnum).HP = GetNpcMaxHP(NpcNum)
        MapNpc(mapNum, mapnpcnum).maxHP = GetNpcMaxHP(NpcNum)
        MapNpc(mapNum, mapnpcnum).MP = GetNpcMaxMP(NpcNum)
        MapNpc(mapNum, mapnpcnum).SP = GetNpcMaxSP(NpcNum)
        MapNpc(mapNum, mapnpcnum).Gold = GetNpcMaxGold(NpcNum)
        MapNpc(mapNum, mapnpcnum).Respawn = GetNpcRespawn(NpcNum)
        MapNpc(mapNum, mapnpcnum).Attack_with_Poison = GetNpcAttack_with_Poison(NpcNum)
                
        MapNpc(mapNum, mapnpcnum).Dir = Int(Rnd * 4)
        'first check to see if there is a spawn npc tile atribute
        For i = 0 To MAX_MAPX Step 1
            For n = 0 To MAX_MAPY Step 1
                If map(mapNum).Tile(i, n).type = TILE_TYPE_NPC_SPAWN And map(mapNum).Tile(i, n).Data1 = mapnpcnum Then
                    'there so spawn npc here
                    MapNpc(mapNum, mapnpcnum).x = i
                    MapNpc(mapNum, mapnpcnum).y = n
                    Spawned = True
                End If
            Next n
        Next i
        
        
        ' Well try 100 times to randomly place the sprite
        If Not Spawned Then
            For i = 1 To 100
                x = Int(Rnd * MAX_MAPX)
                y = Int(Rnd * MAX_MAPY)
                
                ' Check if the tile is walkable
                If map(mapNum).Tile(x, y).type = TILE_TYPE_WALKABLE Then
                    MapNpc(mapNum, mapnpcnum).x = x
                    MapNpc(mapNum, mapnpcnum).y = y
                    Spawned = True
                    Exit For
                End If
            Next i
        End If
        
        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    If map(mapNum).Tile(x, y).type = TILE_TYPE_WALKABLE Then
                        MapNpc(mapNum, mapnpcnum).x = x
                        MapNpc(mapNum, mapnpcnum).y = y
                        Spawned = True
                    End If
                Next x
            Next y
        End If
             
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            packet = "SPAWNNPC" & SEP_CHAR & mapnpcnum & SEP_CHAR & MapNpc(mapNum, mapnpcnum).num & SEP_CHAR & MapNpc(mapNum, mapnpcnum).x & SEP_CHAR & MapNpc(mapNum, mapnpcnum).y & SEP_CHAR & MapNpc(mapNum, mapnpcnum).Dir & SEP_CHAR & END_CHAR
            Call SendDataToMap(mapNum, packet)
        End If
    End If
End Sub

Sub SpawnMapNpcs(ByVal mapNum As Long)
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapNum)
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
    If player(Victim).GettingMap = YES Then
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If (GetPlayerMap(Attacker) = GetPlayerMap(Victim)) And (GetTickCount > player(Attacker).AttackTimer + 950) Then
        
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    'If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                    '    Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                    'Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER And GetPlayerAccess(Attacker) < ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", RGB_AlertColor)
                        Else
                            ' Check if map is attackable
                            If map(GetPlayerMap(Attacker)).Moral > MAP_MORAL_SAFE Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 0 Then
                                    Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", RGB_AlertColor)
                                Else
                                    If GetPlayerLevel(Victim) < 0 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", RGB_AlertColor)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "This is a safe zone!", RGB_AlertColor)
                            End If
                        End If
                    'End If
                End If
            
            Case DIR_DOWN
                If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    'If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                    '    Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                    'Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER And GetPlayerAccess(Attacker) < ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", RGB_AlertColor)
                        Else
                            ' Check if map is attackable
                            If map(GetPlayerMap(Attacker)).Moral > MAP_MORAL_SAFE Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 0 Then
                                    Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", RGB_AlertColor)
                                Else
                                    If GetPlayerLevel(Victim) < 0 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", RGB_AlertColor)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "This is a safe zone!", RGB_AlertColor)
                            End If
                        End If
                    'End If
                End If
        
            Case DIR_LEFT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    'If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                     '   Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                    'Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER And GetPlayerAccess(Attacker) < ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", RGB_AlertColor)
                        Else
                            ' Check if map is attackable
                            If map(GetPlayerMap(Attacker)).Moral > MAP_MORAL_SAFE Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 0 Then
                                    Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", RGB_AlertColor)
                                Else
                                    If GetPlayerLevel(Victim) < 0 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", RGB_AlertColor)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "This is a safe zone!", RGB_AlertColor)
                            End If
                        End If
                    'End If
                End If
            
            Case DIR_RIGHT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    'If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                    '    Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                    'Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER And GetPlayerAccess(Attacker) < ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", RGB_AlertColor)
                        Else
                            ' Check if map is attackable
                            If map(GetPlayerMap(Attacker)).Moral > MAP_MORAL_SAFE Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 0 Then
                                    Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", RGB_AlertColor)
                                Else
                                    If GetPlayerLevel(Victim) < 0 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", RGB_AlertColor)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "This is a safe zone!", RGB_AlertColor)
                            End If
                        End If
                    End If
                'End If
        End Select
    End If
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal mapnpcnum As Long) As Boolean
Dim mapNum As Long, NpcNum As Long

    CanAttackNpc = False
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker), mapnpcnum).num <= 0 Then
        Exit Function
    End If
    
    mapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(mapNum, mapnpcnum).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum, mapnpcnum).HP <= 0 Then
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > player(Attacker).AttackTimer + 950 Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If (MapNpc(mapNum, mapnpcnum).y + 1 = GetPlayerY(Attacker)) And (MapNpc(mapNum, mapnpcnum).x = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", RGB_AlertColor)
                        End If
                    End If
                
                Case DIR_DOWN
                    If (MapNpc(mapNum, mapnpcnum).y - 1 = GetPlayerY(Attacker)) And (MapNpc(mapNum, mapnpcnum).x = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", RGB_AlertColor)
                        End If
                    End If
                
                Case DIR_LEFT
                    If (MapNpc(mapNum, mapnpcnum).y = GetPlayerY(Attacker)) And (MapNpc(mapNum, mapnpcnum).x + 1 = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", RGB_AlertColor)
                        End If
                    End If
                
                Case DIR_RIGHT
                    If (MapNpc(mapNum, mapnpcnum).y = GetPlayerY(Attacker)) And (MapNpc(mapNum, mapnpcnum).x - 1 = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", RGB_AlertColor)
                        End If
                    End If
            End Select
        End If
    End If
End Function

Function CanNpcAttackPlayer(ByVal mapnpcnum As Long, ByVal index As Long) As Boolean
Dim mapNum As Long, NpcNum As Long
    
    CanNpcAttackPlayer = False
    
    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(index) = False Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index), mapnpcnum).num <= 0 Then
        Exit Function
    End If
    
    mapNum = GetPlayerMap(index)
    NpcNum = MapNpc(mapNum, mapnpcnum).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapNum, mapnpcnum).HP <= 0 Then
        Exit Function
    End If
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(mapNum, mapnpcnum).AttackTimer + 1000 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If player(index).GettingMap = YES Then
        Exit Function
    End If
    
    MapNpc(mapNum, mapnpcnum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NpcNum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(mapNum, mapnpcnum).y) And (GetPlayerX(index) = MapNpc(mapNum, mapnpcnum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(mapNum, mapnpcnum).y) And (GetPlayerX(index) = MapNpc(mapNum, mapnpcnum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(mapNum, mapnpcnum).y) And (GetPlayerX(index) + 1 = MapNpc(mapNum, mapnpcnum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(mapNum, mapnpcnum).y) And (GetPlayerX(index) - 1 = MapNpc(mapNum, mapnpcnum).x) Then
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
    
    If player(Victim).Char(player(Victim).CharNum).Access >= 1 Then
        Exit Sub
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
        
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", RGB_WHITE)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", RGB_AlertColor)
            SendStatsInfo (Attacker)
            SendStatsInfo (Victim)
        Else
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", RGB_WHITE)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", RGB_AlertColor)
            SendStatsInfo (Attacker)
            SendStatsInfo (Victim)
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
         Exp = Int(GetPlayerExp(Victim) / ((10 + (Rnd() * 10))))
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If
        
        If Exp = 0 Then
            Call PlayerMsg(Victim, "You lost no experience points.", RGB_AlertColor)
            Call PlayerMsg(Attacker, "You received no experience points from that weak insignificant player.", RGB_HelpColor)
        Else
        'HERE1281
            If map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_ARENA Then
                'Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call PlayerMsg(Victim, "You lost no experience points.", RGB_AlertColor)
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call PlayerMsg(Attacker, "You got " & Exp & " experience points for killing " & GetPlayerName(Victim) & ".", RGB_HelpColor)
            Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call PlayerMsg(Victim, "You lost " & Exp & " experience points.", RGB_AlertColor)
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call PlayerMsg(Attacker, "You got " & Exp & " experience points for killing " & GetPlayerName(Victim) & ".", RGB_AlertColor)
            End If
            
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
        SendStatsInfo (Victim)
                
        ' Check for a level up
        'Call CheckPlayerLevelUp(Attacker)
        
        ' Check if target is player who died and if so set target to 0
        If player(Attacker).TargetType = TARGET_TYPE_PLAYER And player(Attacker).target = Victim Then
            player(Attacker).target = 0
            player(Attacker).TargetType = 0
        End If
        
        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                If map(GetPlayerMap(Attacker)).Moral <> MAP_MORAL_ARENA Then
                    Call SetPlayerPK(Attacker, YES)
                    'If GetPlayerAccess(Attacker) < 1 Then
                    '    Call SetPlayerColour(Attacker, 12, False)
                    'End If
                    Call SendPlayerData(Attacker)
                    Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
                End If
            End If
        Else
            Call SetPlayerPK(Victim, NO)
            'If GetPlayerAccess(Victim) < 1 Then
            '    Call SetPlayerColour(Victim, 15, False)
            'End If
            Call SendPlayerData(Victim)
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
        SendStatsInfo (Victim)
        SendStatsInfo (Attacker)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", RGB_AlertColor)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", RGB_AlertColor)
        Else
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", RGB_AlertColor)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", RGB_AlertColor)
            If Item(n).Poisons = True And Rnd() > 0.75 And GetPlayerPoison(Victim) = False Then
            Call setPlayerPoison(Victim, True, Int(Rnd() * 5 + Item(n).Poison_length), Int(Rnd() * 5 + Item(n).Poison_vital))
            Call PlayerMsg(Victim, "You have been poisoned.", RGB_AlertColor)
            Call PlayerMsg(Attacker, "You poisoned " & GetPlayerName(Victim) & " with your " & Trim(Item(n).Name) & ".", RGB_AlertColor)
        End If
        End If
        
        
        SendStatsInfo (Victim)
    End If
    
    ' Reset timer for attacking
    player(Attacker).AttackTimer = GetTickCount
End Sub

Sub NpcAttackPlayer(ByVal mapnpcnum As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim mapNum As Long

    ' Check for subscript out of range
    If mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim), mapnpcnum).num <= 0 Then
        Exit Sub
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & mapnpcnum & SEP_CHAR & END_CHAR)
    
    mapNum = GetPlayerMap(Victim)
    Name = Trim(Npc(MapNpc(mapNum, mapnpcnum).num).Name)
    
    If Damage >= GetPlayerHP(Victim) Then
        ' Say damage
        Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", RGB_AlertColor)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by a " & Name, RGB_AlertColor)
        
        ' Drop all worn items by victim
        'If GetPlayerWeaponSlot(Victim) > 0 Then
        '    Call PlayerMapDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
        'End If
        'If GetPlayerArmorSlot(Victim) > 0 Then
        '    Call PlayerMapDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
        'End If
        'If GetPlayerHelmetSlot(Victim) > 0 Then
        '    Call PlayerMapDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
        'End If
        'If GetPlayerShieldSlot(Victim) > 0 Then
        '    Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
        'End If
        
        'NEW
        'remove gold
        Dim lngGoldTaken As Long
        lngGoldTaken = HasItem(Victim, 2) 'Player(Victim).Char(Player(Victim).CharNum).
        If lngGoldTaken >= 1 Then
            Call TakeItem(Victim, 2, lngGoldTaken)
            MapNpc(mapNum, mapnpcnum).Gold = MapNpc(mapNum, mapnpcnum).Gold + lngGoldTaken
            'Npc(MapNpc(MapNum, MapNpcNum).Num).Gold = Npc(MapNpc(MapNum, MapNpcNum).Num).Gold + lngGoldTaken
            Call PlayerMsg(Victim, "You have lost " & lngGoldTaken & " gold.", RGB_AlertColor)
            'SaveNpc (MapNpc(MapNum, MapNpcNum).Num)
        End If
        
        ' Calculate exp to give attacker
        Exp = Int(GetPlayerExp(Victim) / ((10 + (Rnd() * 10))))
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If
        
        If Exp = 0 Then
            Call PlayerMsg(Victim, "You lost no experience points.", RGB_AlertColor)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            Call PlayerMsg(Victim, "You lost " & Exp & " experience points.", RGB_AlertColor)
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
        SendStatsInfo (Victim)
        
        ' Set NPC target to 0
        MapNpc(mapNum, mapnpcnum).target = 0
        
        ' If the player the attacker killed was a pk then take it away
        'IMPORTANT change the turning red code
        If GetPlayerPK(Victim) = YES Then
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        'check to see if npc does poison attack
        If GetNpcAttack_with_Poison(mapnpcnum) = True And Rnd() > 0.75 And GetPlayerPoison(Victim) = False Then
            Call setPlayerPoison(Victim, True, Int(Rnd() * 3) + GetNpcAttack_with_Poison_length(mapnpcnum), Int(Rnd() * 4) + GetNpcAttack_with_Poison_vital(mapnpcnum))
            Call PlayerMsg(Victim, "You have been poisoned.", RGB_AlertColor)
        End If
        Call SendHP(Victim)
        SendStatsInfo (Victim)
        ' Say damage
        Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", RGB_AlertColor)
    End If
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal mapnpcnum As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim n As Long, i As Long
Dim str As Long, def As Long, mapNum As Long, NpcNum As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Damage < 0 Then
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
    
    mapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(mapNum, mapnpcnum).num
    Name = Trim(Npc(NpcNum).Name)
        
     If Npc(NpcNum).type = Item(n).weaponType Then
        Damage = CLng(Damage * 1.175)
     Else
        Damage = CLng(Damage * 0.875)
     End If
        
    If Damage >= MapNpc(mapNum, mapnpcnum).HP Then
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points, killing it.", RGB_AlertColor)
            SendStatsInfo (Attacker)
        Else
            Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points, killing it.", RGB_AlertColor)
            SendStatsInfo (Attacker)
        End If
                        
        ' Calculate exp to give attacker
        'EXP CALC
        str = Npc(NpcNum).str
        def = Npc(NpcNum).def
        'OLD EXP
        Exp = str * def * 2
        'NEW EXP SYSTEM
        If Npc(NpcNum).ExpGiven < 0 Then Npc(NpcNum).ExpGiven = 0
        Exp = Npc(NpcNum).ExpGiven
        
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If player(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You have gained " & Exp & " experience points.", RGB_AlertColor)
        Else
            Exp = Exp / 2
            
            If Exp < 0 Then
                Exp = 1
            End If
            
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You have gained " & Exp & " party experience points.", RGB_AlertColor)
            
            n = player(Attacker).PartyPlayer
            If n > 0 Then
                Call SetPlayerExp(n, GetPlayerExp(n) + Exp)
                Call PlayerMsg(n, "You have gained " & Exp & " party experience points.", RGB_AlertColor)
            End If
        End If
                                
        ' Drop the goods if they get it
        n = Int(Rnd * Npc(NpcNum).DropChance) + 1
        If n = 1 Then
            Call SpawnItem(Npc(NpcNum).DropItem, Npc(NpcNum).DropItemValue, mapNum, MapNpc(mapNum, mapnpcnum).x, MapNpc(mapNum, mapnpcnum).y)
            'Call SpawnItem(2, Npc(NpcNum).Gold, MapNum, MapNpc(MapNum, MapNpcNum).x + 1, MapNpc(MapNum, MapNpcNum).y + 1)
            Call GiveItem(Attacker, 2, MapNpc(mapNum, mapnpcnum).Gold)
            Call SendPlayerData(Attacker)
        Else
            If MapNpc(mapNum, mapnpcnum).Gold > Npc(NpcNum).Gold Then
                Call GiveItem(Attacker, 2, MapNpc(mapNum, mapnpcnum).Gold)
                Call SendPlayerData(Attacker)
            End If
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapNum, mapnpcnum).num = 0
        MapNpc(mapNum, mapnpcnum).SpawnWait = GetTickCount
        MapNpc(mapNum, mapnpcnum).HP = 0
        If MapNpc(mapNum, mapnpcnum).Respawn = False Then
            map(mapNum).Npc(mapnpcnum) = 0
            map(mapNum).Npc(mapnpcnum) = 0
            Call ClearMapNpc(mapnpcnum, mapNum)
        End If
            
        Call SendDataToMap(mapNum, "NPCDEAD" & SEP_CHAR & mapnpcnum & SEP_CHAR & END_CHAR)
        
        ' Check for level up
        'Call CheckPlayerLevelUp(Attacker)
        
        ' Check for level up party member
        'If Player(Attacker).InParty = YES Then
        '    Call CheckPlayerLevelUp(Player(Attacker).PartyPlayer)
        'End If
    
        ' Check if target is npc that died and if so set target to 0
        If player(Attacker).TargetType = TARGET_TYPE_NPC And player(Attacker).target = mapnpcnum Then
            player(Attacker).target = 0
            player(Attacker).TargetType = 0
        End If
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapNum, mapnpcnum).HP = MapNpc(mapNum, mapnpcnum).HP - Damage
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", RGB_AlertColor)
        Else
            Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", RGB_AlertColor)
        End If
        
        ' Check if we should send a message
        If MapNpc(mapNum, mapnpcnum).target = 0 And MapNpc(mapNum, mapnpcnum).target <> Attacker Then
            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), RGB_SayColor)
            End If
        End If
        
        ' Set the NPC target to the player
        MapNpc(mapNum, mapnpcnum).target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(mapNum, mapnpcnum).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapNum, i).num = MapNpc(mapNum, mapnpcnum).num Then
                    MapNpc(mapNum, i).target = Attacker
                End If
            Next i
        End If
    End If

    SendMapNpcsToMap (player(Attacker).Char(player(Attacker).CharNum).map)
    SendStatsInfo (Attacker)
    ' Reset attack timer
    player(Attacker).AttackTimer = GetTickCount
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
Dim packet As String
Dim ShopNum As Long, OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    ShopNum = map(GetPlayerMap(index)).Shop
    If ShopNum > 0 Then
        If Trim(Shop(ShopNum).LeaveSay) <> "" Then
            Call PlayerMsg(index, Trim(Shop(ShopNum).Name) & ": " & Trim(Shop(ShopNum).LeaveSay), RGB_SayColor)
        End If
    End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    Call SendLeaveMap(index, OldMap)
    
    Call SetPlayerMap(index, mapNum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    DoEvents
    
    ' Check if there is an npc on the map and say hello if so
    ShopNum = map(GetPlayerMap(index)).Shop
    If ShopNum > 0 Then
        If Trim(Shop(ShopNum).JoinSay) <> "" Then
            Call PlayerMsg(index, Trim(Shop(ShopNum).Name) & ": " & Trim(Shop(ShopNum).JoinSay), RGB_SayColor)
        End If
    End If
            
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapNum) = YES
    
    player(index).GettingMap = YES
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & mapNum & SEP_CHAR & map(mapNum).Revision & SEP_CHAR & END_CHAR)
    DoEvents
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim packet As String
Dim mapNum As Long
Dim x As Long
Dim y As Long
Dim i As Long
Dim Moved As Byte
Dim splitArr() As String

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
                If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type <> TILE_TYPE_KEY Or (map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                        
                        packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(index)).Up > 0 Then
                    Call PlayerWarp(index, map(GetPlayerMap(index)).Up, GetPlayerX(index), MAX_MAPY)
                    DoEvents
                    'Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    Moved = YES
                End If
            End If
                    
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < MAX_MAPY Then
                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type <> TILE_TYPE_KEY Or (map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                        
                        packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    DoEvents
                    'Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    Moved = YES
                End If
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type <> TILE_TYPE_KEY Or (map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                        
                        packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), packet)
                        
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(index)).Left > 0 Then
                    Call PlayerWarp(index, map(GetPlayerMap(index)).Left, MAX_MAPX, GetPlayerY(index))
                    DoEvents
                    'Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    Moved = YES
                End If
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < MAX_MAPX Then
                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type <> TILE_TYPE_KEY Or (map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                        
                        packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    DoEvents
                    Moved = YES
                End If
            End If
    End Select
        
    ' Check to see if the tile is a warp tile, and if so warp them
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_WARP Then
        mapNum = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        x = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        y = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3
                        
        Call PlayerWarp(index, mapNum, x, y)
        Moved = YES
    End If
    ' Check to see if the tile is a warp tile, and if so warp them
'    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_WARP_DOOR Then
'        MapNum = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
'        x = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
'        y = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3
'
'        Call PlayerWarp(index, MapNum, x, y)
'        Call SendSound(index, "doorslam")
'        Moved = YES
'    End If
'
    ' Check to see if the tile is a damage tile
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_DAMAGE Then
        If Val(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1) >= GetPlayerHP(index) Then
            Call PlayerMsg(index, "You were killed.", RGB_AlertColor)
            ' Warp player away
            Call PlayerWarp(index, START_MAP, START_X, START_Y)
            
            ' Restore vitals
            Call SetPlayerHP(index, GetPlayerMaxHP(index))
            Call SetPlayerMP(index, GetPlayerMaxMP(index))
            Call SetPlayerSP(index, GetPlayerMaxSP(index))
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
        Else
            Call PlayerMsg(index, "You got damaged for " & Val(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1) & " HP.", RGB_AlertColor)
            ' Player not dead, just do the damage
            Call SetPlayerHP(index, GetPlayerHP(index) - Val(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1))
            Call SendHP(index)
        End If
        Moved = YES
    End If
    
    ' Check to see if the tile is a heal tile
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_HEAL Then
        If GetPlayerHP(index) + Val(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1) >= GetPlayerMaxHP(index) Then
            Call PlayerMsg(index, "You were healed to full health.", RGB_HelpColor)
            Call SetPlayerHP(index, GetPlayerMaxHP(index))
            Call SendHP(index)
        Else
            Call PlayerMsg(index, "You were healed for " & Val(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1) & " HP.", RGB_AlertColor)
            Call SetPlayerHP(index, GetPlayerHP(index) + Val(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1))
            Call SendSound(index, "heal")
            Call SendHP(index)
        End If
        Call SendSound(index, "heal")
        Moved = YES
    End If
    
    ' Player warp with level requirement. not implimented on client but coded on the server
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_WARP_LEVEL Then
        mapNum = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        x = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        y = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3
        If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data5 = 0 Then map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data5 = 32767
        If (GetPlayerLevel(index) >= map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data4 And GetPlayerLevel(index) <= map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data5) Then
            Call PlayerWarp(index, mapNum, x, y)
        Else
            
                Call PlayerWarp(index, mapNum, GetPlayerX(index) - 1, GetPlayerY(index))
                Call PlayerMsg(index, "You are of the wrong level", RGB_AlertColor)
            
        End If
        'Call PlayerWarp(index, MapNum, x, y)
        Moved = YES
    End If
    
    ' Check to see if the tile is a Sign tile - find PlayerMapGetSign
    'If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SIGN Then
    '    Dim signToLoad As Long
    '    signToLoad = Val(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
    '    Call PlayerSign(index, Trim(Signs(signToLoad).header), Trim(Signs(signToLoad).msg))
    '    Moved = YES
   ' End If
    
    ' Check for key trigger open
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).type = TILE_TYPE_KEYOPEN Then
        x = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        y = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        
        If map(GetPlayerMap(index)).Tile(x, y).type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
            TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", RGB_HelpColor)
        End If
    End If
    
    'Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' They tried to hack
    If Moved = NO Then
        Call HackingAttempt(index, "Position Modification")
    End If
End Sub

'Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long)
'Dim packet As String
'Dim MapNum As Long
'Dim x As Long
'Dim y As Long
'Dim MapNum1 As Long
'Dim x1 As Long
'Dim y1 As Long
'Dim i As Long
'Dim Moved As Byte
'Dim splitArr() As String
'
'    ' Check for subscript out of range
'    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
'        Exit Sub
'    End If
'    MapNum = GetPlayerMap(index)
'    x = GetPlayerX(index)
'    y = GetPlayerY(index)
'    Call SetPlayerDir(index, Dir)
'
'    Moved = NO
'
'    Select Case Dir
'        Case DIR_UP
'            ' Check to make sure not outside of boundries
'            If y > 0 Then
'                ' Check to make sure that the tile is walkable
'                If map(MapNum).Tile(x, y - 1).Type <> TILE_TYPE_BLOCKED Then
'                    ' Check to see if the tile is a key and if it is check if its opened
'                    If map(MapNum).Tile(x, y - 1).Type <> TILE_TYPE_KEY Or (map(MapNum).Tile(x, y - 1).Type = TILE_TYPE_KEY And TempTile(MapNum).DoorOpen(x, y - 1) = YES) Then
'                        Call SetPlayerY(index, y - 1)
'
'                        packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
'                        Call SendDataToMapBut(index, MapNum, packet)
'                        Moved = YES
'                    End If
'                End If
'            Else
'                ' Check to see if we can move them to the another map
'                If map(MapNum).Up > 0 Then
'                    Call PlayerWarp(index, map(MapNum).Up, x, MAX_MAPY)
'                    Moved = YES
'                End If
'            End If
'
'        Case DIR_DOWN
'            ' Check to make sure not outside of boundries
'            If y < MAX_MAPY Then
'                ' Check to make sure that the tile is walkable
'                If map(MapNum).Tile(x, y + 1).Type <> TILE_TYPE_BLOCKED Then
'                    ' Check to see if the tile is a key and if it is check if its opened
'                    If map(MapNum).Tile(x, y + 1).Type <> TILE_TYPE_KEY Or (map(MapNum).Tile(x, y + 1).Type = TILE_TYPE_KEY And TempTile(MapNum).DoorOpen(x, y + 1) = YES) Then
'                        Call SetPlayerY(index, y + 1)
'
'                        packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
'                        Call SendDataToMapBut(index, MapNum, packet)
'                        Moved = YES
'                    End If
'                End If
'            Else
'                ' Check to see if we can move them to the another map
'                If map(MapNum).Down > 0 Then
'                    Call PlayerWarp(index, map(MapNum).Down, x, 0)
'                    Moved = YES
'                End If
'            End If
'
'        Case DIR_LEFT
'            ' Check to make sure not outside of boundries
'            If x > 0 Then
'                ' Check to make sure that the tile is walkable
'                If map(MapNum).Tile(x - 1, y).Type <> TILE_TYPE_BLOCKED Then
'                    ' Check to see if the tile is a key and if it is check if its opened
'                    If map(MapNum).Tile(x - 1, y).Type <> TILE_TYPE_KEY Or (map(MapNum).Tile(x - 1, y).Type = TILE_TYPE_KEY And TempTile(MapNum).DoorOpen(x - 1, y) = YES) Then
'                        Call SetPlayerX(index, x - 1)
'
'                        packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
'                        Call SendDataToMapBut(index, MapNum, packet)
'                        Moved = YES
'                    End If
'                End If
'            Else
'                ' Check to see if we can move them to the another map
'                If map(MapNum).Left > 0 Then
'                    Call PlayerWarp(index, map(MapNum).Left, MAX_MAPX, y)
'                    Moved = YES
'                End If
'            End If
'
'        Case DIR_RIGHT
'            ' Check to make sure not outside of boundries
'            If x < MAX_MAPX Then
'                ' Check to make sure that the tile is walkable
'                If map(MapNum).Tile(x + 1, y).Type <> TILE_TYPE_BLOCKED Then
'                    ' Check to see if the tile is a key and if it is check if its opened
'                    If map(MapNum).Tile(x + 1, y).Type <> TILE_TYPE_KEY Or (map(MapNum).Tile(x + 1, y).Type = TILE_TYPE_KEY And TempTile(MapNum).DoorOpen(x + 1, y) = YES) Then
'                        Call SetPlayerX(index, x + 1)
'
'                        packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
'                        Call SendDataToMapBut(index, MapNum, packet)
'                        Moved = YES
'                    End If
'                End If
'            Else
'                'MapNum = MapNum
'                ' Check to see if we can move them to the another map
'                If map(MapNum).Right > 0 Then
'
'                    Call PlayerWarp(index, map(MapNum).Right, 0, y)
'                    Moved = YES
'                End If
'            End If
'    End Select
'
'    ' Check to see if the tile is a warp tile, and if so warp them
'    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_WARP Then
'        MapNum1 = map(MapNum).Tile(x, y).Data1
'        x1 = map(MapNum).Tile(x, y).Data2
'        y1 = map(MapNum).Tile(x, y).Data3
'
'        Call PlayerWarp(index, MapNum1, x1, y1)
'        Moved = YES
'    End If
'    ' Check to see if the tile is a warp tile, and if so warp them
''    If map(MapNum).Tile(x, y).Type = TILE_TYPE_WARP_DOOR Then
''        MapNum = map(MapNum).Tile(x, y).Data1
''        x = map(MapNum).Tile(x, y).Data2
''        y = map(MapNum).Tile(x, y).Data3
''
''        Call PlayerWarp(index, MapNum, x, y)
''        Call SendSound(index, "doorslam")
''        Moved = YES
''    End If
''
'    ' Check to see if the tile is a damage tile
'    If map(MapNum).Tile(x, y).Type = TILE_TYPE_DAMAGE Then
'        If Val(map(MapNum).Tile(x, y).Data1) >= GetPlayerHP(index) Then
'            Call PlayerMsg(index, "You were killed.", Red)
'            ' Warp player away
'            Call PlayerWarp(index, START_MAP, START_X, START_Y)
'
'            ' Restore vitals
'            Call SetPlayerHP(index, GetPlayerMaxHP(index))
'            Call SetPlayerMP(index, GetPlayerMaxMP(index))
'            Call SetPlayerSP(index, GetPlayerMaxSP(index))
'            Call SendHP(index)
'            Call SendMP(index)
'            Call SendSP(index)
'        Else
'            Call PlayerMsg(index, "You got damaged for " & Val(map(MapNum).Tile(x, y).Data1) & " HP.", Red)
'            ' Player not dead, just do the damage
'            Call SetPlayerHP(index, GetPlayerHP(index) - Val(map(MapNum).Tile(x, y).Data1))
'            Call SendHP(index)
'        End If
'        Moved = YES
'    End If
'
'    ' Check to see if the tile is a heal tile
'    If map(MapNum).Tile(x, y).Type = TILE_TYPE_HEAL Then
'        If GetPlayerHP(index) + Val(map(MapNum).Tile(x, y).Data1) >= GetPlayerMaxHP(index) Then
'            Call PlayerMsg(index, "You were healed to full health.", Red)
'            Call SetPlayerHP(index, GetPlayerMaxHP(index))
'            Call SendHP(index)
'        Else
'            Call PlayerMsg(index, "You were healed for " & Val(map(MapNum).Tile(x, y).Data1) & " HP.", Red)
'            Call SetPlayerHP(index, GetPlayerHP(index) + Val(map(MapNum).Tile(x, y).Data1))
'            Call SendSound(index, "heal")
'            Call SendHP(index)
'        End If
'        Call SendSound(index, "heal")
'        Moved = YES
'    End If
'
'    'check to see if its a level up square
'    If map(MapNum).Tile(x, y).Type = TILE_TYPE_LEVEL Then
'        Call CheckPlayerLevelUp(index)
'    End If
'
'    ' Player warp with level requirement. not implimented on client but coded on the server
'    If map(MapNum).Tile(x, y).Type = TILE_TYPE_WARP_LEVEL Then
'        MapNum1 = map(MapNum).Tile(x, y).Data1
'        x1 = map(MapNum).Tile(x, y).Data2
'        y1 = map(MapNum).Tile(x, y).Data3
'        If map(MapNum).Tile(x, y).Data5 = 0 Then map(MapNum).Tile(x, y).Data5 = 32767
'        If (GetPlayerLevel(index) >= map(MapNum).Tile(x, y).Data4 And GetPlayerLevel(index) <= map(MapNum).Tile(x, y).Data5) Then
'            Call PlayerWarp(index, MapNum1, x1, y1)
'        Else
'            Debug.Print "DIR: " & GetPlayerDir(index)
'                Select Case GetPlayerDir(index)
'                    Case Is = 0
'                        Call PlayerWarp(index, MapNum, x, y + 1)
'                    Case Is = 1
'                        Call PlayerWarp(index, MapNum, x, y - 1)
'                    Case Is = 2
'                        Call PlayerWarp(index, MapNum, x + 1, y)
'                    Case Is = 3
'                        Call PlayerWarp(index, MapNum, x - 1, y)
'                End Select
'
'                Call PlayerMsg(index, "You are of the wrong level", Red)
'
'        End If
'        'Call PlayerWarp(index, MapNum, x, y)
'        Moved = YES
'    End If
'
'    ' Check to see if the tile is a Sign tile - find PlayerMapGetSign
'    'If Map(MapNum).Tile(x, y).Type = TILE_TYPE_SIGN Then
'    '    Dim signToLoad As Long
'    '    signToLoad = Val(Map(MapNum).Tile(x, y).Data1)
'    '    Call PlayerSign(index, Trim(Signs(signToLoad).header), Trim(Signs(signToLoad).msg))
'    '    Moved = YES
'   ' End If
'
'    ' Check for key trigger open
'    If map(MapNum).Tile(x, y).Type = TILE_TYPE_KEYOPEN Then
'        x1 = map(MapNum).Tile(x, y).Data1
'        y1 = map(MapNum).Tile(x, y).Data2
'
'        If map(MapNum).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(MapNum).DoorOpen(x1, y1) = NO Then
'            TempTile(MapNum).DoorOpen(x1, y1) = YES
'            TempTile(MapNum).DoorTimer = GetTickCount
'
'            Call SendDataToMap(MapNum, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
'            Call MapMsg(MapNum, "A door has been unlocked.", White)
'        End If
'    End If
'
'    ' They tried to hack
'    If Moved = NO Then
'        Call HackingAttempt(index, "Position Modification")
'    End If
'End Sub

Function CanNpcMove(ByVal mapNum As Long, ByVal mapnpcnum As Long, ByVal Dir) As Boolean
Dim i As Long, n As Long
Dim x As Long, y As Long

    CanNpcMove = False
    
    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    
    x = MapNpc(mapNum, mapnpcnum).x
    y = MapNpc(mapNum, mapnpcnum).y
    
    CanNpcMove = True
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = map(mapNum).Tile(x, y - 1).type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum, mapnpcnum).x) And (GetPlayerY(i) = MapNpc(mapNum, mapnpcnum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapnpcnum) And (MapNpc(mapNum, i).num > 0) And (MapNpc(mapNum, i).x = MapNpc(mapNum, mapnpcnum).x) And (MapNpc(mapNum, i).y = MapNpc(mapNum, mapnpcnum).y - 1) Then
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
                n = map(mapNum).Tile(x, y + 1).type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum, mapnpcnum).x) And (GetPlayerY(i) = MapNpc(mapNum, mapnpcnum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapnpcnum) And (MapNpc(mapNum, i).num > 0) And (MapNpc(mapNum, i).x = MapNpc(mapNum, mapnpcnum).x) And (MapNpc(mapNum, i).y = MapNpc(mapNum, mapnpcnum).y + 1) Then
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
                n = map(mapNum).Tile(x - 1, y).type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum, mapnpcnum).x - 1) And (GetPlayerY(i) = MapNpc(mapNum, mapnpcnum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapnpcnum) And (MapNpc(mapNum, i).num > 0) And (MapNpc(mapNum, i).x = MapNpc(mapNum, mapnpcnum).x - 1) And (MapNpc(mapNum, i).y = MapNpc(mapNum, mapnpcnum).y) Then
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
                n = map(mapNum).Tile(x + 1, y).type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum, mapnpcnum).x + 1) And (GetPlayerY(i) = MapNpc(mapNum, mapnpcnum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapnpcnum) And (MapNpc(mapNum, i).num > 0) And (MapNpc(mapNum, i).x = MapNpc(mapNum, mapnpcnum).x + 1) And (MapNpc(mapNum, i).y = MapNpc(mapNum, mapnpcnum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Sub NpcMove(ByVal mapNum As Long, ByVal mapnpcnum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim packet As String
Dim x As Long
Dim y As Long
Dim i As Long

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    MapNpc(mapNum, mapnpcnum).Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            MapNpc(mapNum, mapnpcnum).y = MapNpc(mapNum, mapnpcnum).y - 1
            packet = "NPCMOVE" & SEP_CHAR & mapnpcnum & SEP_CHAR & MapNpc(mapNum, mapnpcnum).x & SEP_CHAR & MapNpc(mapNum, mapnpcnum).y & SEP_CHAR & MapNpc(mapNum, mapnpcnum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(mapNum, packet)
    
        Case DIR_DOWN
            MapNpc(mapNum, mapnpcnum).y = MapNpc(mapNum, mapnpcnum).y + 1
            packet = "NPCMOVE" & SEP_CHAR & mapnpcnum & SEP_CHAR & MapNpc(mapNum, mapnpcnum).x & SEP_CHAR & MapNpc(mapNum, mapnpcnum).y & SEP_CHAR & MapNpc(mapNum, mapnpcnum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(mapNum, packet)
    
        Case DIR_LEFT
            MapNpc(mapNum, mapnpcnum).x = MapNpc(mapNum, mapnpcnum).x - 1
            packet = "NPCMOVE" & SEP_CHAR & mapnpcnum & SEP_CHAR & MapNpc(mapNum, mapnpcnum).x & SEP_CHAR & MapNpc(mapNum, mapnpcnum).y & SEP_CHAR & MapNpc(mapNum, mapnpcnum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(mapNum, packet)
    
        Case DIR_RIGHT
            MapNpc(mapNum, mapnpcnum).x = MapNpc(mapNum, mapnpcnum).x + 1
            packet = "NPCMOVE" & SEP_CHAR & mapnpcnum & SEP_CHAR & MapNpc(mapNum, mapnpcnum).x & SEP_CHAR & MapNpc(mapNum, mapnpcnum).y & SEP_CHAR & MapNpc(mapNum, mapnpcnum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(mapNum, packet)
    End Select
End Sub

Sub NpcMoveAway(ByVal y As Long, ByVal x As Long, ByVal speed As Long, ByVal target As Long)
Dim i As Long, didwalk As Boolean
didwalk = False

i = Int(Rnd * 5)

' Lets move the npc
Select Case i
    Case 0
        ' Up
        If MapNpc(y, x).y < GetPlayerY(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_UP) Then
                Call NpcMove(y, x, DIR_UP, speed)
                didwalk = True
            End If
        End If
        ' Down
        If MapNpc(y, x).y > GetPlayerY(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_DOWN) Then
                Call NpcMove(y, x, DIR_DOWN, speed)
                didwalk = True
            End If
        End If
        ' Left
        If MapNpc(y, x).x < GetPlayerX(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_LEFT) Then
                Call NpcMove(y, x, DIR_LEFT, speed)
                didwalk = True
            End If
        End If
        ' Right
        If MapNpc(y, x).x > GetPlayerX(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_RIGHT) Then
                Call NpcMove(y, x, DIR_RIGHT, speed)
                didwalk = True
            End If
        End If
    
    Case 1
        ' Right
        If MapNpc(y, x).x > GetPlayerX(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_RIGHT) Then
                Call NpcMove(y, x, DIR_RIGHT, speed)
                didwalk = True
            End If
        End If
        ' Left
        If MapNpc(y, x).x < GetPlayerX(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_LEFT) Then
                Call NpcMove(y, x, DIR_LEFT, speed)
                didwalk = True
            End If
        End If
        ' Down
        If MapNpc(y, x).y > GetPlayerY(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_DOWN) Then
                Call NpcMove(y, x, DIR_DOWN, speed)
                didwalk = True
            End If
        End If
        ' Up
        If MapNpc(y, x).y < GetPlayerY(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_UP) Then
                Call NpcMove(y, x, DIR_UP, speed)
                didwalk = True
            End If
        End If
        
    Case 2
        ' Down
        If MapNpc(y, x).y > GetPlayerY(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_DOWN) Then
                Call NpcMove(y, x, DIR_DOWN, speed)
                didwalk = True
            End If
        End If
        ' Up
        If MapNpc(y, x).y < GetPlayerY(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_UP) Then
                Call NpcMove(y, x, DIR_UP, speed)
                didwalk = True
            End If
        End If
        ' Right
        If MapNpc(y, x).x > GetPlayerX(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_RIGHT) Then
                Call NpcMove(y, x, DIR_RIGHT, speed)
                didwalk = True
            End If
        End If
        ' Left
        If MapNpc(y, x).x < GetPlayerX(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_LEFT) Then
                Call NpcMove(y, x, DIR_LEFT, speed)
                didwalk = True
            End If
        End If
    
    Case 3
        ' Left
        If MapNpc(y, x).x < GetPlayerX(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_LEFT) Then
                Call NpcMove(y, x, DIR_LEFT, speed)
                didwalk = True
            End If
        End If
        ' Right
        If MapNpc(y, x).x > GetPlayerX(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_RIGHT) Then
                Call NpcMove(y, x, DIR_RIGHT, speed)
                didwalk = True
            End If
        End If
        ' Up
        If MapNpc(y, x).y < GetPlayerY(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_UP) Then
                Call NpcMove(y, x, DIR_UP, speed)
                didwalk = True
            End If
        End If
        ' Down
        If MapNpc(y, x).y > GetPlayerY(target) And didwalk = False Then
            If CanNpcMove(y, x, DIR_DOWN) Then
                Call NpcMove(y, x, DIR_DOWN, speed)
                didwalk = True
            End If
        End If
End Select
End Sub

Sub NpcDir(ByVal mapNum As Long, ByVal mapnpcnum As Long, ByVal Dir As Long)
Dim packet As String

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapnpcnum <= 0 Or mapnpcnum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    MapNpc(mapNum, mapnpcnum).Dir = Dir
    packet = "NPCDIR" & SEP_CHAR & mapnpcnum & SEP_CHAR & Dir & SEP_CHAR & END_CHAR
    Call SendDataToMap(mapNum, packet)
End Sub

Sub JoinGame(ByVal index As Long)
    ' Set the flag so we know the person is in the game
    player(index).InGame = True
        
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITER Then
        Call GlobalMsg(GetPlayerName(index) & " has joined " & GAME_NAME & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has joined " & GAME_NAME & "!", White)
    End If
        
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(index, "LOGINOK" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
    
    ' Send some more little goodies, no need to explain these
'    Call CheckEquippedItems(index)
'    Call SendClasses(index)
'    Call SendItems(index)
'    Call SendSigns(index)
'    Call SendNpcs(index)
'    Call SendShops(index)
'    Call SendSpells(index)
'    Call SendPrayers(index)
'    Call SendInventory(index)
'    Call SendWornEquipment(index)
'    Call SendHP(index)
'    Call SendMP(index)
'    Call SendSP(index)
'    Call SendPP(index)
'    Call SendStats(index)
'    Call SendStatsInfo(index)
'    Call SendWeatherTo(index)
'    Call SendTimeTo(index)
    Call SendMap(index, GetPlayerMap(index))
    'Call SendPets(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(index, "INGAME" & SEP_CHAR & GAME_NAME & SEP_CHAR & END_CHAR)
End Sub

Sub LeftGame(ByVal index As Long)
Dim n As Long

    If player(index).InGame = True Then
        player(index).InGame = False
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) = 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' Check for boot map
        If map(GetPlayerMap(index)).BootMap > 0 Then
            Call SetPlayerX(index, map(GetPlayerMap(index)).BootX)
            Call SetPlayerY(index, map(GetPlayerMap(index)).BootY)
            Call SetPlayerMap(index, map(GetPlayerMap(index)).BootMap)
        End If
        
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If player(index).InParty = YES Then
            n = player(index).PartyPlayer
            
            Call PlayerMsg(n, GetPlayerName(index) & " has left " & GAME_NAME & ", disbanning party.", RGB_AlertColor)
            player(n).InParty = NO
            player(n).PartyPlayer = 0
        End If
            
        Call SavePlayer(index, False)
    
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
End Sub

Function GetTotalMapPlayers(ByVal mapNum As Long) As Long
Dim i As Long, n As Long

    n = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = mapNum Then
            n = n + 1
        End If
    Next i
    
    GetTotalMapPlayers = n
End Function

Function GetNpcMaxHP(ByVal NpcNum As Long)
Dim x As Long, y As Long
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxHP = 0
        Exit Function
    End If
    'New HP SYSTEM
    'x = Npc(NpcNum).STR
    'y = Npc(NpcNum).DEF
    GetNpcMaxHP = Npc(NpcNum).HP
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
        
    GetNpcMaxSP = Npc(NpcNum).speed * 2
End Function

Function GetNpcMaxGold(ByVal NpcNum As Long)
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxGold = 0
        Exit Function
    End If
        
    GetNpcMaxGold = Npc(NpcNum).Gold
End Function

Function GetNpcRespawn(ByVal NpcNum As Long)
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcRespawn = 0
        Exit Function
    End If
        
    GetNpcRespawn = Npc(NpcNum).Respawn
End Function

Function GetNpcAttack_with_Poison(ByVal NpcNum As Long)
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcAttack_with_Poison = False
        Exit Function
    End If
        
    GetNpcAttack_with_Poison = Npc(NpcNum).Attack_with_Poison
End Function
Function GetNpcAttack_with_Poison_length(ByVal NpcNum As Long) As Long
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcAttack_with_Poison_length = 0
        Exit Function
    End If
        
    GetNpcAttack_with_Poison_length = Npc(NpcNum).Poison_length
End Function
Function GetNpcAttack_with_Poison_vital(ByVal NpcNum As Long) As Long
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcAttack_with_Poison_vital = False
        Exit Function
    End If
        
    GetNpcAttack_with_Poison_vital = Npc(NpcNum).Poison_vital
End Function

Function GetPlayerHPRegen(ByVal index As Long)
Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerHPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerCON(index) / 2)
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
    
    i = Int(GetPlayerWIZ(index) / 2)
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
    
    i = Int(GetPlayerDEX(index) / 2)
    If i < 2 Then i = 2
    
    GetPlayerSPRegen = i
End Function

Function GetNpcHPRegen(ByVal NpcNum As Long)
Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcHPRegen = 0
        Exit Function
    End If
    
    i = Int(Npc(NpcNum).def / 3)
    If i < 1 Then i = 1
    
    GetNpcHPRegen = i
End Function

Sub CheckPlayerLevelUp(ByVal index As Long)
Dim i As Long
Dim extraEXP As Long
    ' Check if attacker got a level up
    Debug.Print GetPlayerLevel(index)
    Debug.Print GetPlayerNextLevel(index)
    
    If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
        extraEXP = GetPlayerExp(index) - GetPlayerNextLevel(index)
        Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
        If extraEXP < 0 Then extraEXP = 0
        ' Get the ammount of skill points to add
        i = Int(GetPlayerINT(index) / 10)
        If i < 1 Then i = 1
        If i > 3 Then i = 3
            
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)
        Call SetPlayerExp(index, extraEXP)
        Debug.Print "EXTRA: " & extraEXP
        Call GlobalMsg(GetPlayerName(index) & " has gained a level!", Brown)
        Call SendStatsInfo(index)
        Call PlayerMsg(index, "You have gained a level!  You now have " & GetPlayerPOINTS(index) & " stat points to distribute.", RGB_AlertColor)
    Else
        Call PlayerMsg(index, "You need more exp", RGB_AlertColor)
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
        Call PlayerMsg(index, "You do not have this spell!", RGB_AlertColor)
        Exit Sub
    End If
    
    i = GetSpellReqLevel(index, SpellNum)
    MPReq = (i + Spell(SpellNum).Data1 + Spell(SpellNum).Data2 + Spell(SpellNum).Data3)
    
    ' Check if they have enough MP
    If GetPlayerMP(index) < MPReq Then
        Call PlayerMsg(index, "Not enough mana points!", RGB_AlertColor)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If i > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & i & " to cast this spell.", RGB_AlertColor)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < player(index).AttackTimer + 1000 Then
        Exit Sub
    End If
    
    ' Check if the spell is a give item and do that instead of a stat modification
    If Spell(SpellNum).type = SPELL_TYPE_GIVEITEM Then
        n = FindOpenInvSlot(index, Spell(SpellNum).Data1)
        
        If n > 0 Then
            Call GiveItem(index, Spell(SpellNum).Data1, Spell(SpellNum).Data2)
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & ".", RGB_AlertColor)
            
            ' Take away the mana points
            Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
            Call SendMP(index)
            Casted = True
        Else
            Call PlayerMsg(index, "Your inventory is full!", RGB_AlertColor)
        End If
        
        Exit Sub
    End If
        
    n = player(index).target
    
    If player(index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
            If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And map(GetPlayerMap(index)).Moral >= MAP_MORAL_NONE And GetPlayerAccess(index) >= GetPlayerAccess(n) Then ' And GetPlayerAccess(n) <= 0 Then
'                If GetPlayerLevel(n) + 5 >= GetPlayerLevel(Index) Then
'                    If GetPlayerLevel(n) - 5 <= GetPlayerLevel(Index) Then
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", RGB_AlertColor)
                
                        Select Case Spell(SpellNum).type
                            Case SPELL_TYPE_SUBHP
                                If map(GetPlayerMap(index)).Moral <> MAP_MORAL_SAFE Then
                                    Damage = (Int(GetPlayerINT(index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(n)
                                    If Damage > 0 Then
                                        Call AttackPlayer(index, n, Damage)
                                    Else
                                        Call PlayerMsg(index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", RGB_AlertColor)
                                    End If
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
            
            'send sound
            Call SendSound(index, Trim(Spell(SpellNum).sound), True)
                ' Take away the mana points
                Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
                Call SendMP(index)
                Casted = True
            Else
                If GetPlayerMap(index) = GetPlayerMap(n) And Spell(SpellNum).type >= SPELL_TYPE_ADDHP And Spell(SpellNum).type <= SPELL_TYPE_ADDSP Then
                    Select Case Spell(SpellNum).type
                    
                        Case SPELL_TYPE_ADDHP
                            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", RGB_AlertColor)
                            Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                            Call SendHP(n)
                                    
                        Case SPELL_TYPE_ADDMP
                            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", RGB_AlertColor)
                            Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                            Call SendMP(n)
                    
                        Case SPELL_TYPE_ADDSP
                            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", RGB_AlertColor)
                            Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                            Call SendMP(n)
                    End Select
                    
                    'send sound
            Call SendSound(index, Trim(Spell(SpellNum).sound), True)
                    ' Take away the mana points
                    Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
                    Call SendMP(index)
                    Casted = True
                Else
                    Call PlayerMsg(index, "Could not cast spell!", RGB_AlertColor)
                End If
            End If
        Else
            Call PlayerMsg(index, "Could not cast spell!", RGB_AlertColor)
        End If
    Else
        If Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(index), n).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on a " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", RGB_AlertColor)
            
            Select Case Spell(SpellNum).type
                Case SPELL_TYPE_ADDHP
                    MapNpc(GetPlayerMap(index), n).HP = MapNpc(GetPlayerMap(index), n).HP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBHP
                    
                    Damage = (Int(GetPlayerINT(index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(index), n).num).def / 2)
                    If Damage > 0 Then
                        Call AttackNpc(index, n, Damage)
                    Else
                        Call PlayerMsg(index, "The spell was to weak to hurt " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & "!", RGB_AlertColor)
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
        
        'send sound
            Call SendSound(index, Trim(Spell(SpellNum).sound), True)
            ' Take away the mana points
            Call SetPlayerMP(index, GetPlayerMP(index) - MPReq)
            Call SendMP(index)
            Casted = True
        Else
            Call PlayerMsg(index, "Could not cast spell!", RGB_AlertColor)
        End If
    End If

    If Casted = True Then
        player(index).AttackTimer = GetTickCount
        player(index).CastedSpell = YES
    End If
End Sub

Function GetSpellReqLevel(ByVal index As Long, ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).Data1 - Int(GetClassINT(GetPlayerClass(index)) / 4)
End Function

Function GetPrayerReqLevel(ByVal index As Long, ByVal PrayerNum As Long)
    GetPrayerReqLevel = Prayer(PrayerNum).Data1 - Int(GetClassINT(GetPlayerClass(index)) / 4)
End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
Dim i As Long, n As Long

    CanPlayerCriticalHit = False
    
    If GetPlayerWeaponSlot(index) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerDEX(index) / 2) + Int(GetPlayerLevel(index) / 2)
    
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
            i = Int(GetPlayerDEX(index) / 2) + Int(GetPlayerLevel(index) / 2)
        
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Sub CheckEquippedItems(ByVal index As Long)
Dim Slot As Long, itemnum As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    Slot = GetPlayerWeaponSlot(index)
    If Slot > 0 Then
        itemnum = GetPlayerInvItemNum(index, Slot)
        
        If itemnum > 0 Then
            If Item(itemnum).type <> ITEM_TYPE_WEAPON Then
                Call SetPlayerWeaponSlot(index, 0)
            End If
        Else
            Call SetPlayerWeaponSlot(index, 0)
        End If
    End If

    Slot = GetPlayerArmorSlot(index)
    If Slot > 0 Then
        itemnum = GetPlayerInvItemNum(index, Slot)
        
        If itemnum > 0 Then
            If Item(itemnum).type <> ITEM_TYPE_ARMOR Then
                Call SetPlayerArmorSlot(index, 0)
            End If
        Else
            Call SetPlayerArmorSlot(index, 0)
        End If
    End If

    Slot = GetPlayerHelmetSlot(index)
    If Slot > 0 Then
        itemnum = GetPlayerInvItemNum(index, Slot)
        
        If itemnum > 0 Then
            If Item(itemnum).type <> ITEM_TYPE_HELMET Then
                Call SetPlayerHelmetSlot(index, 0)
            End If
        Else
            Call SetPlayerHelmetSlot(index, 0)
        End If
    End If

    Slot = GetPlayerShieldSlot(index)
    If Slot > 0 Then
        itemnum = GetPlayerInvItemNum(index, Slot)
        
        If itemnum > 0 Then
            If Item(itemnum).type <> ITEM_TYPE_SHIELD Then
                Call SetPlayerShieldSlot(index, 0)
            End If
        Else
            Call SetPlayerShieldSlot(index, 0)
        End If
    End If
End Sub












'~~~~~~~~~~~~~~~~#
Sub CastPrayer(ByVal index As Long, ByVal PrayerSlot As Long)
Dim PrayerNum As Long, PPReq As Long, i As Long, n As Long, Damage As Long
Dim Casted As Boolean

    Casted = False
    
    ' Prevent subscript out of range
    If PrayerSlot <= 0 Or PrayerSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    PrayerNum = GetPlayerPrayer(index, PrayerSlot)
    
    ' Make sure player has the spell
    If Not HasPrayer(index, PrayerNum) Then
        Call PlayerMsg(index, "You do not have this prayer!", RGB_AlertColor)
        Exit Sub
    End If
    
    i = GetPrayerReqLevel(index, PrayerNum)
    PPReq = (i + Prayer(PrayerNum).Data1 + Prayer(PrayerNum).Data2 + Prayer(PrayerNum).Data3) / 10
    
    ' Check if they have enough PP
    If GetPlayerPP(index) < PPReq Then
        Call PlayerMsg(index, "Not enough prayer points!", RGB_AlertColor)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If i > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & i & " to cast this prayer.", RGB_AlertColor)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < player(index).AttackTimer + 1000 Then
        Exit Sub
    End If
        '
        'PRAYER_TYPE_HEAL
        'PRAYER_TYPE_CURE
        'PRAYER_TYPE_ENHANCE
    n = player(index).target
    
    If player(index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
                If GetPlayerMap(index) = GetPlayerMap(n) And Prayer(PrayerNum).type >= PRAYER_TYPE_HEAL And Prayer(PrayerNum).type <= PRAYER_TYPE_ENHANCE Then
                    Select Case Prayer(PrayerNum).type
                    
                        Case PRAYER_TYPE_HEAL
                            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Prayer(PrayerNum).Name), RGB_AlertColor) '& Trim(Spell(SpellNum).name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerHP(n, GetPlayerHP(n) + Prayer(PrayerNum).Data1)
                            Call SendHP(n)
                                    
                        Case PRAYER_TYPE_CURE
                            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Prayer(PrayerNum).Name), RGB_AlertColor) '& Trim(Spell(SpellNum).name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call setPlayerPoison(index, False, 0, 0)
                            Call PlayerMsg(index, "You have been cured!", RGB_AlertColor)
                            'Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                            'Call SendMP(n)
                    
                        Case PRAYER_TYPE_ENHANCE
                            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Prayer(PrayerNum).Name), RGB_AlertColor) '& Trim(Spell(SpellNum).name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call PlayerMsg(index, "You boost your ego :)!", RGB_AlertColor)
                            'Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                            'Call SendMP(n)
                        Case PRAYER_TYPE_BOOST
                            
                    End Select
                    
                    'send sound
            Call SendSound(index, Trim(Prayer(PrayerNum).sound), True)
                    ' Take away the mana points
                    Call SetPlayerPP(index, GetPlayerPP(index) - PPReq)
                    Call SendPP(index)
                    Casted = True
                Else
                    Call PlayerMsg(index, "Could not cast prayer!", RGB_AlertColor)
                End If
            
        Else
            Call PlayerMsg(index, "Could not cast prayer!", RGB_AlertColor)
        End If
    Else
        Call PlayerMsg(index, "Could not cast prayer!", RGB_AlertColor)
    End If

    If Casted = True Then
        player(index).AttackTimer = GetTickCount
        player(index).CastedSpell = YES
    End If
End Sub


Function getItemNo(ByVal Name As String) As String
Dim i As Long
    For i = 1 To MAX_ITEMS Step 1
        If Trim(LCase(Item(i).Name)) = Trim(LCase(Name)) Then
            getItemNo = i
            Exit Function
        End If
    Next i
    getItemNo = 0
End Function
