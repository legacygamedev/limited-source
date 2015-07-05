Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal Index As Long) As Long
Dim WeaponSlot As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    GetPlayerDamage = Int(GetPlayerSTR(Index) / 2)
    GetPlayerDamage = GetPlayerDamage + (Rnd * (GetPlayerLUCK(Index)) / 10 - GetPlayerLUCK(Index) / 20)
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(Index)
        
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2
        
        If GetPlayerInvItemDur(Index, WeaponSlot) > -1 And Item(GetPlayerInvItemNum(Index, WeaponSlot)).Type <> ITEM_TYPE_LAMP Then
            Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) - 1)
        
            If GetPlayerInvItemDur(Index, WeaponSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, WeaponSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, WeaponSlot) & "/" & Trim(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If GetPlayerDamage < 0 Then
        GetPlayerDamage = 0
    End If
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long, WeponSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    WeponSlot = GetPlayerWeaponSlot(Index)
    ArmorSlot = GetPlayerArmorSlot(Index)
    HelmSlot = GetPlayerHelmetSlot(Index)
    ShieldSlot = GetPlayerShieldSlot(Index)
    GetPlayerProtection = Int(GetPlayerDEF(Index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data2
        If GetPlayerInvItemDur(Index, ArmorSlot) > -1 And Item(GetPlayerInvItemNum(Index, WeponSlot)).Type <> ITEM_TYPE_LAMP Then
            Call SetPlayerInvItemDur(Index, ArmorSlot, GetPlayerInvItemDur(Index, ArmorSlot) - 1)
        
            If GetPlayerInvItemDur(Index, ArmorSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ArmorSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, ArmorSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, ArmorSlot) & "/" & Trim(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, HelmSlot)).Data2
        If GetPlayerInvItemDur(Index, HelmSlot) > -1 And Item(GetPlayerInvItemNum(Index, WeponSlot)).Type <> ITEM_TYPE_LAMP Then
            Call SetPlayerInvItemDur(Index, HelmSlot, GetPlayerInvItemDur(Index, HelmSlot) - 1)

            If GetPlayerInvItemDur(Index, HelmSlot) <= 0 Then
                Call BattleMsg(Index, "Your " & Trim(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, HelmSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, HelmSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, HelmSlot) & "/" & Trim(Item(GetPlayerInvItemNum(Index, HelmSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data2
        If GetPlayerInvItemDur(Index, ShieldSlot) > -1 And Item(GetPlayerInvItemNum(Index, WeponSlot)).Type <> ITEM_TYPE_LAMP Then
            Call SetPlayerInvItemDur(Index, ShieldSlot, GetPlayerInvItemDur(Index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(Index, ShieldSlot) <= 0 Then
                Call BattleMsg(Index, "Your " & Trim(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ShieldSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, ShieldSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, ShieldSlot) & "/" & Trim(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
End Function

Function FindOpenPlayerSlot() As Long
Dim I As Long

    FindOpenPlayerSlot = 0
    
    For I = 1 To MAX_PLAYERS
        If Not IsConnected(I) Then
            FindOpenPlayerSlot = I
            Exit Function
        End If
    Next I
End Function

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim I As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, I) = ItemNum Then
                FindOpenInvSlot = I
                Exit Function
            End If
        Next I
    End If
    
    For I = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, I) = 0 Then
            FindOpenInvSlot = I
            Exit Function
        End If
    Next I
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
Dim I As Long

    FindOpenMapItemSlot = 0
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If
    
    For I = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, I).num = 0 Then
            FindOpenMapItemSlot = I
            Exit Function
        End If
    Next I
End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
Dim I As Long

    FindOpenSpellSlot = 0
    
    For I = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, I) = 0 Then
            FindOpenSpellSlot = I
            Exit Function
        End If
    Next I
End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
Dim I As Long

    HasSpell = False
    
    For I = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, I) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next I
End Function
Function TotalOnlinePlayers() As Long
Dim I As Long
TotalOnlinePlayers = 0

For I = 1 To MAX_PLAYERS
    If IsPlaying(I) Then
        TotalOnlinePlayers = TotalOnlinePlayers + 1
    End If
Next I
End Function

Function FindPlayer(ByVal Name As String) As Long
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(I)) >= Len(Trim(Name)) Then
                If UCase(Mid(GetPlayerName(I), 1, Len(Trim(Name)))) = UCase(Trim(Name)) Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I
    
    FindPlayer = 0
End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim I As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    For I = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, I)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next I
End Function

Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim I As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For I = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, I) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) - ItemVal)
                    Call SendInventoryUpdate(Index, I)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, I)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If I = GetPlayerWeaponSlot(Index) Then
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
                        
                    Case ITEM_TYPE_LAMP
                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If I = GetPlayerWeaponSlot(Index) Then
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
                            If I = GetPlayerArmorSlot(Index) Then
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
                            If I = GetPlayerHelmetSlot(Index) Then
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
                            If I = GetPlayerShieldSlot(Index) Then
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
                End Select

                
                n = Item(GetPlayerInvItemNum(Index, I)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_LAMP) Then
                    TakeItem = True
                End If
            End If
                            
            If TakeItem = True Then
                Call SetPlayerInvItemNum(Index, I, 0)
                Call SetPlayerInvItemValue(Index, I, 0)
                Call SetPlayerInvItemDur(Index, I, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(Index, I)
                Exit Sub
            End If
        End If
    Next I
End Sub

Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim I As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    I = FindOpenInvSlot(Index, ItemNum)
    
    ' Check to see if inventory is full
    If I <> 0 Then
        Call SetPlayerInvItemNum(Index, I, ItemNum)
        Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(Index, I, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(Index, I)
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
    End If
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim I As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    I = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(I, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, X, Y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim I As Long
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    I = MapItemSlot
    
    If I <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
        MapItem(MapNum, I).num = ItemNum
        MapItem(MapNum, I).Value = ItemVal
        
        If ItemNum <> 0 Then
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                MapItem(MapNum, I).Dur = ItemDur
            Else
                MapItem(MapNum, I).Dur = 0
            End If
        Else
            MapItem(MapNum, I).Dur = 0
        End If
        
        MapItem(MapNum, I).X = X
        MapItem(MapNum, I).Y = Y
            
        Packet = "SPAWNITEM" & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
End Sub

Sub SpawnAllMapsItems()
Dim I As Long
    
    For I = 1 To MAX_MAPS
        Call SpawnMapItems(I)
    Next I
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
Dim X As Long
Dim Y As Long
Dim I As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Spawn what we have
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            ' Check if the tile type is an item or a saved tile incase someone drops something
            
            If (Map(MapNum).Tile(X, Y).Type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(X, Y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).Tile(X, Y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, 1, MapNum, X, Y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, Map(MapNum).Tile(X, Y).Data2, MapNum, X, Y)
                End If
            End If
        Next X
    Next Y
End Sub

Sub PlayerMapGetItem(ByVal Index As Long)
Dim I As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String

    If IsPlaying(Index) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Index)
    
    For I = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, I).num > 0) And (MapItem(MapNum, I).num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, I).X = GetPlayerX(Index)) And (MapItem(MapNum, I).Y = GetPlayerY(Index)) Then
                ' Find open slot
                n = FindOpenInvSlot(Index, MapItem(MapNum, I).num)
                
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, I).num)
                    If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Then
                        Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, I).Value)
                        Msg = "You picked up " & MapItem(MapNum, I).Value & " " & Trim(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(Index, n, 0)
                        Msg = "You picked up a " & Trim(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
                    End If
                    Call SetPlayerInvItemDur(Index, n, MapItem(MapNum, I).Dur)
                        
                    ' Erase item from the map
                    MapItem(MapNum, I).num = 0
                    MapItem(MapNum, I).Value = 0
                    MapItem(MapNum, I).Dur = 0
                    MapItem(MapNum, I).X = 0
                    MapItem(MapNum, I).Y = 0
                        
                    Call SendInventoryUpdate(Index, n)
                    Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                    Call PlayerMsg(Index, Msg, Yellow)
                    Exit Sub
                Else
                    Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                    Exit Sub
                End If
            End If
        End If
    Next I
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
Dim I As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        I = FindOpenMapItemSlot(GetPlayerMap(Index))
        
        If I <> 0 Then
            MapItem(GetPlayerMap(Index), I).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(Index) Then
                        Call SetPlayerArmorSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(Index) Then
                        Call SetPlayerWeaponSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(Index) Then
                        Call SetPlayerHelmetSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(Index) Then
                        Call SetPlayerShieldSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(Index), I).num = GetPlayerInvItemNum(Index, InvNum)
            MapItem(GetPlayerMap(Index), I).X = GetPlayerX(Index)
            MapItem(GetPlayerMap(Index), I).Y = GetPlayerY(Index)
                        
            If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                    MapItem(GetPlayerMap(Index), I).Value = GetPlayerInvItemValue(Index, InvNum)
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(Index), I).Value = Amount
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Amount & " " & Trim(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(Index), I).Value = 0
                If Item(GetPlayerInvItemNum(Index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                    If Item(GetPlayerInvItemNum(Index, InvNum)).Data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - Ind.", Yellow)
                    Else
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - " & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Data1 & ".", Yellow)
                    End If
                Else
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(Index, InvNum, 0)
                Call SetPlayerInvItemValue(Index, InvNum, 0)
                Call SetPlayerInvItemDur(Index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(Index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(I, MapItem(GetPlayerMap(Index), I).num, Amount, MapItem(GetPlayerMap(Index), I).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
        Else
            Call PlayerMsg(Index, "To many items already on the ground.", BrightRed)
        End If
    End If
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Packet As String
Dim NpcNum As Long
Dim I As Long, X As Long, Y As Long
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
        
        'Check if theres a spawn tile for the specific npc
        For X = 0 To MAX_MAPX
              For Y = 0 To MAX_MAPY
                  If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_NPCSPAWN Then
                       If Map(MapNum).Tile(X, Y).Data1 = MapNpcNum Then
                           MapNpc(MapNum, MapNpcNum).X = X
                           MapNpc(MapNum, MapNpcNum).Y = Y
                           Spawned = True
                           Exit For
                       End If
                  End If
              Next Y
        Next X
        
        ' Well try 100 times to randomly place the sprite
        If Not Spawned Then
            For I = 1 To 100
                X = Int(Rnd * MAX_MAPX)
                Y = Int(Rnd * MAX_MAPY)
            
                ' Check if the tile is walkable
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                    MapNpc(MapNum, MapNpcNum).X = X
                    MapNpc(MapNum, MapNpcNum).Y = Y
                    Spawned = True
                    Exit For
                End If
            Next I
        End If
        
        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                        MapNpc(MapNum, MapNpcNum).X = X
                        MapNpc(MapNum, MapNpcNum).Y = Y
                        Spawned = True
                    End If
                Next X
            Next Y
        End If
             
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        Else
            MapNpc(MapNum, MapNpcNum).num = 0
            MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
            MapNpc(MapNum, MapNpcNum).HP = 0
            Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
    End If
    
    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim I As Long

    For I = 1 To MAX_MAP_NPCS
        Call SpawnNpc(I, MapNum)
    Next I
End Sub

Sub SpawnAllMapNpcs()
Dim I As Long

    For I = 1 To MAX_MAPS
        Call SpawnMapNpcs(I)
    Next I
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        If Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).Type <> ITEM_TYPE_LAMP Then
            AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
        Else
            AttackSpeed = 1000
        End If
    Else
        AttackSpeed = 1000
    End If
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
    If (GetPlayerMap(Attacker) = GetPlayerMap(Victim)) And (GetTickCount > Player(Attacker).AttackTimer + AttackSpeed) Then
        
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
                                Call PlayerMsg(Attacker, "You are below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack another player yet!", BrightRed)
                            Else
                                If GetPlayerLevel(Victim) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack this player yet!", BrightRed)
                                Else
                                    If Trim(GetPlayerGuild(Attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                        If Trim(GetPlayerGuild(Attacker)) <> Trim(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                            Call PlayerMsg(Attacker, "You cant attack a guild member!", BrightRed)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If

            Case DIR_DOWN
                If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
                                Call PlayerMsg(Attacker, "You are below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack another player yet!", BrightRed)
                            Else
                                If GetPlayerLevel(Victim) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack this player yet!", BrightRed)
                                Else
                                    If Trim(GetPlayerGuild(Attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                        If Trim(GetPlayerGuild(Attacker)) <> Trim(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                            Call PlayerMsg(Attacker, "You cant attack a guild member!", BrightRed)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
        
            Case DIR_LEFT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
                                Call PlayerMsg(Attacker, "You are below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack another player yet!", BrightRed)
                            Else
                                If GetPlayerLevel(Victim) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack this player yet!", BrightRed)
                                Else
                                    If Trim(GetPlayerGuild(Attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                        If Trim(GetPlayerGuild(Attacker)) <> Trim(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                            Call PlayerMsg(Attacker, "You cant attack a guild member!", BrightRed)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
            
            Case DIR_RIGHT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
                                Call PlayerMsg(Attacker, "You are below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack another player yet!", BrightRed)
                            Else
                                If GetPlayerLevel(Victim) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack this player yet!", BrightRed)
                                Else
                                    If Trim(GetPlayerGuild(Attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                        If Trim(GetPlayerGuild(Attacker)) <> Trim(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                            Call PlayerMsg(Attacker, "You cant attack a guild member!", BrightRed)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
        End Select
    End If
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        If Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).Type <> ITEM_TYPE_LAMP Then
            AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
        Else
            AttackSpeed = 1000
        End If
    Else
        AttackSpeed = 1000
    End If

CanAttackNpc = False
 
' Check for subscript out of range
If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
Exit Function
End If
 
' Check for subscript out of range
If MapNpc(GetPlayerMap(Attacker), MapNpcNum).num <= 0 Then
Exit Function
End If
 
MapNum = GetPlayerMap(Attacker)
NpcNum = MapNpc(MapNum, MapNpcNum).num
 
' Make sure the npc isn't already dead
If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
Exit Function
End If
 
' Make sure they are on the same map
If IsPlaying(Attacker) Then
    If NpcNum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (MapNpc(MapNum, MapNpcNum).Y + 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
 
            Case DIR_DOWN
                If (MapNpc(MapNum, MapNpcNum).Y - 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
 
            Case DIR_LEFT
                If (MapNpc(MapNum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X + 1 = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
 
            Case DIR_RIGHT
                If (MapNpc(MapNum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X - 1 = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
        End Select
    End If
End If
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
    
    CanNpcAttackPlayer = False
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Index) = False Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index), MapNpcNum).num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Index)
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
    If Player(Index).GettingMap = YES Then
        Exit Function
    End If
    
    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NpcNum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum, MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum, MapNpcNum).X) Then
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
Dim I As Long

    '  Removes one SP when you attack, you can also set it to 2 or 3.
    If GetPlayerSP(Attacker) > 0 Then
    Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - 1)
    Call SendSP(Attacker)
    End If

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

If Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA Then
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        'Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
        'Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)
    
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        
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
            Exp = Int(GetPlayerExp(Victim) / 10)
            
            ' Make sure we dont get less then 0
            If Exp < 0 Then
                Exp = 0
            End If
            
            If GetPlayerLevel(Victim) = MAX_LEVEL Then
                Call BattleMsg(Victim, "You cant lose any experience!", BrightRed, 1)
                Call BattleMsg(Attacker, GetPlayerName(Victim) & " is the max level!", BrightBlue, 0)
            Else
                If Exp = 0 Then
                    Call BattleMsg(Victim, "You lost no experience.", BrightRed, 1)
                    Call BattleMsg(Attacker, "You received no experience.", BrightBlue, 0)
                Else
                    Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                    Call BattleMsg(Victim, "You lost " & Exp & " experience.", BrightRed, 1)
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                    Call BattleMsg(Attacker, "You got " & Exp & " experience for killing " & GetPlayerName(Victim) & ".", BrightBlue, 0)
                End If
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
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!", BrightRed)
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
        'Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
        'Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)
        If n = 0 Then
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
    End If
ElseIf Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        'Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
        'Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)
        If n = 0 Then
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
            'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
    
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed in the arena by " & GetPlayerName(Attacker), BrightRed)
            
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
        If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
    Else
        ' Player not dead, just do the damage
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
        
        ' Check for a weapon and say damage
        'Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
        'Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)
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
    Player(Attacker).AttackTimer = GetTickCount
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
        'Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)
        
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
        'Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)
        
        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
    End If
    
    Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & END_CHAR)
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim n As Long, I As Long, q As Integer, X As Long
Dim STR As Long, DEF As Long, MapNum As Long, NpcNum As Long

    ' Removes one SP when you attack, you can also set it to 2 or 3.
    'If GetPlayerSP(Attacker) > 0 Then
    'Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - 1)
    'Call SendSP(Attacker)
    'End If

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
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim(Npc(NpcNum).Name)
        
    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        
        If Trim(Name) = vbNullString Then Exit Sub
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
            For X = 1 To MAX_PARTY_MEMBERS
                If Player(Attacker).Party.Member(X) > 0 Then q = q + 1
            Next X
            
            'MsgBox "in party" & q
            If q >= 2 Then 'Remember, if they aren't in a party they'll only get one person, so this has to be at least two
                Exp = Exp * (1 / q)  ' 3/4 experience
                'MsgBox Exp
                For X = 1 To MAX_PARTY_MEMBERS
                    If Player(Attacker).Party.Member(X) > 0 Then
                        If Player(Player(Attacker).Party.Member(X)).Party.ShareExp = True Then
                            If GetPlayerLevel(Player(Attacker).Party.Member(X)) = MAX_LEVEL Then
                                Call SetPlayerExp(Player(Attacker).Party.Member(X), Experience(MAX_LEVEL))
                                Call BattleMsg(Player(Attacker).Party.Member(X), "You cant gain anymore experience!", BrightBlue, 0)
                            Else
                                Call SetPlayerExp(Player(Attacker).Party.Member(X), GetPlayerExp(Player(Attacker).Party.Member(X)) + Exp)
                                Call BattleMsg(Player(Attacker).Party.Member(X), "You have gained " & Exp & " party experience.", BrightBlue, 0)
                            End If
                        End If
                    End If
                Next X
            End If
        End If
                                
        For I = 1 To MAX_NPC_DROPS
            ' Drop the goods if they get it
            n = Int(Rnd * Npc(NpcNum).ItemNPC(I).Chance) + 1
            If n = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(I).ItemNum, Npc(NpcNum).ItemNPC(I).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
            End If
        Next I
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
       
        ' Check for level up party member
        If Player(Attacker).InParty = YES Then
            For X = 1 To MAX_PARTY_MEMBERS
                Call CheckPlayerLevelUp(Player(Attacker).Party.Member(X))
            Next X
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
        MapNpc(MapNum, MapNpcNum).HP = MapNpc(MapNum, MapNpcNum).HP - Damage
        
        ' Check for a weapon and say damage
        'Call BattleMsg(Attacker, "You hit a " & Name & " for " & Damage & " damage.", White, 0)
        
        If n = 0 Then
            'Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        Else
            'Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If
        
        ' Check if we should send a message
        If MapNpc(MapNum, MapNpcNum).Target = 0 And MapNpc(MapNum, MapNpcNum).Target <> Attacker Then
            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay) & "", SayColor)
            End If
        End If
        
        ' Set the NPC target to the player
        MapNpc(MapNum, MapNpcNum).Target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum, I).num = MapNpc(MapNum, MapNpcNum).num Then
                    MapNpc(MapNum, I).Target = Attacker
                End If
            Next I
        End If
    End If
        
    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)
        
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    'If Trim(Shop(ShopNum).LeaveSay) <> "" Then
        'Call PlayerMsg(Index, Trim(Shop(ShopNum).Name) & " : " & Trim(Shop(ShopNum).LeaveSay) & "", SayColor)
    'End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    Call SendLeaveMap(Index, OldMap)
    
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)
                
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Player(Index).GettingMap = YES
    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "warp" & SEP_CHAR & END_CHAR)
    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)
    
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim MapNum As Long
Dim X As Long
Dim Y As Long
Dim I As Long
Dim Moved As Byte

    ' They tried to hack
    'If Moved = NO Then
        'Call HackingAttempt(index, "Position Modification")
        'Exit Sub
    'End If
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    Call SetPlayerDir(Index, Dir)
    
    Moved = NO
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKICON And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKICON And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKICON Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), MAX_MAPY)
                    Moved = YES
                End If
            End If
                    
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < MAX_MAPY Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKICON Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Down > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    Moved = YES
                End If
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKICON Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, MAX_MAPX, GetPlayerY(Index))
                    Moved = YES
                End If
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < MAX_MAPX Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKICON Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Right > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Moved = YES
                End If
            End If
    End Select
    
    If GetPlayerX(Index) < 0 Or GetPlayerY(Index) < 0 Or GetPlayerX(Index) > MAX_MAPX Or GetPlayerY(Index) > MAX_MAPY Or GetPlayerMap(Index) <= 0 Then
        Call HackingAttempt(Index, "Position Modification")
        Exit Sub
    End If
    
    'healing tiles code
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SendHP(Index)
        Call SendMP(Index)
        Call PlayerMsg(Index, "You feel a sudden rush through your body as you regain strength!", BrightGreen)
    End If
    
    'Check for kill tile, and if so kill them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KILL Then
        Call SetPlayerHP(Index, 0)
        Call PlayerMsg(Index, "You embrace the cold finger of death and feel your life extinguished.", BrightRed)
        
        ' Warp player away
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Index
        Else
            Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        End If
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
        Call SendHP(Index)
        Call SendMP(Index)
        Call SendSP(Index)
        Moved = YES
    End If

    If GetPlayerX(Index) + 1 <= MAX_MAPX Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)
            
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerX(Index) - 1 >= 0 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)
            
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(Index) - 1 >= 0 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1
            
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(Index) + 1 <= MAX_MAPY Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1
            
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
            
    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3
                        
        Call PlayerWarp(Index, MapNum, X, Y)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KEYOPEN Then
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        
        If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
            TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            If Trim(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = "" Then
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", White)
            Else
                Call MapMsg(GetPlayerMap(Index), Trim(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
            End If
            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
        End If
    End If
        
    ' Check for shop
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SHOP Then
       If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 > 0 Then
            Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
        Else
            Call PlayerMsg(Index, "There is no shop here.", BrightRed)
        End If
    End If
        
    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
            Call PlayerMsg(Index, "You already have this sprite!", BrightRed)
            Exit Sub
        Else
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 = 0 Then
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Else
                If Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(Index, "This sprite will cost you " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 & " " & Trim(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", Yellow)
                Else
                    Call PlayerMsg(Index, "This sprite will cost you a " & Trim(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", Yellow)
                End If
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    
    ' Check if player stepped on a class changing tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 > -1 Then
            If GetPlayerClass(Index) <> Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
                Call PlayerMsg(Index, "You arent the required class!", BrightRed)
                Exit Sub
            End If
        End If
        
        If GetPlayerClass(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
            Call PlayerMsg(Index, "You are already this class!", BrightRed)
        Else
            If Player(Index).Char(Player(Index).CharNum).Sex = 0 Then
                If GetPlayerSprite(Index) = Class(GetPlayerClass(Index)).MaleSprite Then
                    Call SetPlayerSprite(Index, Class(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).MaleSprite)
                End If
            Else
                If GetPlayerSprite(Index) = Class(GetPlayerClass(Index)).FemaleSprite Then
                    Call SetPlayerSprite(Index, Class(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).FemaleSprite)
                End If
            End If
            
            Call SetPlayerSTR(Index, (Player(Index).Char(Player(Index).CharNum).STR - Class(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF - Class(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi - Class(GetPlayerClass(Index)).Magi))
            Call SetPlayerLUCK(Index, (Player(Index).Char(Player(Index).CharNum).Luck - Class(GetPlayerClass(Index)).Luck))
            
            Call SetPlayerClass(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)

            Call SetPlayerSTR(Index, (Player(Index).Char(Player(Index).CharNum).STR + Class(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF + Class(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi + Class(GetPlayerClass(Index)).Magi))
            Call SetPlayerLUCK(Index, (Player(Index).Char(Player(Index).CharNum).Luck + Class(GetPlayerClass(Index)).Luck))
            
            
            Call PlayerMsg(Index, "Your new class is a " & Trim(Class(GetPlayerClass(Index)).Name) & "!", BrightGreen)
            
            Call SendStats(Index)
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & END_CHAR)
        End If
    End If
    
    ' Check if player stepped on notice tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_NOTICE Then
        If Trim(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> "" Then
            Call PlayerMsg(Index, Trim(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), Black)
        End If
        If Trim(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2) <> "" Then
            Call PlayerMsg(Index, Trim(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2), Grey)
        End If
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "panned" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 & END_CHAR)
    End If
    
    ' Check if player stepped on sound tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1 & SEP_CHAR & END_CHAR)
    End If
    
    'Check if the player stepped on a Scripted Tile
    If Scripting = 1 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        End If
    End If
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
Dim I As Long, n As Long
Dim X As Long, Y As Long
Dim BX As Long, BY As Long

    CanNpcMove = False
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    
    X = MapNpc(MapNum, MapNpcNum).X
    Y = MapNpc(MapNum, MapNpcNum).Y
    
    CanNpcMove = True
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = Map(MapNum).Tile(X, Y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum, MapNpcNum).X) And (GetPlayerY(I) = MapNpc(MapNum, MapNpcNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next I
                                
                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> MapNpcNum) And (MapNpc(MapNum, I).num > 0) And (MapNpc(MapNum, I).X = MapNpc(MapNum, MapNpcNum).X) And (MapNpc(MapNum, I).Y = MapNpc(MapNum, MapNpcNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next I
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If Y < MAX_MAPY Then
                n = Map(MapNum).Tile(X, Y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum, MapNpcNum).X) And (GetPlayerY(I) = MapNpc(MapNum, MapNpcNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next I
                
                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> MapNpcNum) And (MapNpc(MapNum, I).num > 0) And (MapNpc(MapNum, I).X = MapNpc(MapNum, MapNpcNum).X) And (MapNpc(MapNum, I).Y = MapNpc(MapNum, MapNpcNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next I
            Else
                CanNpcMove = False
            End If
                
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum, MapNpcNum).X - 1) And (GetPlayerY(I) = MapNpc(MapNum, MapNpcNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next I
                
                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> MapNpcNum) And (MapNpc(MapNum, I).num > 0) And (MapNpc(MapNum, I).X = MapNpc(MapNum, MapNpcNum).X - 1) And (MapNpc(MapNum, I).Y = MapNpc(MapNum, MapNpcNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next I
            Else
                CanNpcMove = False
            End If
                
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If X < MAX_MAPX Then
                n = Map(MapNum).Tile(X + 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum, MapNpcNum).X + 1) And (GetPlayerY(I) = MapNpc(MapNum, MapNpcNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next I
                
                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> MapNpcNum) And (MapNpc(MapNum, I).num > 0) And (MapNpc(MapNum, I).X = MapNpc(MapNum, MapNpcNum).X + 1) And (MapNpc(MapNum, I).Y = MapNpc(MapNum, MapNpcNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next I
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim X As Long
Dim Y As Long
Dim I As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_DOWN
            MapNpc(MapNum, MapNpcNum).Y = MapNpc(MapNum, MapNpcNum).Y + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_LEFT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X - 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case DIR_RIGHT
            MapNpc(MapNum, MapNpcNum).X = MapNpc(MapNum, MapNpcNum).X + 1
            Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
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
Dim MOTD As String
Dim f As Long

    ' Set the flag so we know the person is in the game
    Player(Index).InGame = True
    
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(Index, "LOGINOK" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
    
    ReDim Player(Index).Party.Member(1 To MAX_PARTY_MEMBERS)
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendEmoticons(Index)
    Call SendArrows(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendStats(Index)
    Call SendWeatherTo(Index)
    Call SendTimeTo(Index)
    Call SendGameClockTo(Index)
    Call SendOnlineList
    Call SendGameClockTo(Index)
    Call SendNewsTo(Index)
    Call DisabledTimeTo(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    Call SendPlayerData(Index)
    
    If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "JoinGame " & Index
    Else
        MOTD = GetVar("motd.ini", "MOTD", "Msg")
        
        ' Send a global message that he/she joined
        If GetPlayerAccess(Index) <= ADMIN_MONITER Then
            Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", 7)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", 15)
        End If
    
        ' Send them welcome
        Call PlayerMsg(Index, "Welcome to " & GAME_NAME & "!", 15)
        
        ' Send motd
        If Trim(MOTD) <> "" Then
            Call PlayerMsg(Index, "MOTD: " & MOTD, 11)
        End If
    End If
    
    ' Send whos online
    Call SendWhosOnline(Index)
    Call ShowPLR(Index)

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

        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        Call RemovePMember(Index)
        
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "LeftGame " & Index
        Else
            ' Check for boot map
            If Map(GetPlayerMap(Index)).BootMap > 0 Then
                Call SetPlayerX(Index, Map(GetPlayerMap(Index)).BootX)
                Call SetPlayerY(Index, Map(GetPlayerMap(Index)).BootY)
                Call SetPlayerMap(Index, Map(GetPlayerMap(Index)).BootMap)
            End If
                  
            ' Send a global message that he/she left
            If GetPlayerAccess(Index) <= 1 Then
                Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", 7)
            Else
                Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", 15)
            End If
        End If
        
        Call SavePlayer(Index)
        
        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".", True)
        Call SendLeftGame(Index)
        Call RemovePLR
        For n = 1 To MAX_PLAYERS
            Call ShowPLR(n)
        Next n
    End If
    
    Call ClearPlayer(Index)
    Call SendOnlineList
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
Dim I As Long, n As Long

    n = 0
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
            n = n + 1
        End If
    Next I
    
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
        
    GetNpcMaxSP = Npc(NpcNum).Luck * 2
End Function

Function GetPlayerHPRegen(ByVal Index As Long)
Dim I As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
            GetPlayerHPRegen = 0
            Exit Function
        End If
        
        I = Int(GetPlayerDEF(Index) / 2)
        If I < 2 Then I = 2
        
        GetPlayerHPRegen = I
    End If
End Function

Function GetPlayerMPRegen(ByVal Index As Long)
Dim I As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
            GetPlayerMPRegen = 0
            Exit Function
        End If
        
        I = Int(GetPlayerMAGI(Index) / 2)
        If I < 2 Then I = 2
        
        GetPlayerMPRegen = I
    End If
End Function

Function GetPlayerSPRegen(ByVal Index As Long)
Dim I As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
            GetPlayerSPRegen = 0
            Exit Function
        End If
        
        I = Int(GetPlayerLUCK(Index) / 2)
        If I < 2 Then I = 2
        
        GetPlayerSPRegen = I
    End If
End Function

Function GetNpcHPRegen(ByVal NpcNum As Long)
Dim I As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcHPRegen = 0
        Exit Function
    End If
    
    I = Int(Npc(NpcNum).DEF / 3)
    If I < 1 Then I = 1
    
    GetNpcHPRegen = I
End Function

Sub CheckPlayerLevelUp(ByVal Index As Long)
Dim I As Long
Dim d As Long
Dim c As Long
    c = 0
    
    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        If GetPlayerLevel(Index) < MAX_LEVEL Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerLevelUp " & Index
            Else
                Do Until GetPlayerExp(Index) < GetPlayerNextLevel(Index)
                    DoEvents
                    If GetPlayerLevel(Index) < MAX_LEVEL Then
                        If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
                            d = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
                            Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                            I = Int(GetPlayerLUCK(Index) / 10)
                            If I < 1 Then I = 1
                            If I > 3 Then I = 3
                                
                            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + I)
                            Call SetPlayerExp(Index, d)
                            c = c + 1
                        End If
                    End If
                Loop
                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " levels!", 6)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a level!", 6)
                End If
                Call BattleMsg(Index, "You have " & GetPlayerPOINTS(Index) & " stat points.", 9, 0)
            End If
            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If
        
        If GetPlayerLevel(Index) = MAX_LEVEL Then
            Call SetPlayerExp(Index, Experience(MAX_LEVEL))
        End If
    End If
    
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendStats(Index)
End Sub

Sub CastSpell(ByVal Index As Long, ByVal SpellSlot As Long)
Dim SpellNum As Long, I As Long, n As Long, Damage As Long
Dim Casted As Boolean

    Casted = False
    
    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then
        Call BattleMsg(Index, "You do not have this spell!", BrightRed, 0)
        Exit Sub
    End If
    
    I = GetSpellReqLevel(Index, SpellNum)

    ' Check if they have enough MP
    If GetPlayerMP(Index) < Spell(SpellNum).MPCost Then
        Call BattleMsg(Index, "Not enough mana!", BrightRed, 0)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If I > GetPlayerLevel(Index) Then
        Call BattleMsg(Index, "You must be level " & I & " to cast this spell.", BrightRed, 0)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < Player(Index).AttackTimer + 1000 Then
        Exit Sub
    End If
    
    ' Is the spell script? If so execute it.
    If Spell(SpellNum).Type = SPELL_TYPE_SCRIPTED Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedSpell " & Index & "," & Spell(SpellNum).Data1
        Exit Sub
    End If
        
Dim X As Long, Y As Long

If Spell(SpellNum).AE = 1 Then
    For Y = GetPlayerY(Index) - Spell(SpellNum).Range To GetPlayerY(Index) + Spell(SpellNum).Range
        For X = GetPlayerX(Index) - Spell(SpellNum).Range To GetPlayerX(Index) + Spell(SpellNum).Range
            n = -1
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) = True Then
                    If GetPlayerMap(Index) = GetPlayerMap(I) Then
                        If GetPlayerX(I) = X And GetPlayerY(I) = Y Then
                            If I = Index Then
                                If Spell(SpellNum).Type = SPELL_TYPE_ADDHP Or Spell(SpellNum).Type = SPELL_TYPE_ADDMP Then
                                    Player(Index).Target = I
                                    Player(Index).TargetType = TARGET_TYPE_PLAYER
                                    n = Player(Index).Target
                                End If
                            Else
                                Player(Index).Target = I
                                Player(Index).TargetType = TARGET_TYPE_PLAYER
                                n = Player(Index).Target
                            End If
                        End If
                    End If
                End If
            Next I
            
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(GetPlayerMap(Index), I).num > 0 Then
                    If Npc(MapNpc(GetPlayerMap(Index), I).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), I).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        If MapNpc(GetPlayerMap(Index), I).X = X And MapNpc(GetPlayerMap(Index), I).Y = Y Then
                            Player(Index).Target = I
                            Player(Index).TargetType = TARGET_TYPE_NPC
                            n = Player(Index).Target
                        End If
                    End If
                End If
            Next I
                
        Casted = False
        If n > 0 Then
            If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
                If IsPlaying(n) Then
                    If n <> Index Then
                        Player(Index).TargetType = TARGET_TYPE_PLAYER
                        If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
                            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                    
                            Select Case Spell(SpellNum).Type
                                Case SPELL_TYPE_SUBHP
                            
                                    Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(n)
                                    If Damage > 0 Then
                                        Call AttackPlayer(Index, n, Damage)
                                    Else
                                        Call BattleMsg(Index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BrightRed, 0)
                                    End If
                        
                                Case SPELL_TYPE_SUBMP
                                    Call SetPlayerMP(n, GetPlayerMP(n) - Spell(SpellNum).Data1)
                                    Call SendMP(n)
                    
                                'Case SPELL_TYPE_SUBSP
                                    'Call SetPlayerSP(n, GetPlayerSP(n) - Spell(SpellNum).Data1)
                                    'Call SendSP(n)
                            End Select

                            Casted = True
                        Else
                            If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDMP Then
                                Select Case Spell(SpellNum).Type
                                
                                    Case SPELL_TYPE_ADDHP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                                        Call SendHP(n)
                                                
                                    Case SPELL_TYPE_ADDMP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                                        Call SendMP(n)
                                
                                    'Case SPELL_TYPE_ADDSP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        'Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                                        'Call SendMP(n)
                                End Select
                                
                                Casted = True
                            Else
                                Call PlayerMsg(Index, "Could not cast spell!", BrightRed)
                            End If
                        End If
                    Else
                        Player(Index).TargetType = TARGET_TYPE_PLAYER
                        If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
                        Else
                            If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDMP Then
                                Select Case Spell(SpellNum).Type
                                
                                    Case SPELL_TYPE_ADDHP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                                        Call SendHP(n)
                                                
                                    Case SPELL_TYPE_ADDMP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                                        Call SendMP(n)
                                
                                    'Case SPELL_TYPE_ADDSP
                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        'Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                                        'Call SendMP(n)
                                End Select

                                Casted = True
                            Else
                                Call BattleMsg(Index, "Could not cast spell!", BrightRed, 0)
                            End If
                        End If
                    End If
                Else
                    Call BattleMsg(Index, "Could not cast spell!", BrightRed, 0)
                End If
            Else
                Player(Index).TargetType = TARGET_TYPE_NPC
                If Npc(MapNpc(GetPlayerMap(Index), n).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), n).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    If Spell(SpellNum).Type >= SPELL_TYPE_SUBHP And Spell(SpellNum).Type <= SPELL_TYPE_SUBMP Then
                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on a " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                        Select Case Spell(SpellNum).Type
                            
                            Case SPELL_TYPE_SUBHP
                                Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(Index), n).num).DEF / 2)
                                If Damage > 0 Then
                                    Call AttackNpc(Index, n, Damage)
                                Else
                                    Call BattleMsg(Index, "The spell was to weak to hurt " & Trim(Npc(MapNpc(GetPlayerMap(Index), n).num).Name) & "!", BrightRed, 0)
                                End If
                            
                            Case SPELL_TYPE_SUBMP
                                MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP - Spell(SpellNum).Data1

                            'Case SPELL_TYPE_SUBSP
                                'MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP - Spell(SpellNum).Data1
                        End Select
                        
                        Casted = True
                    Else
                        Select Case Spell(SpellNum).Type
                            Case SPELL_TYPE_ADDHP
                                'MapNpc(GetPlayerMap(Index), n).HP = MapNpc(GetPlayerMap(Index), n).HP + Spell(SpellNum).Data1
                            
                            Case SPELL_TYPE_ADDMP
                                'MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP + Spell(SpellNum).Data1
                            
                            'Case SPELL_TYPE_ADDSP
                                'MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP + Spell(SpellNum).Data1
                        End Select
                        Casted = False
                    End If
                Else
                    Call BattleMsg(Index, "Could not cast spell!", BrightRed, 0)
                End If
            End If
        End If
        If Casted = True Then
            Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & END_CHAR)
            'Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
        End If
        Next X
    Next Y
    
    Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
    Call SendMP(Index)
Else
    n = Player(Index).Target
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
            If GetPlayerName(n) <> GetPlayerName(Index) Then
                If CInt(Sqr((GetPlayerX(Index) - GetPlayerX(n)) ^ 2 + ((GetPlayerY(Index) - GetPlayerY(n)) ^ 2))) > Spell(SpellNum).Range Then
                    Call BattleMsg(Index, "You are too far away to hit the target.", BrightRed, 0)
                    Exit Sub
                End If
            End If
            Player(Index).TargetType = TARGET_TYPE_PLAYER
            If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
        
                Select Case Spell(SpellNum).Type
                    Case SPELL_TYPE_SUBHP
                
                        Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(n)
                        If Damage > 0 Then
                            Call AttackPlayer(Index, n, Damage)
                        Else
                            Call BattleMsg(Index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BrightRed, 0)
                        End If
            
                    Case SPELL_TYPE_SUBMP
                        Call SetPlayerMP(n, GetPlayerMP(n) - Spell(SpellNum).Data1)
                        Call SendMP(n)
        
                    'Case SPELL_TYPE_SUBSP
                        'Call SetPlayerSP(n, GetPlayerSP(n) - Spell(SpellNum).Data1)
                        'Call SendSP(n)
                End Select
            
                ' Take away the mana points
                Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                Call SendMP(Index)
                Casted = True
            Else
                If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDMP Then
                    Select Case Spell(SpellNum).Type
                    
                        Case SPELL_TYPE_ADDHP
                            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                            Call SendHP(n)
                                    
                        Case SPELL_TYPE_ADDMP
                            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                            Call SendMP(n)
                    
                        'Case SPELL_TYPE_ADDSP
                            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            'Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                            'Call SendMP(n)
                    End Select
                    
                    ' Take away the mana points
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                    Call SendMP(Index)
                    Casted = True
                Else
                    Call BattleMsg(Index, "Could not cast spell!", BrightRed, 0)
                End If
            End If
        Else
            Call PlayerMsg(Index, "Could not cast spell!", BrightRed)
        End If
    Else
        If CInt(Sqr((GetPlayerX(Index) - MapNpc(GetPlayerMap(Index), n).X) ^ 2 + ((GetPlayerY(Index) - MapNpc(GetPlayerMap(Index), n).Y) ^ 2))) > Spell(SpellNum).Range Then
            Call BattleMsg(Index, "You are too far away to hit the target.", BrightRed, 0)
            Exit Sub
        End If
        
        Player(Index).TargetType = TARGET_TYPE_NPC
        
        If Npc(MapNpc(GetPlayerMap(Index), n).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), n).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
            'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim(Spell(SpellNum).Name) & " on a " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_ADDHP
                    MapNpc(GetPlayerMap(Index), n).HP = MapNpc(GetPlayerMap(Index), n).HP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBHP
                    
                    Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(Index), n).num).DEF / 2)
                    If Damage > 0 Then
                        Call AttackNpc(Index, n, Damage)
                    Else
                        Call BattleMsg(Index, "The spell was to weak to hurt " & Trim(Npc(MapNpc(GetPlayerMap(Index), n).num).Name) & "!", BrightRed, 0)
                    End If
                    
                Case SPELL_TYPE_ADDMP
                    MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBMP
                    MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP - Spell(SpellNum).Data1
            
                'Case SPELL_TYPE_ADDSP
                    'MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP + Spell(SpellNum).Data1
                
                'Case SPELL_TYPE_SUBSP
                    'MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP - Spell(SpellNum).Data1
            End Select
        
            ' Take away the mana points
            Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
            Call SendMP(Index)
            Casted = True
        Else
            Call BattleMsg(Index, "Could not cast spell!", BrightRed, 0)
        End If
    End If
End If

    If Casted = True Then
        Player(Index).AttackTimer = GetTickCount
        Player(Index).CastedSpell = YES
        Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & Player(Index).CastedSpell & SEP_CHAR & END_CHAR)
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
    End If
End Sub
Function GetSpellReqLevel(ByVal Index As Long, ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).LevelReq ' - Int(GetClassMAGI(GetPlayerClass(index)) / 4)
End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
Dim I As Long, n As Long

    CanPlayerCriticalHit = False
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            I = Int(GetPlayerSTR(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
    
            n = Int(Rnd * 100 - GetPlayerLUCK(Index) / 10) + 1
            If n <= I Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
Dim I As Long, n As Long, ShieldSlot As Long

    CanPlayerBlockHit = False
    
    ShieldSlot = GetPlayerShieldSlot(Index)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            I = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
        
            n = Int(Rnd * 100) + 1
            If n <= I Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Sub CheckEquippedItems(ByVal Index As Long)
Dim Slot As Long, ItemNum As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    Slot = GetPlayerWeaponSlot(Index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
                Call SetPlayerWeaponSlot(Index, 0)
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
End Sub

Public Sub SetPMember(ByVal Leader As Byte, ByVal MemberIndex As Byte)
Dim I As Integer
For I = 1 To MAX_PARTY_MEMBERS
    If Player(Leader).Party.Member(I) = 0 Then
        Player(Leader).Party.Member(I) = MemberIndex
        Exit For
    End If
Next I
Player(MemberIndex).Party.Leader = Leader

For I = 1 To MAX_PARTY_MEMBERS
    If Player(Leader).Party.Member(I) > 0 Then UpdateParty Player(Leader).Party.Member(I)
Next I
End Sub

Public Sub RemovePMember(ByVal Index As Byte)
Dim I, b, q As Integer
I = 0

b = Player(Index).Party.Leader

If Player(Index).Party.Leader = Index Then 'Change the party leader!
    For I = 1 To MAX_PARTY_MEMBERS
        If Player(Index).Party.Member(I) > 0 And Player(Index).Party.Member(I) <> Index Then
        Call ChangePLeader(Player(Index).Party.Member(I))
        Exit For
        End If
    Next I
End If

I = 0

For q = 1 To MAX_PARTY_MEMBERS ' find which member the player is
    If Player(Index).Party.Member(q) = Index Then
        Exit For
    End If
Next q

For I = 1 To MAX_PARTY_MEMBERS ' removes player from other members party
    If Player(Index).Party.Member(I) > 0 Then Player(Player(Index).Party.Member(I)).Party.Member(q) = 0
Next I

Player(Index).Party.Leader = 0 'no leader
Player(Index).InvitedBy = 0

For I = 1 To MAX_PARTY_MEMBERS ' clears player's party
    Player(Index).Party.Member(I) = 0
Next I

Player(Index).InParty = False 'not in party

q = 0

If b > 0 Then
    For I = 1 To MAX_PARTY_MEMBERS 'check to see if we need to clear out the party leader
        If Player(b).Party.Member(I) > 0 Then q = q + 1
    Next I

    If q < 1 Then
    Call PlayerMsg(b, "Your party has been disbanded!", White)
    Player(b).InParty = False

        For I = 1 To MAX_PARTY_MEMBERS ' clears player's party
            Player(b).Party.Member(I) = 0
        Next I
    End If
End If

End Sub

Public Sub ChangePLeader(ByVal Index As Byte)
Dim I As Integer

Player(Index).Party.Leader = Index

For I = 1 To MAX_PARTY_MEMBERS
    If Player(Index).Party.Member(I) > 0 Then Player(Player(Index).Party.Member(I)).Party.Leader = Index
Next I

For I = 1 To MAX_PARTY_MEMBERS
    If Player(Index).Party.Member(I) > 0 Then Call PlayerMsg(Player(Index).Party.Member(I), "Leadership has been passed to " & GetPlayerName(Index) & "!", Pink)
Next I

End Sub

Public Sub UpdateParty(ByVal Index As Byte)
Player(Index).Party = Player(Player(Index).Party.Leader).Party
End Sub

Public Sub SetPShare(ByVal Index As Byte, ByVal share As Boolean)
Player(Index).Party.ShareExp = share
End Sub

Function GetPLeader(ByVal Index As Byte) As Byte
    GetPLeader = Player(Index).Party.Leader
End Function

Function GetPMember(ByVal Index As Byte, ByVal Member As Byte) As Byte
    GetPMember = Player(Index).Party.Member(Member)
End Function

Function GetPShare(ByVal Index As Byte) As Boolean
    GetPShare = Player(Index).Party.ShareExp
End Function

Public Sub ShowPLR(ByVal Index As Long)
Dim ls As ListItem
On Error Resume Next

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(Index) = True Then
        frmServer.lvUsers.ListItems.Remove Index
    End If
    Set ls = frmServer.lvUsers.ListItems.add(Index, , Index)
    
    If IsPlaying(Index) = False Then
        ls.SubItems(1) = ""
        ls.SubItems(2) = ""
        ls.SubItems(3) = ""
        ls.SubItems(4) = ""
        ls.SubItems(5) = ""
    Else
        ls.SubItems(1) = GetPlayerLogin(Index)
        ls.SubItems(2) = GetPlayerName(Index)
        ls.SubItems(3) = GetPlayerLevel(Index)
        ls.SubItems(4) = GetPlayerSprite(Index)
        ls.SubItems(5) = GetPlayerAccess(Index)
    End If
End Sub

Public Sub RemovePLR()
    frmServer.lvUsers.ListItems.Clear
End Sub
Function CanAttackPlayerWithArrow(ByVal Attacker As Long, ByVal Victim As Long) As Boolean

CanAttackPlayerWithArrow = False

' Check If map Is attackable
If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
' Make sure they are high enough level
If GetPlayerLevel(Attacker) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
    Call PlayerMsg(Attacker, "You are below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack another player yet!", BrightRed)
Else
If GetPlayerLevel(Victim) < Val(GetVar("Data.ini", "Config", "pklevel")) Then
    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " Is below level " & Val(GetVar("Data.ini", "Config", "pklevel")) & ", you cannot attack this player yet!", BrightRed)
Else
If Trim(GetPlayerGuild(Attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
If Trim(GetPlayerGuild(Attacker)) <> Trim(GetPlayerGuild(Victim)) Then
    CanAttackPlayerWithArrow = True
Else
    Call PlayerMsg(Attacker, "You cant attack a guild member!", BrightRed)
End If
Else
    CanAttackPlayerWithArrow = True
End If
End If
End If
Else
    Call PlayerMsg(Attacker, "This Is a safe zone!", BrightRed)
End If

End Function

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
Dim AttackSpeed As Long
    
    AttackSpeed = 1000
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        If Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).Type <> ITEM_TYPE_LAMP Then
            AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
        End If
    End If
    

CanAttackNpcWithArrow = False

' Check For subscript out of range
If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
Exit Function
End If

' Check For subscript out of range
If MapNpc(GetPlayerMap(Attacker), MapNpcNum).num <= 0 Then
Exit Function
End If

MapNum = GetPlayerMap(Attacker)
NpcNum = MapNpc(MapNum, MapNpcNum).num

' Make sure the npc isn't already dead
If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
Exit Function
End If

' Make sure they are On the same map
If IsPlaying(Attacker) Then
    If NpcNum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
        ' Check If at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpcWithArrow = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), Green)
                    End If

            Case DIR_DOWN
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpcWithArrow = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), Green)
                    End If

            Case DIR_LEFT
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpcWithArrow = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), Green)
                    End If

            Case DIR_RIGHT
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpcWithArrow = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & ": " & Trim(Npc(NpcNum).AttackSay), Green)
                End If
        End Select
    End If
End If
End Function
