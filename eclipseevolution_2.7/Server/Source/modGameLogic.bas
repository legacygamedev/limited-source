Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim WeaponSlot As Long, RingSlot As Long, NecklaceSlot As Long
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    ' GetPlayerDamage in script - TODO LATER - Can't get it to work. :(
    ' If Scripting = 1 Then
    ' GetPlayerDamage = MyScript.RunCodeReturn("Scripts\Main.txt", "GetPlayerDamage ", index)
    ' Else
    GetPlayerDamage = Int(GetPlayerSTR(Index) / 2)
' End If

    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If

    If GetPlayerWeaponSlot(Index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(Index)

        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2

        If GetPlayerInvItemDur(Index, WeaponSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) - 1)

            If GetPlayerInvItemDur(Index, WeaponSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " has broken.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, WeaponSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, WeaponSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If GetPlayerRingSlot(Index) > 0 Then
        RingSlot = GetPlayerRingSlot(Index)

        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2

        If GetPlayerInvItemDur(Index, RingSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, RingSlot, GetPlayerInvItemDur(Index, RingSlot) - 1)

            If GetPlayerInvItemDur(Index, RingSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, RingSlot)).Name) & " has broken.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, RingSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, RingSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, RingSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, RingSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, RingSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If GetPlayerNecklaceSlot(Index) > 0 Then
        NecklaceSlot = GetPlayerNecklaceSlot(Index)

        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, NecklaceSlot)).Data2

        If GetPlayerInvItemDur(Index, NecklaceSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, NecklaceSlot, GetPlayerInvItemDur(Index, NecklaceSlot) - 1)

            If GetPlayerInvItemDur(Index, NecklaceSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, NecklaceSlot)).Name) & " has broken.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, NecklaceSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, NecklaceSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, NecklaceSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, NecklaceSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, NecklaceSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If



    If GetPlayerDamage < 0 Then
        GetPlayerDamage = 0
    End If
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long, LegsSlot As Long

    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    ArmorSlot = GetPlayerArmorSlot(Index)
    HelmSlot = GetPlayerHelmetSlot(Index)
    ShieldSlot = GetPlayerShieldSlot(Index)
    LegsSlot = GetPlayerLegsSlot(Index)
    GetPlayerProtection = Int(GetPlayerDEF(Index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data2
        If GetPlayerInvItemDur(Index, ArmorSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, ArmorSlot, GetPlayerInvItemDur(Index, ArmorSlot) - 1)

            If GetPlayerInvItemDur(Index, ArmorSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " has broken.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ArmorSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, ArmorSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, ArmorSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, HelmSlot)).Data2
        If GetPlayerInvItemDur(Index, HelmSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, HelmSlot, GetPlayerInvItemDur(Index, HelmSlot) - 1)

            If GetPlayerInvItemDur(Index, HelmSlot) <= 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " has broken.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, HelmSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, HelmSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, HelmSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data2
        If GetPlayerInvItemDur(Index, ShieldSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, ShieldSlot, GetPlayerInvItemDur(Index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(Index, ShieldSlot) <= 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " has broken.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ShieldSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, ShieldSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, ShieldSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If LegsSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, LegsSlot)).Data2
        If GetPlayerInvItemDur(Index, LegsSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, LegsSlot, GetPlayerInvItemDur(Index, LegsSlot) - 1)

            If GetPlayerInvItemDur(Index, LegsSlot) <= 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, LegsSlot)).Name) & " has broken.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, LegsSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, LegsSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, LegsSlot)).Name) & " " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, LegsSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, LegsSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If
End Function

Function FindOpenPlayerSlot() As Long
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If Not IsConnected(I) Then
            FindOpenPlayerSlot = I
            Exit Function
        End If
    Next I
End Function

Public Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check to see if they already have the item.
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, I) = ItemNum Then
                FindOpenInvSlot = I
                Exit Function
            End If
        Next I
    End If

    ' Try to find an open inventory slot.
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, I) = 0 Then
            FindOpenInvSlot = I
            Exit Function
        End If
    Next I
End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check to see if they already have the item.
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        For I = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, I) = ItemNum Then
                FindOpenBankSlot = I
                Exit Function
            End If
        Next I
    End If

    ' Try to find an open bank slot.
    For I = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, I) = 0 Then
            FindOpenBankSlot = I
            Exit Function
        End If
    Next I
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
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

    For I = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, I) = 0 Then
            FindOpenSpellSlot = I
            Exit Function
        End If
    Next I
End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim I As Long

    For I = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, I) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next I
End Function

Function TotalOnlinePlayers() As Long
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next I
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim I As Long

    Name = LCase$(Name)

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If Len(GetPlayerName(I)) >= Len(Name) Then
                If LCase$(GetPlayerName(I)) = Name Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check to see if the player has the item.
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
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
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
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

                    Case ITEM_TYPE_LEGS
                        If GetPlayerLegsSlot(Index) > 0 Then
                            If I = GetPlayerLegsSlot(Index) Then
                                Call SetPlayerLegsSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(Index) > 0 Then
                            If I = GetPlayerRingSlot(Index) Then
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

                    Case ITEM_TYPE_NECKLACE
                        If GetPlayerNecklaceSlot(Index) > 0 Then
                            If I = GetPlayerNecklaceSlot(Index) Then
                                Call SetPlayerNecklaceSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select


                n = Item(GetPlayerInvItemNum(Index, I)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_LEGS) And (n <> ITEM_TYPE_RING) And (n <> ITEM_TYPE_NECKLACE) Then
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

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    I = FindOpenInvSlot(Index, ItemNum)

    ' Check to see if inventory is full
    If I > 0 Then
        Call SetPlayerInvItemNum(Index, I, ItemNum)
        Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) + ItemVal)

        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_NECKLACE) Then
            Call SetPlayerInvItemDur(Index, I, Item(ItemNum).Data1)
        End If

        Call SendInventoryUpdate(Index, I)
    Else
        Call PlayerMsg(Index, "Your inventory has reached its maximum capacity!", BRIGHTRED)
    End If
End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim I As Long, n As Long
    Dim TakeBankItem As Boolean

    TakeBankItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    For I = 1 To MAX_BANK
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(Index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                ' Is what we are trying to take away more then what they have? If so just set it to zero
                If ItemVal >= GetPlayerBankItemValue(Index, I) Then
                    TakeBankItem = True
                Else
                    Call SetPlayerBankItemValue(Index, I, GetPlayerBankItemValue(Index, I) - ItemVal)
                    Call SendBankUpdate(Index, I)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerBankItemNum(Index, I)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If I = GetPlayerWeaponSlot(Index) Then
                                Call SetPlayerWeaponSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerWeaponSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(Index) > 0 Then
                            If I = GetPlayerArmorSlot(Index) Then
                                Call SetPlayerArmorSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerArmorSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(Index) > 0 Then
                            If I = GetPlayerHelmetSlot(Index) Then
                                Call SetPlayerHelmetSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerHelmetSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(Index) > 0 Then
                            If I = GetPlayerShieldSlot(Index) Then
                                Call SetPlayerShieldSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerShieldSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_LEGS
                        If GetPlayerLegsSlot(Index) > 0 Then
                            If I = GetPlayerLegsSlot(Index) Then
                                Call SetPlayerLegsSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerLegsSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(Index) > 0 Then
                            If I = GetPlayerRingSlot(Index) Then
                                Call SetPlayerRingSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerRingSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_NECKLACE
                        If GetPlayerNecklaceSlot(Index) > 0 Then
                            If I = GetPlayerNecklaceSlot(Index) Then
                                Call SetPlayerNecklaceSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerNecklaceSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                End Select


                n = Item(GetPlayerBankItemNum(Index, I)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_LEGS) And (n <> ITEM_TYPE_RING) And (n <> ITEM_TYPE_NECKLACE) Then
                    TakeBankItem = True
                End If
            End If

            If TakeBankItem = True Then
                Call SetPlayerBankItemNum(Index, I, 0)
                Call SetPlayerBankItemValue(Index, I, 0)
                Call SetPlayerBankItemDur(Index, I, 0)

                ' Send the Bank update
                Call SendBankUpdate(Index, I)
                Exit Sub
            End If
        End If
    Next I
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal BankSlot As Long)
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    I = BankSlot

    ' Check to see if Bankentory is full
    If I > 0 Then
        Call SetPlayerBankItemNum(Index, I, ItemNum)
        Call SetPlayerBankItemValue(Index, I, GetPlayerBankItemValue(Index, I) + ItemVal)

        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_NECKLACE) Then
            Call SetPlayerBankItemDur(Index, I, Item(ItemNum).Data1)
        End If
    Else
        Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "Bank full!" & END_CHAR)
    End If
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot.
    I = FindOpenMapItemSlot(MapNum)

    Call SpawnItemSlot(I, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, X, Y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If MapItemSlot < 1 Or MapItemSlot > MAX_MAP_ITEMS Then
        Exit Sub
    End If

    I = MapItemSlot

    If I > 0 Then
        MapItem(MapNum, I).num = ItemNum
        MapItem(MapNum, I).Value = ItemVal

        If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_NECKLACE) Then
            MapItem(MapNum, I).Dur = ItemDur
        Else
            MapItem(MapNum, I).Dur = 0
        End If

        MapItem(MapNum, I).X = X
        MapItem(MapNum, I).Y = Y

        Call SendDataToMap(MapNum, "SPAWNITEM" & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR)
    End If
End Sub

Sub SpawnAllMapsItems()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call SpawnMapItems(I)
    Next I
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim X As Integer
    Dim Y As Integer

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn all the mapped items on their specified tile.
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_ITEM Then
                If (Item(Map(MapNum).Tile(X, Y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(MapNum).Tile(X, Y).Data1).Stackable = 1) And Map(MapNum).Tile(X, Y).Data2 <= 0 Then
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
        If (MapItem(MapNum, I).num > 0) Then
            If (MapItem(MapNum, I).num <= MAX_ITEMS) Then
        
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, I).X = GetPlayerX(Index)) Then
                    If (MapItem(MapNum, I).Y = GetPlayerY(Index)) Then
                    
                        ' Find open slot
                        n = FindOpenInvSlot(Index, MapItem(MapNum, I).num)
        
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventory
                            Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, I).num)
                            If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, n)).Stackable = 1 Then
                                Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, I).Value)
                                Msg = "You pickup " & MapItem(MapNum, I).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
                            Else
                                Call SetPlayerInvItemValue(Index, n, 0)
                                Msg = "You pickup " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
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
                            Call PlayerMsg(Index, Msg, YELLOW)
                            Exit Sub
                        Else
                            Call PlayerMsg(Index, "Your inventory has reached its maximum capacity!", BRIGHTRED)
                            Exit Sub
                        End If
                    End If
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

    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
            I = FindOpenMapItemSlot(GetPlayerMap(Index))
    
            If I <> 0 Then
                MapItem(GetPlayerMap(Index), I).Dur = 0
    
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                    Case ITEM_TYPE_ARMOR
                        If InvNum = GetPlayerArmorSlot(Index) Then
                            Call SetPlayerArmorSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
    
                    Case ITEM_TYPE_WEAPON
                        If InvNum = GetPlayerWeaponSlot(Index) Then
                            Call SetPlayerWeaponSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
    
                    Case ITEM_TYPE_HELMET
                        If InvNum = GetPlayerHelmetSlot(Index) Then
                            Call SetPlayerHelmetSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
    
                    Case ITEM_TYPE_SHIELD
                        If InvNum = GetPlayerShieldSlot(Index) Then
                            Call SetPlayerShieldSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                    Case ITEM_TYPE_LEGS
                        If InvNum = GetPlayerLegsSlot(Index) Then
                            Call SetPlayerLegsSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                    Case ITEM_TYPE_RING
                        If InvNum = GetPlayerRingSlot(Index) Then
                            Call SetPlayerRingSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                    Case ITEM_TYPE_NECKLACE
                        If InvNum = GetPlayerNecklaceSlot(Index) Then
                            Call SetPlayerNecklaceSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                End Select
    
                MapItem(GetPlayerMap(Index), I).num = GetPlayerInvItemNum(Index, InvNum)
                MapItem(GetPlayerMap(Index), I).X = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), I).Y = GetPlayerY(Index)
    
                If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                        MapItem(GetPlayerMap(Index), I).Value = GetPlayerInvItemValue(Index, InvNum)
                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                        Call SetPlayerInvItemDur(Index, InvNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), I).Value = Amount
                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                    End If
                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), I).Value = 0
    
                    ' Normally messages for item drops would go here but it's scripted now
    
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                End If
    
                ' Send inventory update
                Call SendInventoryUpdate(Index, InvNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(I, MapItem(GetPlayerMap(Index), I).num, Amount, MapItem(GetPlayerMap(Index), I).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "onitemdrop " & Index & "," & GetPlayerMap(Index) & "," & MapItem(GetPlayerMap(Index), I).num & "," & Amount & "," & MapItem(GetPlayerMap(Index), I).Dur & "," & I & "," & InvNum
                End If
    
            Else
                Call PlayerMsg(Index, "To many items already on the ground.", BRIGHTRED)
            End If
        End If
        
    End If
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim NPCnum As Long
    Dim I As Long
    Dim X As Long
    Dim Y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Or MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    NPCnum = Map(MapNum).NPC(MapNpcNum)

    If NPCnum > 0 Then
        If GameTime = TIME_NIGHT Then
            If NPC(NPCnum).SpawnTime = 1 Then
                MapNPC(MapNum, MapNpcNum).num = 0
                MapNPC(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNPC(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
                Exit Sub
            End If
        Else
            If NPC(NPCnum).SpawnTime = 2 Then
                MapNPC(MapNum, MapNpcNum).num = 0
                MapNPC(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNPC(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
                Exit Sub
            End If
        End If

        MapNPC(MapNum, MapNpcNum).num = NPCnum
        MapNPC(MapNum, MapNpcNum).Target = 0

        MapNPC(MapNum, MapNpcNum).HP = GetNpcMaxHP(NPCnum)
        MapNPC(MapNum, MapNpcNum).MP = GetNpcMaxMP(NPCnum)
        MapNPC(MapNum, MapNpcNum).SP = GetNpcMaxSP(NPCnum)

        MapNPC(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

        ' This means the admin wants to do a random spawn. [Mellowz]
        If Map(MapNum).SpawnX(MapNpcNum) = 0 Or Map(MapNum).SpawnY(MapNpcNum) = 0 Then
            For I = 1 To 100
                X = Int(Rnd * MAX_MAPX)
                Y = Int(Rnd * MAX_MAPY)
    
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                    MapNPC(MapNum, MapNpcNum).X = X
                    MapNPC(MapNum, MapNpcNum).Y = Y
                    Spawned = True
                    Exit For
                End If
            Next I

            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then
                For Y = 0 To MAX_MAPY
                    For X = 0 To MAX_MAPX
                        If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                            MapNPC(MapNum, MapNpcNum).X = X
                            MapNPC(MapNum, MapNpcNum).Y = Y
                            Spawned = True
                        End If
                    Next X
                Next Y
            End If
        Else
            ' We subtract one because Rand is ListIndex 0. [Mellowz]
            MapNPC(MapNum, MapNpcNum).X = Map(MapNum).SpawnX(MapNpcNum) - 1
            MapNPC(MapNum, MapNpcNum).Y = Map(MapNum).SpawnY(MapNpcNum) - 1
            Spawned = True
        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR
            Call SendDataToMap(MapNum, packet)
        End If
    End If

    ' Enable this to display HP when monsters spawn.
    ' Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNPC(MapNum, MapNpcNum).num) & END_CHAR)
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim I As Long

    For I = 1 To MAX_MAP_NPCS
        If Map(MapNum).NPC(I) > 0 Then
            Call SpawnNpc(I, MapNum)
        End If
    Next I
End Sub

Sub SpawnAllMapNpcs()
    Dim I As Long

    For I = 1 To MAX_MAPS
        If PlayersOnMap(I) = YES Then
            Call SpawnMapNpcs(I)
        End If
    Next I
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
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
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BRIGHTRED)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < PKMINLVL Then
                                    Call PlayerMsg(Attacker, "You are below level " & PKMINLVL & ", you cannot attack another player yet!", BRIGHTRED)
                                Else
                                    If GetPlayerLevel(Victim) < PKMINLVL Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & PKMINLVL & ", you cannot attack this player yet!", BRIGHTRED)
                                    Else
                                        If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                            If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                                CanAttackPlayer = True
                                            Else
                                                Call PlayerMsg(Attacker, "You cant attack a guild member!", BRIGHTRED)
                                            End If
                                        Else
                                            CanAttackPlayer = True
                                        End If
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "This is a safe zone!", BRIGHTRED)
                            End If
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If

            Case DIR_DOWN
                If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure that they dont have access
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < PKMINLVL Then
                                Call PlayerMsg(Attacker, "You are below level " & PKMINLVL & ", you cannot attack another player yet!", BRIGHTRED)
                            Else
                                If GetPlayerLevel(Victim) < PKMINLVL Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & PKMINLVL & ", you cannot attack this player yet!", BRIGHTRED)
                                Else
                                    If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                        If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                            Call PlayerMsg(Attacker, "You cant attack a guild member!", BRIGHTRED)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "This is a safe zone!", BRIGHTRED)
                        End If
                    End If
                End If
                If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                    CanAttackPlayer = True
                End If

            Case DIR_LEFT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < PKMINLVL Then
                                Call PlayerMsg(Attacker, "You are below level " & PKMINLVL & ", you cannot attack another player yet!", BRIGHTRED)
                            Else
                                If GetPlayerLevel(Victim) < PKMINLVL Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & PKMINLVL & ", you cannot attack this player yet!", BRIGHTRED)
                                Else
                                    If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                        If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                                Call PlayerMsg(Attacker, "You cant attack a guild member!", BRIGHTRED)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "This is a safe zone!", BRIGHTRED)
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
                            If GetPlayerLevel(Attacker) < PKMINLVL Then
                                Call PlayerMsg(Attacker, "You are below level " & PKMINLVL & ", you cannot attack another player yet!", BRIGHTRED)
                            Else
                                If GetPlayerLevel(Victim) < PKMINLVL Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & PKMINLVL & ", you cannot attack this player yet!", BRIGHTRED)
                                Else
                                    If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                        If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                            Call PlayerMsg(Attacker, "You cant attack a guild member!", BRIGHTRED)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "This is a safe zone!", BRIGHTRED)
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
        End Select
    End If
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim MapNum As Long
    Dim NPCnum As Long
    Dim AttackSpeed As Long

    ' Check for sub-script out of range.
    If Not IsPlaying(Attacker) Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If MapNPC(GetPlayerMap(Attacker), MapNpcNum).num = 0 Then
        Exit Function
    End If

    ' Get the players weapon attack speed.
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    ' Get the players map number.
    MapNum = GetPlayerMap(Attacker)

    ' Get the NPCs map index.
    NPCnum = MapNPC(MapNum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNPC(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Checks to see if the player can attack.
    If GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (MapNPC(MapNum, MapNpcNum).Y + 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                    If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpc = True
                    Else
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(NPCnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(NPCnum).Name) & " : " & Trim$(NPC(NPCnum).AttackSay), GREEN)
                        End If
                    End If
                End If

            Case DIR_DOWN
                If (MapNPC(MapNum, MapNpcNum).Y - 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                    If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpc = True
                    Else
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(NPCnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(NPCnum).Name) & " : " & Trim$(NPC(NPCnum).AttackSay), GREEN)
                        End If
                    End If
                End If

            Case DIR_LEFT
                If (MapNPC(MapNum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum, MapNpcNum).X + 1 = GetPlayerX(Attacker)) Then
                    If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpc = True
                    Else
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(NPCnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(NPCnum).Name) & " : " & Trim$(NPC(NPCnum).AttackSay), GREEN)
                        End If
                    End If
                End If

            Case DIR_RIGHT
                If (MapNPC(MapNum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum, MapNpcNum).X - 1 = GetPlayerX(Attacker)) Then
                    If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpc = True
                    Else
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(NPCnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(NPCnum).Name) & " : " & Trim$(NPC(NPCnum).AttackSay), GREEN)
                        End If
                    End If
                End If
        End Select
    End If
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long
    Dim NPCnum As Long

    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Make sure the NPC map number isn't out-of-range.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Make sure that it's a valid NPC.
    If MapNPC(GetPlayerMap(Index), MapNpcNum).num < 1 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Index)
    NPCnum = MapNPC(MapNum, MapNpcNum).num

    ' Make sure that the NPC isn't already dead.
    If MapNPC(MapNum, MapNpcNum).HP < 1 Then
        Exit Function
    End If

    ' Make sure that NPCs don't attack more then once a second.
    If GetTickCount < MapNPC(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we don't attack a player if they are switching maps.
    If Player(Index).GettingMap = YES Then
        Exit Function
    End If

    MapNPC(MapNum, MapNpcNum).AttackTimer = GetTickCount

    If IsPlaying(Index) Then
        If NPCnum > 0 Then
            If (GetPlayerY(Index) + 1 = MapNPC(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNPC(MapNum, MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNPC(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNPC(MapNum, MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNPC(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) + 1 = MapNPC(MapNum, MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNPC(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) - 1 = MapNPC(MapNum, MapNpcNum).X) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Exp As Long
    Dim n As Long

    ' Make sure the attack is a valid index.
    If Not IsPlaying(Attacker) Then
        Exit Sub
    End If

    ' Make sure the victim is a valid index.
    If Not IsPlaying(Victim) Then
        Exit Sub
    End If

    ' Remove one SP point every time the player attacks.
    If SP_ATTACK = 1 Then
        If GetPlayerSP(Attacker) > 0 Then
            Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - 1)
            Call SendSP(Attacker)
        Else
            Call PlayerMsg(Attacker, "You feel exhausted from fighting.", BLUE)
            Exit Sub
        End If
    End If
 
    ' If damage is below one, exit this sub routine.
    If Damage < 1 Then
        Exit Sub
    End If

    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    End If

    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & END_CHAR)

    If Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA Then
        If Damage >= GetPlayerHP(Victim) Then
            Call SetPlayerHP(Victim, 0)

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnPVPDeath " & Attacker & "," & Victim
            Else
                Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BRIGHTRED)
            End If

            If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
                If SCRIPTING = 1 Then
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

                    If GetPlayerLegsSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerLegsSlot(Victim), 0)
                    End If

                    If GetPlayerRingSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerRingSlot(Victim), 0)
                    End If

                    If GetPlayerNecklaceSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerNecklaceSlot(Victim), 0)
                    End If
                End If

                ' Calculate exp to give attacker
                Exp = Int(GetPlayerExp(Victim) / 10)

                ' Make sure we dont get less then 0
                If Exp < 0 Then
                    Exp = 0
                End If

                If GetPlayerLevel(Victim) = MAX_LEVEL Then
                    Call BattleMsg(Victim, "You cant lose any experience!", BRIGHTRED, 1)
                    Call BattleMsg(Attacker, GetPlayerName(Victim) & " is the max level!", BRIGHTBLUE, 0)
                Else
                    If Exp = 0 Then
                        Call SetPlayerExp(Victim, 0)
                        Call BattleMsg(Victim, "You didn't lose any experience.", BRIGHTRED, 1)
                        Call BattleMsg(Attacker, "You didn't received any experience.", BRIGHTBLUE, 0)
                    Else
                        Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                        Call BattleMsg(Victim, "You lost " & Exp & " experience.", BRIGHTRED, 1)
                        Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                        Call BattleMsg(Attacker, "You got " & Exp & " experience for killing " & GetPlayerName(Victim) & ".", BRIGHTBLUE, 0)
                    End If
                    
                    Call SendEXP(Victim)
                    Call SendEXP(Attacker)
                End If
            End If

            ' Warp player away
            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
            Else
                If Map(GetPlayerMap(Victim)).BootMap > 0 Then
                    Call PlayerWarp(Victim, Map(GetPlayerMap(Victim)).BootMap, Map(GetPlayerMap(Victim)).BootX, Map(GetPlayerMap(Victim)).BootY)
                Else
                    Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
                End If
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
                    Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a player killer!", BRIGHTRED)
                End If
            Else
                Call SetPlayerPK(Victim, NO)
                Call SendPlayerData(Victim)
                Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a player killer!", BRIGHTRED)
            End If
        Else
            Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
            Call SendHP(Victim)
        End If
    ElseIf Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then
        If Damage >= GetPlayerHP(Victim) Then
            Call SetPlayerHP(Victim, 0)

            ' Check if target is player who died and if so set target to 0
            If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
                Player(Attacker).Target = 0
                Player(Attacker).TargetType = 0
            End If

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnArenaDeath " & Attacker & "," & Victim
            End If
        Else
            Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
            Call SendHP(Victim)
        End If
    End If

    Player(Attacker).AttackTimer = GetTickCount

    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & Player(Victim).Char(Player(Victim).CharNum).Sex & END_CHAR)
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim MapNum As Long

    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If
    
    If Not IsPlaying(Victim) Then
        Exit Sub
    End If

    If Damage < 1 Then
        Exit Sub
    End If

    If MapNPC(GetPlayerMap(Victim), MapNpcNum).num <= 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & END_CHAR)

    MapNum = GetPlayerMap(Victim)

    Name = Trim$(NPC(MapNPC(MapNum, MapNpcNum).num).Name)

    If Damage >= GetPlayerHP(Victim) Then
        Call GlobalMsg(GetPlayerName(Victim) & " was killed by " & Name, BRIGHTRED)

        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            If SCRIPTING = 1 Then
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

                If GetPlayerShieldSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
                End If

                If GetPlayerLegsSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerLegsSlot(Victim), 0)
                End If

                If GetPlayerRingSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerRingSlot(Victim), 0)
                End If
            End If

            ' Calculate exp to give attacker
            Exp = Int(GetPlayerExp(Victim) / 3)

            ' Make sure we dont get less then 0
            If Exp < 0 Then
                Exp = 0
            End If

            If Exp = 0 Then
                Call SetPlayerExp(Victim, 0)
                Call BattleMsg(Victim, "You didn't lose any experience.", BRIGHTRED, 0)
            Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call BattleMsg(Victim, "You lost " & Exp & " experience.", BRIGHTRED, 0)
            End If

            Call SendEXP(Victim)
        End If

        ' Warp player away
        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
        Else
            If Map(GetPlayerMap(Victim)).BootMap > 0 Then
                Call PlayerWarp(Victim, Map(GetPlayerMap(Victim)).BootMap, Map(GetPlayerMap(Victim)).BootX, Map(GetPlayerMap(Victim)).BootY)
            Else
                Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
            End If
        End If

        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)

        ' Set NPC target to 0
        MapNPC(MapNum, MapNpcNum).Target = 0

        ' If the player the attacker killed was a pk then take it away
        If GetPlayerPK(Victim) = YES Then
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
        End If
    Else
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
    End If

    Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & END_CHAR)
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & Player(Victim).Char(Player(Victim).CharNum).Sex & END_CHAR)
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim n As Long, I As Long, q As Integer, X As Long
    Dim MapNum As Long, NPCnum As Long

    ' Removes one SP when you attack.
    If SP_ATTACK = 1 Then
        If GetPlayerSP(Attacker) > 0 Then
            Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - 1)
            Call SendSP(Attacker)
        Else
            Call PlayerMsg(Attacker, "You feel exhausted from fighting.", BLUE)
            Exit Sub
        End If
    End If

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
    NPCnum = MapNPC(MapNum, MapNpcNum).num
    Name = Trim$(NPC(NPCnum).Name)

    If Damage >= MapNPC(MapNum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        Player(Attacker).TargetNPC = 0

        ' Call BattleMsg(Attacker, "You killed a " & Name, BrightRed, 0)
        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnNPCDeath " & Attacker & "," & MapNum & "," & NPCnum & "," & MapNpcNum
        End If

        Dim Add As String

        Add = 0
        If GetPlayerWeaponSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AddEXP
        End If
        If GetPlayerArmorSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerArmorSlot(Attacker))).AddEXP
        End If
        If GetPlayerShieldSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerShieldSlot(Attacker))).AddEXP
        End If
        If GetPlayerLegsSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerLegsSlot(Attacker))).AddEXP
        End If
        If GetPlayerRingSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerRingSlot(Attacker))).AddEXP
        End If
        If GetPlayerNecklaceSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerNecklaceSlot(Attacker))).AddEXP
        End If
        If GetPlayerHelmetSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerHelmetSlot(Attacker))).AddEXP
        End If

        If Add > 0 Then
            If Add < 100 Then
                If Add < 10 Then
                    Add = 0 & ".0" & Right$(Add, 2)
                Else
                    Add = 0 & "." & Right$(Add, 2)
                End If
            Else
                Add = Mid$(Add, 1, 1) & "." & Right$(Add, 2)
            End If
        End If

        ' Calculate exp to give attacker
        If Add > 0 Then
            Exp = NPC(NPCnum).Exp + (NPC(NPCnum).Exp * Val(Add))
        Else
            Exp = NPC(NPCnum).Exp
        End If

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If

        ' Check if in party, if so divide the exp up by 2
        If Player(Attacker).InParty = False Or Player(Attacker).Party.ShareExp = False Then
            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "You can't gain anymore experience!", BRIGHTBLUE, 0)
            Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "You gained " & Exp & " experience.", BRIGHTBLUE, 0)
            End If
        Else
            q = 0
            ' The following code will tell us how many party members we have active
            For X = 1 To MAX_PARTY_MEMBERS
                If Player(Attacker).Party.Member(X) > 0 Then
                    q = q + 1
                End If
            Next X

            ' MsgBox "in party" & q
            If q = 2 Then 'Remember, if they aren't in a party they'll only get one person, so this has to be at least two
                Exp = Exp * 0.75 ' 3/4 experience
                ' MsgBox Exp
                For X = 1 To MAX_PARTY_MEMBERS
                    If Player(Attacker).Party.Member(X) > 0 Then
                        If Player(Player(Attacker).Party.Member(X)).Party.ShareExp = True Then
                            If GetPlayerLevel(Player(Attacker).Party.Member(X)) = MAX_LEVEL Then
                                Call SetPlayerExp(Player(Attacker).Party.Member(X), Experience(MAX_LEVEL))
                                Call BattleMsg(Player(Attacker).Party.Member(X), "You can't gain anymore experience!", BRIGHTBLUE, 0)
                            Else
                                Call SetPlayerExp(Player(Attacker).Party.Member(X), GetPlayerExp(Player(Attacker).Party.Member(X)) + Exp)
                                Call BattleMsg(Player(Attacker).Party.Member(X), "You gained " & Exp & " party experience.", BRIGHTBLUE, 0)
                            End If
                            
                            Call SendEXP(X)
                        End If
                    End If
                Next X
            Else 'if there are 3 or more party members..
                Exp = Exp * 0.5  ' 1/2 experience
                For X = 1 To MAX_PARTY_MEMBERS
                    If Player(Attacker).Party.Member(X) > 0 Then
                        If Player(Player(Attacker).Party.Member(X)).Party.ShareExp = True Then
                            If GetPlayerLevel(Player(Attacker).Party.Member(X)) = MAX_LEVEL Then
                                Call SetPlayerExp(Player(Attacker).Party.Member(X), Experience(MAX_LEVEL))
                                Call BattleMsg(Player(Attacker).Party.Member(X), "You can't gain anymore experience!", BRIGHTBLUE, 0)
                            Else
                                Call SetPlayerExp(Player(Attacker).Party.Member(X), GetPlayerExp(Player(Attacker).Party.Member(X)) + Exp)
                                Call BattleMsg(Player(Attacker).Party.Member(X), "You gained " & Exp & " party experience.", BRIGHTBLUE, 0)
                            End If
                            Call SendEXP(X)
                        End If
                    End If
                Next X
            End If
        End If

        ' Drop the items if they earn it.
        For I = 1 To MAX_NPC_DROPS
            If NPC(NPCnum).ItemNPC(I).ItemNum > 0 Then
                n = Int(Rnd * NPC(NPCnum).ItemNPC(I).Chance) + 1
                If n = 1 Then
                    Call SpawnItem(NPC(NPCnum).ItemNPC(I).ItemNum, NPC(NPCnum).ItemNPC(I).ItemValue, MapNum, MapNPC(MapNum, MapNpcNum).X, MapNPC(MapNum, MapNpcNum).Y)
                End If
            End If
        Next I

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNPC(MapNum, MapNpcNum).num = 0
        MapNPC(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNPC(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)

        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)

        ' Check for level up party member
        If Player(Attacker).InParty = True Then
            For X = 1 To MAX_PARTY_MEMBERS
                Call CheckPlayerLevelUp(Player(Attacker).Party.Member(X))
            Next X
        End If

        ' Check for level up party member
        If Player(Attacker).InParty = True Then
            Call CheckPlayerLevelUp(Player(Attacker).PartyPlayer)
        End If

        ' Check if target is npc that died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_NPC And Player(Attacker).Target = MapNpcNum Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
    Else
        ' NPC not dead, just do the damage
        MapNPC(MapNum, MapNpcNum).HP = MapNPC(MapNum, MapNpcNum).HP - Damage
        Player(Attacker).TargetNPC = MapNpcNum

' Check for a weapon and say damage
' Call BattleMsg(Attacker, "You hit a " & Name & " for " & Damage & " damage.", White, 0)

        If n = 0 Then
        ' Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        Else
        ' Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If

        ' Check if we should send a message
        If MapNPC(MapNum, MapNpcNum).Target = 0 And MapNPC(MapNum, MapNpcNum).Target <> Attacker Then
            If Trim$(NPC(NPCnum).AttackSay) <> vbNullString Then
                Call PlayerMsg(Attacker, "A " & Trim$(NPC(NPCnum).Name) & " : " & Trim$(NPC(NPCnum).AttackSay) & vbNullString, SayColor)
            End If
        End If

        ' Set the NPC target to the player
        MapNPC(MapNum, MapNpcNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNPC(MapNum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For I = 1 To MAX_MAP_NPCS
                If MapNPC(MapNum, I).num = MapNPC(MapNum, MapNpcNum).num Then
                    MapNPC(MapNum, I).Target = Attacker
                End If
            Next I
        End If
    End If

    Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNPC(MapNum, MapNpcNum).num) & END_CHAR)

    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub JoinWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save current number map the player is on.
    OldMap = GetPlayerMap(Index)

    Call SendLeaveMap(Index, OldMap)

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    ' Check to see if anyone is on the map.
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexWornEquipmentFromMap(Index)
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    On Error GoTo WarpErr

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save current number map the player is on.
    OldMap = GetPlayerMap(Index)

    If Not OldMap = MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    ' Check to see if anyone is on the map.
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "warp" & END_CHAR)

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexInventoryFromMap(Index)
    Call SendIndexWornEquipmentFromMap(Index)

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & Index & "," & OldMap & "," & MapNum
    End If
    
    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & Index & " on map " & GetPlayerMap(Index) & ".", "logs\ErrorLog.txt")
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
    Dim packet As String
    Dim MapNum As Long
    Dim X As Long
    Dim Y As Long
    Dim I As Long
    Dim Moved As Byte
    Dim sheet As Long
    Dim a As Long

    ' They tried to hack
    ' If Moved = NO Then
    ' Call HackingAttempt(index, "Position Modification")
    ' Exit Sub
    ' End If

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    If Player(Index).GettingMap = True Then
        Exit Sub
    End If

    ' Check for scrolling to prevent RTE 9
    If GetPlayerX(Index) > MAX_MAPX Or GetPlayerY(Index) > MAX_MAPY Then
        Call PlayerWarp(Index, GetPlayerMap(Index), 0, 0)
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)

    ' Remove SP if the player is running.
    If SP_RUNNING = 1 Then
        If Movement = MOVING_RUNNING Then
            If GetPlayerSP(Index) > 0 Then
                Call SetPlayerSP(Index, GetPlayerSP(Index) - 1)
                Call SendSP(Index)
            Else
                Call PlayerMsg(Index, "You feel exhausted from running.", BLUE)
            End If
        End If
    End If

    Moved = NO


    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) - 1)

                        packet = "playermove" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), packet)
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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) + 1)

                        packet = "playermove" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), packet)
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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) - 1)

                        packet = "playermove" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), packet)
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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) + 1)

                        packet = "playermove" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), packet)
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
        Call HackingAttempt(Index, vbNullString)
        Exit Sub
    End If

    ' healing tiles code
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SendHP(Index)
        Call SendMP(Index)
        Call PlayerMsg(Index, "You feel a sudden rush through your body as you regain strength!", BRIGHTGREEN)
    End If

    ' Check for kill tile, and if so kill them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KILL Then
        Call SetPlayerHP(Index, 0)
        Call PlayerMsg(Index, "You embrace the cold finger of death; and feel your life extinguished", BRIGHTRED)

        ' Warp player away
        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Index
        Else
            If Map(GetPlayerMap(Index)).BootMap > 0 Then
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).BootMap, Map(GetPlayerMap(Index)).BootX, Map(GetPlayerMap(Index)).BootY)
            Else
                Call PlayerWarp(Index, START_MAP, START_X, START_Y)
            End If
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

                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
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

                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
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

                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
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

                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
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

            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", WHITE)
            Else
                Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), WHITE)
            End If
            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & END_CHAR)
        End If
    End If

    ' Check for shop
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SHOP Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 > 0 Then
            Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
        Else
            Call PlayerMsg(Index, "There is no shop here.", BRIGHTRED)
        End If
    End If

    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
            Call PlayerMsg(Index, "You already have this sprite!", BRIGHTRED)
            Exit Sub
        Else
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 = 0 Then
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 0 & END_CHAR)
            Else
                If Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(Index, "This sprite will cost you " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 & " " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", YELLOW)
                Else
                    Call PlayerMsg(Index, "This sprite will cost you a " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", YELLOW)
                End If
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 1 & END_CHAR)
            End If
        End If
    End If

    ' Check if player stepped on house buying tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HOUSE Then
        If Len(Map(GetPlayerMap(Index)).Owner) < 2 Then
            If GetPlayerName(Index) = Map(GetPlayerMap(Index)).Owner Then
                Call PlayerMsg(Index, "You already own this house!", BRIGHTRED)
                Exit Sub
            Else
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 = 0 Then
                    Call SendDataTo(Index, "housebuy" & SEP_CHAR & 0 & END_CHAR)
                Else
                    If Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Type = ITEM_TYPE_CURRENCY Then
                        Call PlayerMsg(Index, "This house will cost you " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 & " " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Name) & "!", YELLOW)
                    Else
                        Call PlayerMsg(Index, "This house will cost you a " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Name) & "!", YELLOW)
                    End If
                    Call SendDataTo(Index, "housebuy" & SEP_CHAR & 1 & END_CHAR)
                End If
            End If
        Else
            Call PlayerMsg(Index, "This house is not for sale!", BRIGHTRED)
            Exit Sub
        End If
    End If

    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 > -1 Then
            If GetPlayerClass(Index) <> Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
                Call PlayerMsg(Index, "You arent the required class!", BRIGHTRED)
                Exit Sub
            End If
        End If

        If GetPlayerClass(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
            Call PlayerMsg(Index, "You are already this class!", BRIGHTRED)
        Else
            If Player(Index).Char(Player(Index).CharNum).Sex = 0 Then
                If GetPlayerSprite(Index) = ClassData(GetPlayerClass(Index)).MaleSprite Then
                    Call SetPlayerSprite(Index, ClassData(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).MaleSprite)
                End If
            Else
                If GetPlayerSprite(Index) = ClassData(GetPlayerClass(Index)).FemaleSprite Then
                    Call SetPlayerSprite(Index, ClassData(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).FemaleSprite)
                End If
            End If

            Call SetPlayerSTR(Index, (Player(Index).Char(Player(Index).CharNum).STR - ClassData(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF - ClassData(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi - ClassData(GetPlayerClass(Index)).Magi))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed - ClassData(GetPlayerClass(Index)).Speed))

            Call SetPlayerClassData(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)

            Call SetPlayerSTR(Index, (Player(Index).Char(Player(Index).CharNum).STR + ClassData(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF + ClassData(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi + ClassData(GetPlayerClass(Index)).Magi))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed + ClassData(GetPlayerClass(Index)).Speed))


            Call PlayerMsg(Index, "Your new class is a " & Trim$(ClassData(GetPlayerClass(Index)).Name) & "!", BRIGHTGREEN)

            Call SendStats(Index)
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Player(Index).Char(Player(Index).CharNum).MAXHP = GetPlayerMaxHP(Index)
            Player(Index).Char(Player(Index).CharNum).MAXMP = GetPlayerMaxMP(Index)
            Player(Index).Char(Player(Index).CharNum).MAXSP = GetPlayerMaxSP(Index)
            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
        End If
    End If

    ' Check if player stepped on notice tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_NOTICE Then
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), BLACK)
        End If
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2), GREY)
        End If
        If Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 = vbNullString Or Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 = vbNullString Then
            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 & END_CHAR)
        End If
    End If

    ' Check if player steppted on minus stat tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_LOWER_STAT Then
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), BLACK)
        End If
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1) <> 0 Then
            Call SetPlayerHP(Index, GetPlayerHP(Index) - Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1))
        End If
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2) <> 0 Then
            Call SetPlayerMP(Index, GetPlayerMP(Index) - Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2))
        End If
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3) <> 0 Then
            Call SetPlayerSP(Index, GetPlayerSP(Index) - Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3))
        End If
    End If

    ' Check if player stepped on sound tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1 & END_CHAR)
    End If

    If SCRIPTING = 1 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        End If
    End If

    ' Check if player stepped on Bank tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_BANK Then
        Call SendDataTo(Index, "openbank" & END_CHAR)
    End If

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
    Dim I As Long
    Dim TileType As Long
    Dim X As Long
    Dim Y As Long

    ' Check for sub-script out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    X = MapNPC(MapNum, MapNpcNum).X
    Y = MapNPC(MapNum, MapNpcNum).Y

    CanNpcMove = True

    Select Case Dir
        Case DIR_UP
            If Y > 0 Then
                ' Get the attribute on the tile.
                TileType = Map(MapNum).Tile(X, Y - 1).Type
                                
                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = MapNum Then
                            If GetPlayerX(I) = MapNPC(MapNum, MapNpcNum).X Then
                                If GetPlayerY(I) = (MapNPC(MapNum, MapNpcNum).Y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I

                ' Check to make sure that there is not another npc in the way.
                For I = 1 To MAX_MAP_NPCS
                    If I <> MapNpcNum Then
                        If MapNPC(MapNum, I).num > 0 Then
                            If MapNPC(MapNum, I).X = MapNPC(MapNum, MapNpcNum).X Then
                                If MapNPC(MapNum, I).Y = (MapNPC(MapNum, MapNpcNum).Y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN
            If Y < MAX_MAPY Then
                ' Get the attribute on the tile.
                TileType = Map(MapNum).Tile(X, Y + 1).Type
                
                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = MapNum Then
                            If GetPlayerX(I) = MapNPC(MapNum, MapNpcNum).X Then
                                If GetPlayerY(I) = (MapNPC(MapNum, MapNpcNum).Y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I

                ' Check to make sure that there is not another npc in the way.
                For I = 1 To MAX_MAP_NPCS
                    If I <> MapNpcNum Then
                        If MapNPC(MapNum, I).num > 0 Then
                            If MapNPC(MapNum, I).X = MapNPC(MapNum, MapNpcNum).X Then
                                If MapNPC(MapNum, I).Y = (MapNPC(MapNum, MapNpcNum).Y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT
            If X > 0 Then
                ' Get the attribute on the tile.
                TileType = Map(MapNum).Tile(X - 1, Y).Type

                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = MapNum Then
                            If GetPlayerX(I) = (MapNPC(MapNum, MapNpcNum).X - 1) Then
                                If GetPlayerY(I) = MapNPC(MapNum, MapNpcNum).Y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I

                ' Check to make sure that there is not another npc in the way.
                For I = 1 To MAX_MAP_NPCS
                    If I <> MapNpcNum Then
                        If MapNPC(MapNum, I).num > 0 Then
                            If MapNPC(MapNum, I).X = (MapNPC(MapNum, MapNpcNum).X - 1) Then
                                If MapNPC(MapNum, I).Y = MapNPC(MapNum, MapNpcNum).Y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT
            If X < MAX_MAPX Then
                ' Get the attribute on the tile.
                TileType = Map(MapNum).Tile(X + 1, Y).Type
                
                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = MapNum Then
                            If GetPlayerX(I) = (MapNPC(MapNum, MapNpcNum).X + 1) Then
                                If GetPlayerY(I) = MapNPC(MapNum, MapNpcNum).Y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I

                ' Check to make sure that there is not another npc in the way.
                For I = 1 To MAX_MAP_NPCS
                    If I <> MapNpcNum Then
                        If MapNPC(MapNum, I).num > 0 Then
                            If MapNPC(MapNum, I).X = (MapNPC(MapNum, MapNpcNum).X + 1) Then
                                If MapNPC(MapNum, I).Y = MapNPC(MapNum, MapNpcNum).Y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Public Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
    ' Check to make sure it's a valid map.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid NPC.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid direction.
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Check to make sure it's a valid movement speed.
    If Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    MapNPC(MapNum, MapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNPC(MapNum, MapNpcNum).Y = MapNPC(MapNum, MapNpcNum).Y - 1

        Case DIR_DOWN
            MapNPC(MapNum, MapNpcNum).Y = MapNPC(MapNum, MapNpcNum).Y + 1

        Case DIR_LEFT
            MapNPC(MapNum, MapNpcNum).X = MapNPC(MapNum, MapNpcNum).X - 1

        Case DIR_RIGHT
            MapNPC(MapNum, MapNpcNum).X = MapNPC(MapNum, MapNpcNum).X + 1
    End Select

    Call SendDataToMap(MapNum, "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR)
End Sub

Public Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
    ' Check to make sure it's a valid map.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid NPC.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid direction.
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNPC(MapNum, MapNpcNum).Dir = Dir

    Call SendDataToMap(MapNum, "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & END_CHAR)
End Sub

Public Sub JoinGame(ByVal Index As Long)
    Dim MOTD As String

    ' Set the flag so we know the person is in the game
    Player(Index).InGame = True

    ' Send an ok to client to start receiving in game data
    Call SendDataTo(Index, "loginok" & SEP_CHAR & Index & END_CHAR)

    ReDim Player(Index).Party.Member(1 To MAX_PARTY_MEMBERS)

    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendEmoticons(Index)
    Call SendElements(Index)
    Call SendArrows(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendInventory(Index)
    Call SendBank(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendPTS(Index)
    Call SendStats(Index)
    Call SendWeatherTo(Index)
    Call SendTimeTo(Index)
    Call SendGameClockTo(Index)
    Call DisabledTimeTo(Index)
    Call SendSprite(Index, Index)
    Call SendPlayerSpells(Index)
    Call SendOnlineList

    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "JoinGame " & Index
    Else
        ' Send a global message that he/she joined.
        If GetPlayerAccess(Index) = 0 Then
            Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", 7)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", 15)
        End If

        Call PlayerMsg(Index, "Welcome to " & GAME_NAME & "!", 15)

        ' Send the player the welcome message.
        MOTD = Trim$(GetVar(App.Path & "\MOTD.ini", "MOTD", "Msg"))
        If LenB(MOTD) <> 0 Then
            Call PlayerMsg(Index, "MOTD: " & MOTD, 11)
        End If

        ' Update all clients with the player.
        Call SendWhosOnline(Index)
    End If

    ' Tell the client the player is in-game.
    Call SendDataTo(Index, "ingame" & END_CHAR)

    ' Update the server console.
    Call ShowPLR(Index)
End Sub

Public Sub LeftGame(ByVal Index As Long)
    Dim n As Long

    If Player(Index).InGame Then
        Player(Index).InGame = False

        ' Stop processing NPCs if no one is on it.
        If GetTotalMapPlayers(GetPlayerMap(Index)) = 0 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If

        ' If player is in party, remove experience decrease.
        Call RemovePMember(Index)

        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "LeftGame " & Index
        Else
            ' Check to see if there is any boot map data.
            If Map(GetPlayerMap(Index)).BootMap > 0 Then
                Call SetPlayerX(Index, Map(GetPlayerMap(Index)).BootX)
                Call SetPlayerY(Index, Map(GetPlayerMap(Index)).BootY)
                Call SetPlayerMap(Index, Map(GetPlayerMap(Index)).BootMap)
            End If

            ' Inform the server that the player logged off.
            If GetPlayerAccess(Index) = 0 Then
                Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", 7)
            Else
                Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", 15)
            End If
        End If

        Call SavePlayer(Index)
        Call SendLeftGame(Index)

        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".", True)

        Call RemovePLR(Index)
    End If

    Call ClearPlayer(Index)
    Call SendOnlineList
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim I As Long

    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                GetTotalMapPlayers = GetTotalMapPlayers + 1
            End If
        End If
    Next I
End Function

Function GetNpcMaxHP(ByVal NPCnum As Long) As Long
    If NPCnum < 1 Or NPCnum > MAX_NPCS Then
        Exit Function
    End If

    GetNpcMaxHP = NPC(NPCnum).MAXHP
End Function

Function GetNpcMaxMP(ByVal NPCnum As Long) As Long
    If NPCnum < 1 Or NPCnum > MAX_NPCS Then
        Exit Function
    End If

    GetNpcMaxMP = NPC(NPCnum).Magi * 2
End Function

Function GetNpcMaxSP(ByVal NPCnum As Long) As Long
    If NPCnum < 1 Or NPCnum > MAX_NPCS Then
        Exit Function
    End If

    GetNpcMaxSP = NPC(NPCnum).Speed * 2
End Function

Function GetPlayerHPRegen(ByVal Index As Long) As Integer
    Dim Total As Integer

    If HP_REGEN = 1 Then
        If Index < 1 Or Index > MAX_PLAYERS Then
            Exit Function
        End If

        If Not IsPlaying(Index) Then
            Exit Function
        End If

        Total = Int(GetPlayerDEF(Index) / 2)
        If Total < 2 Then
            Total = 2
        End If

        GetPlayerHPRegen = Total
    End If
End Function

Function GetPlayerMPRegen(ByVal Index As Long) As Integer
    Dim Total As Integer

    If MP_REGEN = 1 Then
        If Index < 1 Or Index > MAX_PLAYERS Then
            Exit Function
        End If

        If Not IsPlaying(Index) Then
            Exit Function
        End If

        Total = Int(GetPlayerMAGI(Index) / 2)
        If Total < 2 Then
            Total = 2
        End If

        GetPlayerMPRegen = Total
    End If
End Function

Function GetPlayerSPRegen(ByVal Index As Long) As Integer
    Dim Total As Integer

    If SP_REGEN = 1 Then
        If Index < 1 Or Index > MAX_PLAYERS Then
            Exit Function
        End If

        If Not IsPlaying(Index) Then
            Exit Function
        End If

        Total = Int(GetPlayerSPEED(Index) / 2)
        If Total < 2 Then
            Total = 2
        End If

        GetPlayerSPRegen = Total
    End If
End Function

Function GetNpcHPRegen(ByVal NPCnum As Long) As Integer
    Dim Total As Integer

    If NPC_REGEN = 1 Then
        If NPCnum < 1 Or NPCnum > MAX_NPCS Then
            Exit Function
        End If
    
        Total = Int(NPC(NPCnum).DEF / 3)
        If Total < 1 Then
            Total = 1
        End If
    
        GetNpcHPRegen = Total
    End If
End Function

Sub CheckPlayerLevelUp(ByVal Index As Long)
    Dim I As Long
    Dim d As Long
    Dim c As Long
    c = 0

    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        If GetPlayerLevel(Index) < MAX_LEVEL Then
            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerLevelUp " & Index
            Else
                Do Until GetPlayerExp(Index) < GetPlayerNextLevel(Index)
                    DoEvents
                    If GetPlayerLevel(Index) < MAX_LEVEL Then
                        If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
                            d = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
                            Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                            I = Int(GetPlayerSPEED(Index) / 10)
                            If I < 1 Then
                                I = 1
                            End If
                            If I > 3 Then
                                I = 3
                            End If

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
                Call BattleMsg(Index, "You have " & GetPlayerPOINTS(Index) & " stat points", 9, 0)
            End If
            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & END_CHAR)
            Call SendPlayerLevelToAll(Index)
        End If

        If GetPlayerLevel(Index) = MAX_LEVEL Then
            Call SetPlayerExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendPTS(Index)

    Player(Index).Char(Player(Index).CharNum).MAXHP = GetPlayerMaxHP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXMP = GetPlayerMaxMP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXSP = GetPlayerMaxSP(Index)

    Call SendStats(Index)
End Sub

Sub CastSpell(ByVal Index As Long, ByVal SpellSlot As Long)
    Dim SpellNum As Long, I As Long, n As Long, Damage As Long
    Dim Casted As Boolean
    Casted = False

    ' Prevent player from using spells if they have been script locked
    If Player(Index).LockedSpells = True Then
        Exit Sub
    End If

    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    SpellNum = GetPlayerSpell(Index, SpellSlot)

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then
        Call BattleMsg(Index, "You do not have this spell!", BRIGHTRED, 0)
        Exit Sub
    End If

    I = GetSpellReqLevel(SpellNum)

    ' Check if they have enough MP
    If GetPlayerMP(Index) < Spell(SpellNum).MPCost Then
        Call BattleMsg(Index, "Not enough mana!", BRIGHTRED, 0)
        Exit Sub
    End If

    ' Make sure they are the right level
    If I > GetPlayerLevel(Index) Then
        Call BattleMsg(Index, "You need to be " & I & "to cast this spell.", BRIGHTRED, 0)
        Exit Sub
    End If

    ' Check if timer is ok
    If GetTickCount < Player(Index).AttackTimer + 1000 Then
        Exit Sub
    End If

    ' Check if the spell is scripted and do that instead of a stat modification
    If Spell(SpellNum).Type = SPELL_TYPE_SCRIPTED Then

        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedSpell " & Index & "," & Spell(SpellNum).Data1

        Exit Sub
    End If
' End If

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
                                    If Spell(SpellNum).Type = SPELL_TYPE_ADDHP Or Spell(SpellNum).Type = SPELL_TYPE_ADDMP Or Spell(SpellNum).Type = SPELL_TYPE_ADDSP Then
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
                    If MapNPC(GetPlayerMap(Index), I).num > 0 Then
                        If NPC(MapNPC(GetPlayerMap(Index), I).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(GetPlayerMap(Index), I).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            If MapNPC(GetPlayerMap(Index), I).X = X And MapNPC(GetPlayerMap(Index), I).Y = Y Then
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
                            If n = Index Then
                                Select Case Spell(SpellNum).Type

                                    Case SPELL_TYPE_ADDHP
                                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                                        Call SendHP(n)

                                    Case SPELL_TYPE_ADDMP
                                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                                        Call SendMP(n)

                                    Case SPELL_TYPE_ADDSP
                                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                                        Call SendMP(n)
                                End Select

                                Casted = True
                            Else
                                Call PlayerMsg(Index, "Cannot cast spell!", BRIGHTRED)
                            End If
                            If n <> Index Then
                                Player(Index).TargetType = TARGET_TYPE_PLAYER
                                If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)

                                    Select Case Spell(SpellNum).Type
                                        Case SPELL_TYPE_SUBHP

                                            Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(n)
                                            If Damage > 0 Then
                                                Call AttackPlayer(Index, n, Damage)
                                            Else
                                                Call BattleMsg(Index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BRIGHTRED, 0)
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
                                    If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then
                                        Select Case Spell(SpellNum).Type

                                            Case SPELL_TYPE_ADDHP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                                                Call SendHP(n)

                                            Case SPELL_TYPE_ADDMP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                                                Call SendMP(n)

                                            Case SPELL_TYPE_ADDSP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                                                Call SendMP(n)
                                        End Select

                                        Casted = True
                                    Else
                                        Call PlayerMsg(Index, "Cannot cast spell!", BRIGHTRED)
                                    End If
                                End If
                            Else
                                Player(Index).TargetType = TARGET_TYPE_PLAYER
                                If n = Index Then
                                    Select Case Spell(SpellNum).Type

                                        Case SPELL_TYPE_ADDHP
                                            ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                            Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                                            Call SendHP(n)

                                        Case SPELL_TYPE_ADDMP
                                            ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                            Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                                            Call SendMP(n)

                                        Case SPELL_TYPE_ADDSP
                                            ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                            Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                                            Call SendMP(n)
                                    End Select

                                    Casted = True
                                Else
                                    Call PlayerMsg(Index, "Cannot cast spell!", BRIGHTRED)
                                End If
                                If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
                                Else
                                    If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then
                                        Select Case Spell(SpellNum).Type

                                            Case SPELL_TYPE_ADDHP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                                                Call SendHP(n)

                                            Case SPELL_TYPE_ADDMP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                                                Call SendMP(n)

                                            Case SPELL_TYPE_ADDSP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                                                Call SendMP(n)
                                        End Select

                                        Casted = True
                                    Else
                                        Call BattleMsg(Index, "Could not cast spell!", BRIGHTRED, 0)
                                    End If
                                End If
                            End If
                        Else
                            Call BattleMsg(Index, "Could not cast spell!", BRIGHTRED, 0)
                        End If
                    Else
                        Player(Index).TargetType = TARGET_TYPE_NPC
                        If NPC(MapNPC(GetPlayerMap(Index), n).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(GetPlayerMap(Index), n).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            If Spell(SpellNum).Type >= SPELL_TYPE_SUBHP And Spell(SpellNum).Type <= SPELL_TYPE_SUBSP Then
                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                                Select Case Spell(SpellNum).Type

                                    Case SPELL_TYPE_SUBHP
                                        Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - Int(NPC(MapNPC(GetPlayerMap(Index), n).num).DEF / 2)

                                        If Damage > 0 Then
                                            If Spell(SpellNum).Element <> 0 And NPC(MapNPC(GetPlayerMap(Index), n).num).Element <> 0 Then
                                                If Element(Spell(SpellNum).Element).Strong = NPC(MapNPC(GetPlayerMap(Index), n).num).Element Or Element(NPC(MapNPC(GetPlayerMap(Index), n).num).Element).Weak = Spell(SpellNum).Element Then
                                                    Call BattleMsg(Index, "A deadly mix of elements harm the " & Trim$(NPC(MapNPC(GetPlayerMap(Index), n).num).Name) & "!", BLUE, 0)
                                                    Damage = Int(Damage * 1.25)
                                                    If Element(Spell(SpellNum).Element).Strong = NPC(MapNPC(GetPlayerMap(Index), n).num).Element And Element(NPC(MapNPC(GetPlayerMap(Index), n).num).Element).Weak = Spell(SpellNum).Element Then
                                                        Damage = Int(Damage * 1.2)
                                                    End If
                                                End If

                                                If Element(Spell(SpellNum).Element).Weak = NPC(MapNPC(GetPlayerMap(Index), n).num).Element Or Element(NPC(MapNPC(GetPlayerMap(Index), n).num).Element).Strong = Spell(SpellNum).Element Then
                                                    Call BattleMsg(Index, "The " & Trim$(NPC(MapNPC(GetPlayerMap(Index), n).num).Name) & " aborbs much of the elemental damage!", RED, 0)
                                                    Damage = Int(Damage * 0.75)
                                                    If Element(Spell(SpellNum).Element).Weak = NPC(MapNPC(GetPlayerMap(Index), n).num).Element And Element(NPC(MapNPC(GetPlayerMap(Index), n).num).Element).Strong = Spell(SpellNum).Element Then
                                                        Damage = Int(Damage * (2 / 3))
                                                    End If
                                                End If
                                            End If
                                            Call AttackNpc(Index, n, Damage)
                                        Else
                                            Call BattleMsg(Index, "The spell was to weak to hurt " & Trim$(NPC(MapNPC(GetPlayerMap(Index), n).num).Name) & "!", BRIGHTRED, 0)
                                        End If

                                    Case SPELL_TYPE_SUBMP
                                        MapNPC(GetPlayerMap(Index), n).MP = MapNPC(GetPlayerMap(Index), n).MP - Spell(SpellNum).Data1

                                    Case SPELL_TYPE_SUBSP
                                        MapNPC(GetPlayerMap(Index), n).SP = MapNPC(GetPlayerMap(Index), n).SP - Spell(SpellNum).Data1
                                End Select

                                Casted = True
                            Else
                                Select Case Spell(SpellNum).Type
                                    Case SPELL_TYPE_ADDHP
' MapNpc(GetPlayerMap(Index), n).HP = MapNpc(GetPlayerMap(Index), n).HP + Spell(SpellNum).Data1

                                    Case SPELL_TYPE_ADDMP
' MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP + Spell(SpellNum).Data1

                                    Case SPELL_TYPE_ADDSP
                                ' MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP + Spell(SpellNum).Data1
                                End Select
                                Casted = False
                            End If
                        Else
                            Call BattleMsg(Index, "Could not cast spell!", BRIGHTRED, 0)
                        End If
                    End If
                End If
                If Casted = True Then
                    Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & Player(Index).CastedSpell & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & END_CHAR)
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
                        Call BattleMsg(Index, "You are too far away to hit the target.", BRIGHTRED, 0)
                        Exit Sub
                    End If
                End If
                Player(Index).TargetType = TARGET_TYPE_PLAYER
                If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)

                    Select Case Spell(SpellNum).Type
                        Case SPELL_TYPE_SUBHP

                            Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(n)
                            If Damage > 0 Then
                                Call AttackPlayer(Index, n, Damage)
                            Else
                                Call BattleMsg(Index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BRIGHTRED, 0)
                            End If

                        Case SPELL_TYPE_SUBMP
                            Call SetPlayerMP(n, GetPlayerMP(n) - Spell(SpellNum).Data1)
                            Call SendMP(n)

                        Case SPELL_TYPE_SUBSP
                            Call SetPlayerSP(n, GetPlayerSP(n) - Spell(SpellNum).Data1)
                            Call SendSP(n)
                    End Select

                    ' Take away the mana points
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                    Call SendMP(Index)
                    Casted = True
                Else
                    If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then
                        Select Case Spell(SpellNum).Type

                            Case SPELL_TYPE_ADDHP
                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                                Call SendHP(n)

                            Case SPELL_TYPE_ADDMP
                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                                Call SendMP(n)

                            Case SPELL_TYPE_ADDSP
                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                                Call SendMP(n)
                        End Select

                        ' Take away the mana points
                        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                        Call SendMP(Index)
                        Casted = True
                    Else
                        Call BattleMsg(Index, "Could not cast spell!", BRIGHTRED, 0)
                    End If
                End If
            Else
                Call PlayerMsg(Index, "Cannot cast spell!", BRIGHTRED)
            End If
        Else
            If CInt(Sqr((GetPlayerX(Index) - MapNPC(GetPlayerMap(Index), n).X) ^ 2 + ((GetPlayerY(Index) - MapNPC(GetPlayerMap(Index), n).Y) ^ 2))) > Spell(SpellNum).Range Then
                Call BattleMsg(Index, "You are too far away to hit the target.", BRIGHTRED, 0)
                Exit Sub
            End If

            Player(Index).TargetType = TARGET_TYPE_NPC

            If NPC(MapNPC(GetPlayerMap(Index), n).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(GetPlayerMap(Index), n).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)

                Select Case Spell(SpellNum).Type
                    Case SPELL_TYPE_ADDHP
                        MapNPC(GetPlayerMap(Index), n).HP = MapNPC(GetPlayerMap(Index), n).HP + Spell(SpellNum).Data1

                    Case SPELL_TYPE_SUBHP

                        Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - Int(NPC(MapNPC(GetPlayerMap(Index), n).num).DEF / 2)
                        If Damage > 0 Then
                            If Spell(SpellNum).Element <> 0 And NPC(MapNPC(GetPlayerMap(Index), n).num).Element <> 0 Then
                                If Element(Spell(SpellNum).Element).Strong = NPC(MapNPC(GetPlayerMap(Index), n).num).Element Or Element(NPC(MapNPC(GetPlayerMap(Index), n).num).Element).Weak = Spell(SpellNum).Element Then
                                    Call BattleMsg(Index, "A deadly mix of elements harm the " & Trim$(NPC(MapNPC(GetPlayerMap(Index), n).num).Name) & "!", BLUE, 0)
                                    Damage = Int(Damage * 1.25)
                                    If Element(Spell(SpellNum).Element).Strong = NPC(MapNPC(GetPlayerMap(Index), n).num).Element And Element(NPC(MapNPC(GetPlayerMap(Index), n).num).Element).Weak = Spell(SpellNum).Element Then
                                        Damage = Int(Damage * 1.2)
                                    End If
                                End If

                                If Element(Spell(SpellNum).Element).Weak = NPC(MapNPC(GetPlayerMap(Index), n).num).Element Or Element(NPC(MapNPC(GetPlayerMap(Index), n).num).Element).Strong = Spell(SpellNum).Element Then
                                    Call BattleMsg(Index, "The " & Trim$(NPC(MapNPC(GetPlayerMap(Index), n).num).Name) & " aborbs much of the elemental damage!", RED, 0)
                                    Damage = Int(Damage * 0.75)
                                    If Element(Spell(SpellNum).Element).Weak = NPC(MapNPC(GetPlayerMap(Index), n).num).Element And Element(NPC(MapNPC(GetPlayerMap(Index), n).num).Element).Strong = Spell(SpellNum).Element Then
                                        Damage = Int(Damage * (2 / 3))
                                    End If
                                End If
                            End If
                            Call AttackNpc(Index, n, Damage)
                        Else
                            Call BattleMsg(Index, "The spell was to weak to hurt " & Trim$(NPC(MapNPC(GetPlayerMap(Index), n).num).Name) & "!", BRIGHTRED, 0)
                        End If

                    Case SPELL_TYPE_ADDMP
                        MapNPC(GetPlayerMap(Index), n).MP = MapNPC(GetPlayerMap(Index), n).MP + Spell(SpellNum).Data1

                    Case SPELL_TYPE_SUBMP
                        MapNPC(GetPlayerMap(Index), n).MP = MapNPC(GetPlayerMap(Index), n).MP - Spell(SpellNum).Data1

                    Case SPELL_TYPE_ADDSP
                        MapNPC(GetPlayerMap(Index), n).SP = MapNPC(GetPlayerMap(Index), n).SP + Spell(SpellNum).Data1

                    Case SPELL_TYPE_SUBSP
                        MapNPC(GetPlayerMap(Index), n).SP = MapNPC(GetPlayerMap(Index), n).SP - Spell(SpellNum).Data1
                End Select

                ' Take away the mana points
                Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                Call SendMP(Index)
                Casted = True
            Else
                Call BattleMsg(Index, "Could not cast spell!", BRIGHTRED, 0)
            End If
        End If
    End If

    If Casted = True Then
        Player(Index).AttackTimer = GetTickCount
        Player(Index).CastedSpell = YES
        Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & Player(Index).CastedSpell & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & END_CHAR)
    End If
End Sub

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    Dim I As Long
    Dim n As Long

    If GetPlayerWeaponSlot(Index) > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            I = Int(GetPlayerSTR(Index) / 2) + Int(GetPlayerLevel(Index) / 2)

            n = Int(Rnd * 100) + 1
            If n <= I Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim I As Long
    Dim n As Long

    If GetPlayerShieldSlot(Index) > 0 Then
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

Public Sub CheckEquippedItems(ByVal Index As Long)
    Dim ItemNum As Long

    ' Check to make sure the weapon exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
            If Item(ItemNum).Type <> ITEM_TYPE_TWO_HAND Then
                Call SetPlayerWeaponSlot(Index, 0)
            End If
        End If
    Else
        Call SetPlayerWeaponSlot(Index, 0)
    End If

    ' Check to make sure the chest armor exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then
            Call SetPlayerArmorSlot(Index, 0)
        End If
    Else
        Call SetPlayerArmorSlot(Index, 0)
    End If

    ' Check to make sure the helmet exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then
            Call SetPlayerHelmetSlot(Index, 0)
        End If
    Else
        Call SetPlayerHelmetSlot(Index, 0)
    End If

    ' Check to make sure the shield exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then
            Call SetPlayerShieldSlot(Index, 0)
        End If
    Else
        Call SetPlayerShieldSlot(Index, 0)
    End If

    ' Check to make sure the leggings exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_LEGS Then
            Call SetPlayerLegsSlot(Index, 0)
        End If
    Else
        Call SetPlayerLegsSlot(Index, 0)
    End If

    ' Check to make sure the ring exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_RING Then
            Call SetPlayerRingSlot(Index, 0)
        End If
    Else
        Call SetPlayerRingSlot(Index, 0)
    End If

    ' Check to make sure the necklace exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_NECKLACE Then
            Call SetPlayerNecklaceSlot(Index, 0)
        End If
    Else
        Call SetPlayerNecklaceSlot(Index, 0)
    End If
End Sub

Public Sub SetPMember(ByVal LeaderIndex As Long, ByVal MemberIndex As Long)
    Dim I As Integer

    For I = 1 To MAX_PARTY_MEMBERS
        If Player(LeaderIndex).Party.Member(I) = 0 Then
            Player(LeaderIndex).Party.Member(I) = MemberIndex
            Exit For
        End If
    Next I

    Player(MemberIndex).Party.Leader = LeaderIndex

    For I = 1 To MAX_PARTY_MEMBERS
        If Player(LeaderIndex).Party.Member(I) > 0 Then
            Call UpdateParty(Player(LeaderIndex).Party.Member(I))
        End If
    Next I
End Sub

' This sub-routine needs a re-write. [Mellowz]
Public Sub RemovePMember(ByVal Index As Long)
    Dim I As Long
    Dim b As Long
    Dim q As Long

    b = Player(Index).Party.Leader

    ' Change the party leader.
    If Player(Index).Party.Leader = Index Then
        For I = 1 To MAX_PARTY_MEMBERS
            If Player(Index).Party.Member(I) > 0 Then
                If Player(Index).Party.Member(I) <> Index Then
                    Call ChangePLeader(Player(Index).Party.Member(I))
                    Exit For
                End If
            End If
        Next I
    End If

    ' Find which member the player is.
    For q = 1 To MAX_PARTY_MEMBERS
        If Player(Index).Party.Member(q) = Index Then
            Exit For
        End If
    Next q

    For I = 1 To MAX_PARTY_MEMBERS ' removes player from other members party
        If Player(Index).Party.Member(I) > 0 Then
            Player(Player(Index).Party.Member(I)).Party.Member(q) = 0
        End If
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
            If Player(b).Party.Member(I) > 0 Then
                q = q + 1
            End If
        Next I

        If q < 1 Then
            Call PlayerMsg(b, "The party has been disbanded.", WHITE)
            Player(b).InParty = False

            For I = 1 To MAX_PARTY_MEMBERS ' clears player's party
                Player(b).Party.Member(I) = 0
            Next I
        End If
    End If

    For I = 1 To MAX_PARTY_MEMBERS
        If Player(Index).Party.Member(I) > 0 And Player(Index).Party.Member(I) = Index Then
            Call SendDataTo(Index, "removemembers" & SEP_CHAR & END_CHAR)
        End If
        If Player(Index).Party.Member(I) > 0 And Player(Index).Party.Member(I) <> Index Then
            Call SendDataTo(Index, "updatemembers" & SEP_CHAR & I & SEP_CHAR & 0 & END_CHAR)
        End If
    Next I
End Sub

Public Sub ChangePLeader(ByVal Index As Long)
    Dim I As Integer

    Player(Index).Party.Leader = Index

    For I = 1 To MAX_PARTY_MEMBERS
        If Player(Index).Party.Member(I) > 0 Then
            Player(Player(Index).Party.Member(I)).Party.Leader = Index

            Call PlayerMsg(Player(Index).Party.Member(I), "Leadership has been passed to " & GetPlayerName(Index) & "!", PINK)
        End If
    Next I
End Sub

Public Sub UpdateParty(ByVal Index As Long)
    Player(Index).Party = Player(Player(Index).Party.Leader).Party
End Sub

Public Sub SetPShare(ByVal Index As Long, ByVal share As Boolean)
    Player(Index).Party.ShareExp = share
End Sub

Function GetPLeader(ByVal Index As Long) As Long
    GetPLeader = Player(Index).Party.Leader
End Function

Function GetPMember(ByVal Index As Long, ByVal Member As Long) As Long
    GetPMember = Player(Index).Party.Member(Member)
End Function

Function GetPShare(ByVal Index As Long) As Boolean
    GetPShare = Player(Index).Party.ShareExp
End Function

Public Sub ShowPLR(ByVal Index As Long)
    Dim LS As ListItem

    On Error Resume Next

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(Index) Then
        frmServer.lvUsers.ListItems.Remove Index
    End If

    Set LS = frmServer.lvUsers.ListItems.Add(Index, , Index)

    If IsPlaying(Index) Then
        LS.SubItems(1) = GetPlayerLogin(Index)
        LS.SubItems(2) = GetPlayerName(Index)
        LS.SubItems(3) = GetPlayerLevel(Index)
        LS.SubItems(4) = GetPlayerSprite(Index)
        LS.SubItems(5) = GetPlayerAccess(Index)
    End If
End Sub

Public Sub RemovePLR(ByVal Index As Long)
    Dim LS As ListItem
    
    On Error Resume Next

    If Not IsPlaying(Index) Then
        frmServer.lvUsers.ListItems.Remove Index
    
        Set LS = frmServer.lvUsers.ListItems.Add(Index, , Index)
        
        LS.SubItems(1) = vbNullString
        LS.SubItems(2) = vbNullString
        LS.SubItems(3) = vbNullString
        LS.SubItems(4) = vbNullString
        LS.SubItems(5) = vbNullString
    End If
End Sub

Function CanAttackPlayerWithArrow(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    ' Check If map Is attackable
    If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
        ' Make sure they are high enough level
        If GetPlayerLevel(Attacker) < 10 Then
            Call PlayerMsg(Attacker, "Your level is below 10 you can't attack anybody until your level 10 or higher.", BRIGHTRED)
        Else
            If GetPlayerLevel(Victim) < 10 Then
                Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is lower then level 10 for that you can't attack him.", BRIGHTRED)
            Else
                If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                    If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                        CanAttackPlayerWithArrow = True
                    Else
                        Call PlayerMsg(Attacker, "Is in the same guild as you are for that you can't attack him.", BRIGHTRED)
                    End If
                Else
                    CanAttackPlayerWithArrow = True
                End If
            End If
        End If
    Else
        Call PlayerMsg(Attacker, "This is a safe zone!", BRIGHTRED)
    End If
End Function

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim MapNum As Long, NPCnum As Long
    Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    CanAttackNpcWithArrow = False

    ' Check For subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check For subscript out of range
    If MapNPC(GetPlayerMap(Attacker), MapNpcNum).num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    NPCnum = MapNPC(MapNum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNPC(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure they are On the same map
    If IsPlaying(Attacker) Then
        If NPCnum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
            ' Check If at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpcWithArrow = True
                    Else
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(NPCnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(NPCnum).Name) & " :" & Trim$(NPC(NPCnum).AttackSay), GREEN)
                        End If
                    End If

                Case DIR_DOWN
                    If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpcWithArrow = True
                    Else
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(NPCnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(NPCnum).Name) & " :" & Trim$(NPC(NPCnum).AttackSay), GREEN)
                        End If
                    End If

                Case DIR_LEFT
                    If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpcWithArrow = True
                    Else
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(NPCnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(NPCnum).Name) & " :" & Trim$(NPC(NPCnum).AttackSay), GREEN)
                        End If
                    End If

                Case DIR_RIGHT
                    If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpcWithArrow = True
                    Else
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(NPCnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(NPCnum).Name) & " :" & Trim$(NPC(NPCnum).AttackSay), GREEN)
                        End If
                    End If
            End Select
        End If
    End If
End Function

Sub SendIndexWornEquipment(ByVal Index As Long)
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    If GetPlayerArmorSlot(Index) > 0 Then
        Armor = GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        Helmet = GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        Shield = GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))
    End If

    If GetPlayerWeaponSlot(Index) > 0 Then
        Weapon = GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))
    End If

    If GetPlayerLegsSlot(Index) > 0 Then
        Legs = GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))
    End If

    If GetPlayerRingSlot(Index) > 0 Then
        Ring = GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))
    End If

    If GetPlayerNecklaceSlot(Index) > 0 Then
        Necklace = GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))
    End If

    Call SendDataToMap(GetPlayerMap(Index), "itemworn" & SEP_CHAR & Index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR)
End Sub

Sub SendIndexWornEquipmentto(ByVal Index As Long, ByVal From As Long)
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    If GetPlayerArmorSlot(From) > 0 Then
        Armor = GetPlayerInvItemNum(From, GetPlayerArmorSlot(From))
    End If

    If GetPlayerHelmetSlot(From) > 0 Then
        Helmet = GetPlayerInvItemNum(From, GetPlayerHelmetSlot(From))
    End If

    If GetPlayerShieldSlot(From) > 0 Then
        Shield = GetPlayerInvItemNum(From, GetPlayerShieldSlot(Index))
    End If

    If GetPlayerWeaponSlot(From) > 0 Then
        Weapon = GetPlayerInvItemNum(From, GetPlayerWeaponSlot(From))
    End If

    If GetPlayerLegsSlot(From) > 0 Then
        Legs = GetPlayerInvItemNum(From, GetPlayerLegsSlot(From))
    End If

    If GetPlayerRingSlot(From) > 0 Then
        Ring = GetPlayerInvItemNum(From, GetPlayerRingSlot(From))
    End If

    If GetPlayerNecklaceSlot(From) > 0 Then
        Necklace = GetPlayerInvItemNum(From, GetPlayerNecklaceSlot(From))
    End If

    Call SendDataTo(Index, "itemworn" & SEP_CHAR & From & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR)
End Sub


Sub SendIndexWornEquipmentFromMap(ByVal Index As Long)
    Dim packet As String
    Dim I As Long
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) = True Then
            If GetPlayerMap(I) = GetPlayerMap(Index) Then

                Armor = 0
                Helmet = 0
                Shield = 0
                Weapon = 0
                Legs = 0
                Ring = 0
                Necklace = 0

                If GetPlayerArmorSlot(I) > 0 Then
                    Armor = GetPlayerInvItemNum(I, GetPlayerArmorSlot(I))
                End If
                If GetPlayerHelmetSlot(I) > 0 Then
                    Helmet = GetPlayerInvItemNum(I, GetPlayerHelmetSlot(I))
                End If
                If GetPlayerShieldSlot(I) > 0 Then
                    Shield = GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))
                End If
                If GetPlayerWeaponSlot(I) > 0 Then
                    Weapon = GetPlayerInvItemNum(I, GetPlayerWeaponSlot(I))
                End If
                If GetPlayerLegsSlot(I) > 0 Then
                    Legs = GetPlayerInvItemNum(I, GetPlayerLegsSlot(I))
                End If
                If GetPlayerRingSlot(I) > 0 Then
                    Ring = GetPlayerInvItemNum(I, GetPlayerRingSlot(I))
                End If
                If GetPlayerNecklaceSlot(I) > 0 Then
                    Necklace = GetPlayerInvItemNum(I, GetPlayerNecklaceSlot(I))
                End If

                packet = "itemworn" & SEP_CHAR & I & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR
                Call SendDataTo(Index, packet)
            End If
        End If
    Next I
End Sub

Sub AddNewTimer(ByVal Name As String, ByVal Interval As Long)
    On Error Resume Next
    Dim TmpTimer As clsCTimers
    Set TmpTimer = New clsCTimers
    TmpTimer.Name = Name
    TmpTimer.Interval = Interval
    TmpTimer.tmrWait = GetTickCount + Interval
    CTimers.Add TmpTimer, Name
    If Err.Number > 0 Then
        Debug.Print "Err: " & Err.Number
        CTimers.Item(Name).Name = Name
        CTimers.Item(Name).Interval = Interval
        CTimers.Item(Name).tmrWait = GetTickCount + Interval
        Err.Clear
    End If
End Sub

Function GetTimeLeft(ByVal Name As String) As Long
    On Error GoTo Hell
    GetTimeLeft = CTimers.Item(Name).tmrWait - GetTickCount
    Exit Function
Hell:
    GetTimeLeft = -1
End Function

Sub GetRidOfTimer(ByVal Name As String)
    Call CTimers.Remove(Name)
End Sub
Sub ScriptSetTile(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal setx As Long, ByVal sety As Long, ByVal tileset As Long, ByVal layer As Long)
    Dim packet As String
    packet = "tilecheck" & SEP_CHAR & mapper & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & layer & SEP_CHAR

    Select Case layer

        Case 0
            Map(mapper).Tile(X, Y).Ground = sety * 14 + setx
            Map(mapper).Tile(X, Y).GroundSet = tileset
            packet = packet & Map(mapper).Tile(X, Y).Ground & SEP_CHAR & Map(mapper).Tile(X, Y).GroundSet

        Case 1
            Map(mapper).Tile(X, Y).Mask = sety * 14 + setx
            Map(mapper).Tile(X, Y).MaskSet = tileset
            packet = packet & Map(mapper).Tile(X, Y).Mask & SEP_CHAR & Map(mapper).Tile(X, Y).MaskSet

        Case 2
            Map(mapper).Tile(X, Y).Anim = sety * 14 + setx
            Map(mapper).Tile(X, Y).AnimSet = tileset
            packet = packet & Map(mapper).Tile(X, Y).Anim & SEP_CHAR & Map(mapper).Tile(X, Y).AnimSet

        Case 3
            Map(mapper).Tile(X, Y).Mask2 = sety * 14 + setx
            Map(mapper).Tile(X, Y).Mask2Set = tileset
            packet = packet & Map(mapper).Tile(X, Y).Mask2 & SEP_CHAR & Map(mapper).Tile(X, Y).Mask2Set

        Case 4
            Map(mapper).Tile(X, Y).M2Anim = sety * 14 + setx
            Map(mapper).Tile(X, Y).M2AnimSet = tileset
            packet = packet & Map(mapper).Tile(X, Y).M2Anim & SEP_CHAR & Map(mapper).Tile(X, Y).M2AnimSet

        Case 5
            Map(mapper).Tile(X, Y).Fringe = sety * 14 + setx
            Map(mapper).Tile(X, Y).FringeSet = tileset
            packet = packet & Map(mapper).Tile(X, Y).Fringe & SEP_CHAR & Map(mapper).Tile(X, Y).FringeSet

        Case 6
            Map(mapper).Tile(X, Y).FAnim = sety * 14 + setx
            Map(mapper).Tile(X, Y).FAnimSet = tileset
            packet = packet & Map(mapper).Tile(X, Y).FAnim & SEP_CHAR & Map(mapper).Tile(X, Y).FAnimSet

        Case 7
            Map(mapper).Tile(X, Y).Fringe2 = sety * 14 + setx
            Map(mapper).Tile(X, Y).Fringe2Set = tileset
            packet = packet & Map(mapper).Tile(X, Y).Fringe2 & SEP_CHAR & Map(mapper).Tile(X, Y).Fringe2Set

        Case 8
            Map(mapper).Tile(X, Y).F2Anim = sety * 14 + setx
            Map(mapper).Tile(X, Y).F2AnimSet = tileset
            packet = packet & Map(mapper).Tile(X, Y).F2Anim & SEP_CHAR & Map(mapper).Tile(X, Y).F2AnimSet
    End Select

    Call SaveMap(mapper)
    Call SendDataToAll(packet & END_CHAR)
End Sub

Sub ScriptSetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
    Dim packet As String
    
    With Map(mapper).Tile(X, Y)
        .Type = Attrib
        .Data1 = Data1
        .Data2 = Data2
        .Data3 = Data3
        .String1 = String1
        .String2 = String2
        .String3 = String3
    End With

    packet = "tilecheckattribute" & SEP_CHAR & mapper & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR
    With Map(mapper).Tile(X, Y)
        packet = packet & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR
    End With
    
    Call SaveMap(mapper)
    Call SendDataToAll(packet & END_CHAR)
End Sub

Function ItemIsUsable(ByVal Index As Long, ByVal InvNum As Long) As Boolean
    ' Check if the player meets the class requirement.
    If Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq > -1 Then
        If GetPlayerClass(Index) <> Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq Then
            Call PlayerMsg(Index, "You must be a " & GetClassName(Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq) & " to use this item!", BRIGHTRED)
            Exit Function
        End If
    End If

    ' Check if the player meets the access requirement.
    If GetPlayerAccess(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).AccessReq Then
        Call PlayerMsg(Index, "Your access must be higher then " & Item(GetPlayerInvItemNum(Index, InvNum)).AccessReq & "!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the strength requirement.
    If GetPlayerSTR(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).StrReq Then
        Call PlayerMsg(Index, "Your strength is too low to equip this item!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the defense requirement.
    If GetPlayerDEF(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).DefReq Then
        Call PlayerMsg(Index, "Your defense is too low to equip this item!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the magic requirement.
    If GetPlayerMAGI(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).MagicReq Then
        Call PlayerMsg(Index, "Your magic is too low to equip this item!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the speed requirement.
    If GetPlayerSPEED(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).SpeedReq Then
        Call PlayerMsg(Index, "Your speed is too low to equip this item!", BRIGHTRED)
        Exit Function
    End If

    ItemIsUsable = True
End Function

Function ItemIsEquipped(ByVal Index As Long, ByVal ItemNum As Long) As Boolean
    If GetPlayerWeaponSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerArmorSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerShieldSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerHelmetSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerLegsSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerRingSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerNecklaceSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If
End Function
