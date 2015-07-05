Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal index As Long) As Long
Dim WeaponSlot As Long, RingSlot As Long, NecklaceSlot As Long
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
        
        If GetPlayerInvItemDur(index, WeaponSlot) > -1 Then
            Call SetPlayerInvItemDur(index, WeaponSlot, GetPlayerInvItemDur(index, WeaponSlot) - 1)
        
            If GetPlayerInvItemDur(index, WeaponSlot) = 0 Then
                Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, WeaponSlot), 0)
            Else
                If GetPlayerInvItemDur(index, WeaponSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, WeaponSlot) & "/" & Trim(Item(GetPlayerInvItemNum(index, WeaponSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If GetPlayerRingSlot(index) > 0 Then
        RingSlot = GetPlayerRingSlot(index)
        
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(index, RingSlot)).Data2
        
        If GetPlayerInvItemDur(index, RingSlot) > -1 Then
            Call SetPlayerInvItemDur(index, RingSlot, GetPlayerInvItemDur(index, RingSlot) - 1)
        
            If GetPlayerInvItemDur(index, RingSlot) = 0 Then
                Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, RingSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, RingSlot), 0)
            Else
                If GetPlayerInvItemDur(index, RingSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, RingSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, RingSlot) & "/" & Trim(Item(GetPlayerInvItemNum(index, RingSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If GetPlayerNecklaceSlot(index) > 0 Then
        NecklaceSlot = GetPlayerNecklaceSlot(index)
        
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(index, NecklaceSlot)).Data2
        
        If GetPlayerInvItemDur(index, NecklaceSlot) > -1 Then
            Call SetPlayerInvItemDur(index, NecklaceSlot, GetPlayerInvItemDur(index, NecklaceSlot) - 1)
        
            If GetPlayerInvItemDur(index, NecklaceSlot) = 0 Then
                Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, NecklaceSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, NecklaceSlot), 0)
            Else
                If GetPlayerInvItemDur(index, NecklaceSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, NecklaceSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, NecklaceSlot) & "/" & Trim(Item(GetPlayerInvItemNum(index, NecklaceSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If



    If GetPlayerDamage < 0 Then
        GetPlayerDamage = 0
    End If
End Function

Function GetPlayerProtection(ByVal index As Long) As Long
Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long, LegsSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ArmorSlot = GetPlayerArmorSlot(index)
    HelmSlot = GetPlayerHelmetSlot(index)
    ShieldSlot = GetPlayerShieldSlot(index)
    LegsSlot = GetPlayerLegsSlot(index)
    GetPlayerProtection = Int(GetPlayerDEF(index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, ArmorSlot)).Data2
        If GetPlayerInvItemDur(index, ArmorSlot) > -1 Then
            Call SetPlayerInvItemDur(index, ArmorSlot, GetPlayerInvItemDur(index, ArmorSlot) - 1)
        
            If GetPlayerInvItemDur(index, ArmorSlot) = 0 Then
                Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, ArmorSlot), 0)
            Else
                If GetPlayerInvItemDur(index, ArmorSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, ArmorSlot) & "/" & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, HelmSlot)).Data2
        If GetPlayerInvItemDur(index, HelmSlot) > -1 Then
            Call SetPlayerInvItemDur(index, HelmSlot, GetPlayerInvItemDur(index, HelmSlot) - 1)

            If GetPlayerInvItemDur(index, HelmSlot) <= 0 Then
                Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, HelmSlot), 0)
            Else
                If GetPlayerInvItemDur(index, HelmSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, HelmSlot) & "/" & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, ShieldSlot)).Data2
        If GetPlayerInvItemDur(index, ShieldSlot) > -1 Then
            Call SetPlayerInvItemDur(index, ShieldSlot, GetPlayerInvItemDur(index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(index, ShieldSlot) <= 0 Then
                Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ShieldSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, ShieldSlot), 0)
            Else
                If GetPlayerInvItemDur(index, ShieldSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, ShieldSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, ShieldSlot) & "/" & Trim(Item(GetPlayerInvItemNum(index, ShieldSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
    
    If LegsSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, LegsSlot)).Data2
        If GetPlayerInvItemDur(index, LegsSlot) > -1 Then
            Call SetPlayerInvItemDur(index, LegsSlot, GetPlayerInvItemDur(index, LegsSlot) - 1)

            If GetPlayerInvItemDur(index, LegsSlot) <= 0 Then
                Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, LegsSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, LegsSlot), 0)
            Else
                If GetPlayerInvItemDur(index, LegsSlot) <= 10 Then
                    Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, LegsSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, LegsSlot) & "/" & Trim(Item(GetPlayerInvItemNum(index, LegsSlot)).Data1), Yellow, 0)
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

Public Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim I As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(index, I) = ItemNum Then
                FindOpenInvSlot = I
                Exit Function
            End If
        Next I
    End If
    
    For I = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, I) = 0 Then
            FindOpenInvSlot = I
            Exit Function
        End If
    Next I
End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim I As Long
   
   FindOpenBankSlot = 0
   
   ' Check for subscript out of range
   If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
       Exit Function
   End If
   
   If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
       ' If currency then check to see if they already have an instance of the item and add it to that
       For I = 1 To MAX_BANK
           If GetPlayerBankItemNum(index, I) = ItemNum Then
               FindOpenBankSlot = I
               Exit Function
           End If
       Next I
   End If
   
   For I = 1 To MAX_BANK
       ' Try to find an open free slot
       If GetPlayerBankItemNum(index, I) = 0 Then
           FindOpenBankSlot = I
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

Function FindOpenSpellSlot(ByVal index As Long) As Long
Dim I As Long

    FindOpenSpellSlot = 0
    
    For I = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, I) = 0 Then
            FindOpenSpellSlot = I
            Exit Function
        End If
    Next I
End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
Dim I As Long

    HasSpell = False
    
    For I = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, I) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next I
End Function
'Function TotalOnlinePlayers() As Long
'Dim I As Long
'TotalOnlinePlayers = 0
'
'For I = 1 To MAX_PLAYERS
'    If IsPlaying(I) Then
'        TotalOnlinePlayers = TotalOnlinePlayers + 1
'    End If
'Next I
'End Function

Function TotalOnlinePlayers() As Long
'On Error GoTo ErrorHandler
Dim I As Long
    'frmServer.LstPlayers.Clear
    'frmServer.LstAccounts.Clear
    TotalOnlinePlayers = 0

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
            'frmServer.LstPlayers.AddItem Trim(Player(i).Char(Player(i).CharNum).Name)
            'frmServer.LstAccounts.AddItem Trim(Player(i).Login)
        End If
    Next I
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  'ReportError "modGameLogic.bas", "TotalOnlinePlayers", Err.Number, Err.Description
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

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
Dim I As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    For I = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                HasItem = GetPlayerInvItemValue(index, I)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next I
End Function

Sub TakeItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim I As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For I = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, I) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(index, I, GetPlayerInvItemValue(index, I) - ItemVal)
                    Call SendInventoryUpdate(index, I)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(index, I)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(index) > 0 Then
                            If I = GetPlayerWeaponSlot(index) Then
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
                            If I = GetPlayerArmorSlot(index) Then
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
                            If I = GetPlayerHelmetSlot(index) Then
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
                            If I = GetPlayerShieldSlot(index) Then
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
                        
                    Case ITEM_TYPE_LEGS
                        If GetPlayerLegsSlot(index) > 0 Then
                            If I = GetPlayerLegsSlot(index) Then
                                Call SetPlayerLegsSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerLegsSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                    Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(index) > 0 Then
                            If I = GetPlayerRingSlot(index) Then
                                Call SetPlayerRingSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerRingSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                    Case ITEM_TYPE_NECKLACE
                        If GetPlayerNecklaceSlot(index) > 0 Then
                            If I = GetPlayerNecklaceSlot(index) Then
                                Call SetPlayerNecklaceSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select

                
                n = Item(GetPlayerInvItemNum(index, I)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_LEGS) And (n <> ITEM_TYPE_RING) And (n <> ITEM_TYPE_NECKLACE) Then
                    TakeItem = True
                End If
            End If
                            
            If TakeItem = True Then
                Call SetPlayerInvItemNum(index, I, 0)
                Call SetPlayerInvItemValue(index, I, 0)
                Call SetPlayerInvItemDur(index, I, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(index, I)
                Exit Sub
            End If
        End If
    Next I
End Sub

Sub GiveItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim I As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    I = FindOpenInvSlot(index, ItemNum)
    
    ' Check to see if inventory is full
    If I <> 0 Then
        Call SetPlayerInvItemNum(index, I, ItemNum)
        Call SetPlayerInvItemValue(index, I, GetPlayerInvItemValue(index, I) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_NECKLACE) Then
            Call SetPlayerInvItemDur(index, I, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(index, I)
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
    End If
End Sub

Sub TakeBankItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim I As Long, n As Long
Dim TakeBankItem As Boolean

   TakeBankItem = False
   
   ' Check for subscript out of range
   If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
       Exit Sub
   End If
   
   For I = 1 To MAX_BANK
       ' Check to see if the player has the item
       If GetPlayerBankItemNum(index, I) = ItemNum Then
           If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
               ' Is what we are trying to take away more then what they have? If so just set it to zero
               If ItemVal >= GetPlayerBankItemValue(index, I) Then
                   TakeBankItem = True
               Else
                   Call SetPlayerBankItemValue(index, I, GetPlayerBankItemValue(index, I) - ItemVal)
                   Call SendBankUpdate(index, I)
               End If
           Else
               ' Check to see if its any sort of ArmorSlot/WeaponSlot
               Select Case Item(GetPlayerBankItemNum(index, I)).Type
                   Case ITEM_TYPE_WEAPON
                       If GetPlayerWeaponSlot(index) > 0 Then
                           If I = GetPlayerWeaponSlot(index) Then
                               Call SetPlayerWeaponSlot(index, 0)
                               Call SendWornEquipment(index)
                               TakeBankItem = True
                           Else
                               ' Check if the item we are taking isn't already equipped
                               If ItemNum <> GetPlayerBankItemNum(index, GetPlayerWeaponSlot(index)) Then
                                   TakeBankItem = True
                               End If
                           End If
                       Else
                           TakeBankItem = True
                       End If
               
                   Case ITEM_TYPE_ARMOR
                       If GetPlayerArmorSlot(index) > 0 Then
                           If I = GetPlayerArmorSlot(index) Then
                               Call SetPlayerArmorSlot(index, 0)
                               Call SendWornEquipment(index)
                               TakeBankItem = True
                           Else
                               ' Check if the item we are taking isn't already equipped
                               If ItemNum <> GetPlayerBankItemNum(index, GetPlayerArmorSlot(index)) Then
                                   TakeBankItem = True
                               End If
                           End If
                       Else
                           TakeBankItem = True
                       End If
                   
                   Case ITEM_TYPE_HELMET
                       If GetPlayerHelmetSlot(index) > 0 Then
                           If I = GetPlayerHelmetSlot(index) Then
                               Call SetPlayerHelmetSlot(index, 0)
                               Call SendWornEquipment(index)
                               TakeBankItem = True
                           Else
                               ' Check if the item we are taking isn't already equipped
                               If ItemNum <> GetPlayerBankItemNum(index, GetPlayerHelmetSlot(index)) Then
                                   TakeBankItem = True
                               End If
                           End If
                       Else
                           TakeBankItem = True
                       End If
                   
                   Case ITEM_TYPE_SHIELD
                       If GetPlayerShieldSlot(index) > 0 Then
                           If I = GetPlayerShieldSlot(index) Then
                               Call SetPlayerShieldSlot(index, 0)
                               Call SendWornEquipment(index)
                               TakeBankItem = True
                           Else
                               ' Check if the item we are taking isn't already equipped
                               If ItemNum <> GetPlayerBankItemNum(index, GetPlayerShieldSlot(index)) Then
                                   TakeBankItem = True
                               End If
                           End If
                       Else
                           TakeBankItem = True
                       End If
                       
                   Case ITEM_TYPE_LEGS
                       If GetPlayerLegsSlot(index) > 0 Then
                           If I = GetPlayerLegsSlot(index) Then
                               Call SetPlayerLegsSlot(index, 0)
                               Call SendWornEquipment(index)
                               TakeBankItem = True
                           Else
                               ' Check if the item we are taking isn't already equipped
                               If ItemNum <> GetPlayerBankItemNum(index, GetPlayerLegsSlot(index)) Then
                                   TakeBankItem = True
                               End If
                           End If
                       Else
                           TakeBankItem = True
                       End If
                       
                   Case ITEM_TYPE_RING
                       If GetPlayerRingSlot(index) > 0 Then
                           If I = GetPlayerRingSlot(index) Then
                               Call SetPlayerRingSlot(index, 0)
                               Call SendWornEquipment(index)
                               TakeBankItem = True
                           Else
                               ' Check if the item we are taking isn't already equipped
                               If ItemNum <> GetPlayerBankItemNum(index, GetPlayerRingSlot(index)) Then
                                   TakeBankItem = True
                               End If
                           End If
                       Else
                           TakeBankItem = True
                       End If
                       
                   Case ITEM_TYPE_NECKLACE
                       If GetPlayerNecklaceSlot(index) > 0 Then
                           If I = GetPlayerNecklaceSlot(index) Then
                               Call SetPlayerNecklaceSlot(index, 0)
                               Call SendWornEquipment(index)
                               TakeBankItem = True
                           Else
                               ' Check if the item we are taking isn't already equipped
                               If ItemNum <> GetPlayerBankItemNum(index, GetPlayerNecklaceSlot(index)) Then
                                   TakeBankItem = True
                               End If
                           End If
                       Else
                           TakeBankItem = True
                       End If
               End Select

               
               n = Item(GetPlayerBankItemNum(index, I)).Type
               ' Check if its not an equipable weapon, and if it isn't then take it away
               If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_LEGS) And (n <> ITEM_TYPE_RING) And (n <> ITEM_TYPE_NECKLACE) Then
                   TakeBankItem = True
               End If
           End If
                           
           If TakeBankItem = True Then
               Call SetPlayerBankItemNum(index, I, 0)
               Call SetPlayerBankItemValue(index, I, 0)
               Call SetPlayerBankItemDur(index, I, 0)
               
               ' Send the Bank update
               Call SendBankUpdate(index, I)
               Exit Sub
           End If
       End If
   Next I
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal BankSlot As Long)
Dim I As Long

   ' Check for subscript out of range
   If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
       Exit Sub
   End If
   
   I = BankSlot
   
   ' Check to see if Bankentory is full
   If I <> 0 Then
       Call SetPlayerBankItemNum(index, I, ItemNum)
       Call SetPlayerBankItemValue(index, I, GetPlayerBankItemValue(index, I) + ItemVal)
       
       If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_NECKLACE) Then
           Call SetPlayerBankItemDur(index, I, Item(ItemNum).Data1)
       End If
   Else
       Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Bank full!" & SEP_CHAR & END_CHAR)
   End If
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim I As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    I = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(I, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, x, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
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
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_NECKLACE) Then
                MapItem(MapNum, I).Dur = ItemDur
            Else
                MapItem(MapNum, I).Dur = 0
            End If
        Else
            MapItem(MapNum, I).Dur = 0
        End If
        
        MapItem(MapNum, I).x = x
        MapItem(MapNum, I).y = y
            
        Packet = "SPAWNITEM" & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
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
Dim x As Long
Dim y As Long
Dim I As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Spawn what we have
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            ' Check if the tile type is an item or a saved tile incase someone drops something
            
            If (map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If (Item(map(MapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(map(MapNum).Tile(x, y).Data1).Stackable = 1) And map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(map(MapNum).Tile(x, y).Data1, map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If
        Next x
    Next y
End Sub

Sub PlayerMapGetItem(ByVal index As Long)
Dim I As Long
Dim n As Long
Dim MapNum As Long
Dim msg As String

    If IsPlaying(index) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(index)
    
    For I = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, I).num > 0) And (MapItem(MapNum, I).num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, I).x = GetPlayerX(index)) And (MapItem(MapNum, I).y = GetPlayerY(index)) Then
                ' Find open slot
                n = FindOpenInvSlot(index, MapItem(MapNum, I).num)
                
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(index, n, MapItem(MapNum, I).num)
                    If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, n)).Stackable = 1 Then
                        Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, I).Value)
                        msg = "You picked up " & MapItem(MapNum, I).Value & " " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(index, n, 0)
                        msg = "You picked up a " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    End If
                    Call SetPlayerInvItemDur(index, n, MapItem(MapNum, I).Dur)
                        
                    ' Erase item from the map
                    MapItem(MapNum, I).num = 0
                    MapItem(MapNum, I).Value = 0
                    MapItem(MapNum, I).Dur = 0
                    MapItem(MapNum, I).x = 0
                    MapItem(MapNum, I).y = 0
                        
                    Call SendInventoryUpdate(index, n)
                    Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    Call PlayerMsg(index, msg, Yellow)
                    Exit Sub
                Else
                    Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                    Exit Sub
                End If
            End If
        End If
    Next I
End Sub

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
Dim I As Long



    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
        I = FindOpenMapItemSlot(GetPlayerMap(index))
        
        If I <> 0 Then
            MapItem(GetPlayerMap(index), I).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(index) Then
                        Call SetPlayerArmorSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(index) Then
                        Call SetPlayerWeaponSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(index) Then
                        Call SetPlayerHelmetSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(index) Then
                        Call SetPlayerShieldSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
                Case ITEM_TYPE_LEGS
                    If InvNum = GetPlayerLegsSlot(index) Then
                        Call SetPlayerLegsSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
                Case ITEM_TYPE_RING
                    If InvNum = GetPlayerRingSlot(index) Then
                        Call SetPlayerRingSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
                Case ITEM_TYPE_NECKLACE
                    If InvNum = GetPlayerNecklaceSlot(index) Then
                        Call SetPlayerNecklaceSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(index), I).num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), I).x = GetPlayerX(index)
            MapItem(GetPlayerMap(index), I).y = GetPlayerY(index)
                        
            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), I).Value = GetPlayerInvItemValue(index, InvNum)
                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), I).Value = Amount
                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Amount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(index), I).Value = 0
                If Item(GetPlayerInvItemNum(index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_NECKLACE Then
                    If Item(GetPlayerInvItemNum(index, InvNum)).Data1 <= -1 Then
                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & " - Ind.", Yellow)
                    Else
                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & " - " & GetPlayerInvItemDur(index, InvNum) & "/" & Item(GetPlayerInvItemNum(index, InvNum)).Data1 & ".", Yellow)
                    End If
                Else
                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(I, MapItem(GetPlayerMap(index), I).num, Amount, MapItem(GetPlayerMap(index), I).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
             
        If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "onitemdrop " & index & "," & GetPlayerMap(index) & "," & MapItem(GetPlayerMap(index), I).num & "," & Amount & "," & MapItem(GetPlayerMap(index), I).Dur & "," & I & "," & InvNum
        End If
        
        Else
            Call PlayerMsg(index, "To many items already on the ground.", BrightRed)
        End If
    End If
End Sub



Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Packet As String
Dim NpcNum As Long
Dim I As Long, x As Long, y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    Spawned = False
    
    NpcNum = map(MapNum).Npc(MapNpcNum)
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
        
        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxhp(NpcNum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)
                
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)
        
        ' Well try 100 times to randomly place the sprite
        For I = 1 To 100
            x = Int(Rnd * MAX_MAPX)
            y = Int(Rnd * MAX_MAPY)
            
            ' Check if the tile is walkable
            If map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                MapNpc(MapNum, MapNpcNum).x = x
                MapNpc(MapNum, MapNpcNum).y = y
                Spawned = True
                Exit For
            End If
        Next I
        
        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    If map(MapNum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
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
                    If map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                            ' Check to make sure the victim isn't an admin
                            If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                                Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                            Else
                                ' Check if map is attackable
                                If map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                    ' Make sure they are high enough level
                                    If GetPlayerLevel(Attacker) < 10 Then
                                        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                    Else
                                        If GetPlayerLevel(Victim) < 10 Then
                                            Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
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
                        End If
                    ElseIf map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If

            Case DIR_DOWN
                If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    If map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure that they dont have access
                        If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                        Else
                            ' Check to make sure the victim isn't an admin
                            If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                                Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                            Else
                                ' Check if map is attackable
                                If map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                    ' Make sure they are high enough level
                                    If GetPlayerLevel(Attacker) < 10 Then
                                        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                    Else
                                        If GetPlayerLevel(Victim) < 10 Then
                                            Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
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
                            End If
                        End If
                    ElseIf map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
        
            Case DIR_LEFT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                    If map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure that they dont have access
                        If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                        Else
                            ' Check to make sure the victim isn't an admin
                            If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                                Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                            Else
                                ' Check if map is attackable
                                If map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                    ' Make sure they are high enough level
                                    If GetPlayerLevel(Attacker) < 10 Then
                                        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                    Else
                                        If GetPlayerLevel(Victim) < 10 Then
                                            Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
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
                            End If
                        End If
                    ElseIf map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
            
            Case DIR_RIGHT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                    If map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure that they dont have access
                        If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                        Else
                            ' Check to make sure the victim isn't an admin
                            If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                                Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                            Else
                                ' Check if map is attackable
                                If map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                    ' Make sure they are high enough level
                                    If GetPlayerLevel(Attacker) < 10 Then
                                        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                                    Else
                                        If GetPlayerLevel(Victim) < 10 Then
                                            Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
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
                            End If
                        End If
                    ElseIf map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
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
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
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
                If (MapNpc(MapNum, MapNpcNum).y + 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x = GetPlayerX(Attacker)) Then
 If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpc = True
Else
If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(NpcNum).SpawnSecs
Else
Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
End If
End If
End If
            Case DIR_DOWN
                If (MapNpc(MapNum, MapNpcNum).y - 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x = GetPlayerX(Attacker)) Then
 If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpc = True
Else
If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(NpcNum).SpawnSecs
Else
Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
End If
End If
End If
            Case DIR_LEFT
                If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x + 1 = GetPlayerX(Attacker)) Then
 If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpc = True
Else
If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(NpcNum).SpawnSecs
Else
Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
End If
End If
End If

Case DIR_RIGHT
                If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x - 1 = GetPlayerX(Attacker)) Then
 If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpc = True
Else
If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(NpcNum).SpawnSecs
Else
Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
End If
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

If map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA And map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA Then
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        'Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
        'Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)
    
        ' Player is dead
        'Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        
        If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "onPVPdeath " & Attacker & "," & Victim
        End If
        
        If map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
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
        
    End If
ElseIf map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then
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
        'Call GlobalMsg(GetPlayerName(Victim) & " has been killed in the arena by " & GetPlayerName(Attacker), BrightRed)
            
        ' Warp player away
        'Call PlayerWarp(Victim, map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data1, map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data2, map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data3)
        
        ' Restore vitals
        'Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        'Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        'Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        'Call SendHP(Victim)
        'Call SendMP(Victim)
        'Call SendSP(Victim)
                        
        ' Check if target is player who died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
        
        If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "onARENAdeath " & Attacker & "," & Victim
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
        
        If map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
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
Dim n As Long, I As Long, q As Integer, x As Long
Dim STR As Long, DEF As Long, MapNum As Long, NpcNum As Long

    ' Removes one SP when you attack, you can also set it to 2 or 3.
    If GetPlayerSP(Attacker) > 0 Then
    Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - 1)
    Call SendSP(Attacker)
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
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim(Npc(NpcNum).Name)
        
    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        
        'Call BattleMsg(Attacker, "You killed a " & Name, BrightRed, 0)
        If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "onNPCdeath " & Attacker & "," & MapNum & "," & NpcNum & "," & MapNpcNum
        End If

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
        If GetPlayerLegsSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerLegsSlot(Attacker))).AddEXP
        End If
        If GetPlayerRingSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerRingSlot(Attacker))).AddEXP
        End If
        If GetPlayerNecklaceSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerNecklaceSlot(Attacker))).AddEXP
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
            For x = 1 To MAX_PARTY_MEMBERS
            If Player(Attacker).Party.Member(x) > 0 Then q = q + 1
            Next x
            'MsgBox "in party" & q
            If q = 2 Then 'Remember, if they aren't in a party they'll only get one person, so this has to be at least two
                Exp = Exp * 0.75 ' 3/4 experience
                'MsgBox Exp
                For x = 1 To MAX_PARTY_MEMBERS
                    If Player(Attacker).Party.Member(x) > 0 Then
                        If Player(Player(Attacker).Party.Member(x)).Party.ShareExp = True Then
                            If GetPlayerLevel(Player(Attacker).Party.Member(x)) = MAX_LEVEL Then
                                Call SetPlayerExp(Player(Attacker).Party.Member(x), Experience(MAX_LEVEL))
                                Call BattleMsg(Player(Attacker).Party.Member(x), "You cant gain anymore experience!", BrightBlue, 0)
                            Else
                                Call SetPlayerExp(Player(Attacker).Party.Member(x), GetPlayerExp(Player(Attacker).Party.Member(x)) + Exp)
                                Call BattleMsg(Player(Attacker).Party.Member(x), "You have gained " & Exp & " party experience.", BrightBlue, 0)
                            End If
                        End If
                    End If
                Next x
            Else 'if there are 3 or more party members..
                Exp = Exp * 0.5  ' 1/2 experience
                    For x = 1 To MAX_PARTY_MEMBERS
                        If Player(Attacker).Party.Member(x) > 0 Then
                            If Player(Player(Attacker).Party.Member(x)).Party.ShareExp = True Then
                                If GetPlayerLevel(Player(Attacker).Party.Member(x)) = MAX_LEVEL Then
                                    Call SetPlayerExp(Player(Attacker).Party.Member(x), Experience(MAX_LEVEL))
                                    Call BattleMsg(Player(Attacker).Party.Member(x), "You cant gain anymore experience!", BrightBlue, 0)
                                Else
                                    Call SetPlayerExp(Player(Attacker).Party.Member(x), GetPlayerExp(n) + Exp)
                                    Call BattleMsg(Player(Attacker).Party.Member(x), "You have gained " & Exp & " party experience.", BrightBlue, 0)
                                End If
                            End If
                        End If
                    Next x
            End If
        End If
                                
        For I = 1 To MAX_NPC_DROPS
            ' Drop the goods if they get it
            n = Int(Rnd * Npc(NpcNum).ItemNPC(I).Chance) + 1
            If n = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(I).ItemNum, Npc(NpcNum).ItemNPC(I).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
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
            For x = 1 To MAX_PARTY_MEMBERS
                Call CheckPlayerLevelUp(Player(Attacker).Party.Member(x))
            Next x
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
                Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " : " & Trim(Npc(NpcNum).AttackSay) & "", SayColor)
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
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & map(MapNum).Revision & SEP_CHAR & END_CHAR)
    
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendIndexWornEquipmentFromMap(index)
    
    If Scripting = 1 Then
    MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & index
    End If
    
    Packet = "forcehouseclose" & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim MapNum As Long
Dim x As Long
Dim y As Long
Dim I As Long
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
                If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_DOOR) Or ((map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Or map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(index)).Up > 0 Then
                    Call PlayerWarp(index, map(GetPlayerMap(index)).Up, GetPlayerX(index), MAX_MAPY)
                    Moved = YES
                End If
            End If
                    
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < MAX_MAPY Then
                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_DOOR) Or ((map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Or map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                End If
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_DOOR) Or ((map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Or map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                    'BARON DEBUG TIME - CAUSE OF RTE 9'S ABOVE ?
                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(index)).left > 0 Then
                    Call PlayerWarp(index, map(GetPlayerMap(index)).left, MAX_MAPX, GetPlayerY(index))
                    Moved = YES
                End If
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < MAX_MAPX Then
                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If (map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_DOOR) Or ((map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Or map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                        
                        Packet = "PLAYERMOVE" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), Packet)
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                End If
            End If
    End Select
    
    If GetPlayerX(index) < 0 Or GetPlayerY(index) < 0 Or GetPlayerX(index) > MAX_MAPX Or GetPlayerY(index) > MAX_MAPY Or GetPlayerMap(index) <= 0 Then
        Call HackingAttempt(index, "")
        Exit Sub
    End If
    
    'healing tiles code
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SendHP(index)
        Call SendMP(index)
        Call PlayerMsg(index, "You feel a sudden rush through your body as you regain strength!", BrightGreen)
    End If
    
    'Check for kill tile, and if so kill them
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KILL Then
        Call SetPlayerHP(index, 0)
        Call PlayerMsg(index, "You embrace the cold finger of death; and feel your life extinguished", BrightRed)
        
        ' Warp player away
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & index
        Else
            Call PlayerWarp(index, START_MAP, START_X, START_Y)
        End If
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SetPlayerSP(index, GetPlayerMaxSP(index))
        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
        Moved = YES
    End If

    If GetPlayerX(index) + 1 <= MAX_MAPX Then
        If map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
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
        If map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
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
        If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Then
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
        If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Then
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
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_WARP Then
        MapNum = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        x = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        y = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3
                        
        Call PlayerWarp(index, MapNum, x, y)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KEYOPEN Then
        x = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        y = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        
        If map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
            TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            If Trim(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) = "" Then
                Call MapMsg(GetPlayerMap(index), "A door has been unlocked!", White)
            Else
                Call MapMsg(GetPlayerMap(index), Trim(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), White)
            End If
            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
        End If
    End If
        
    ' Check for shop
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SHOP Then
       If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 > 0 Then
            Call SendTrade(index, map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
        Else
            Call PlayerMsg(index, "There is no shop here.", BrightRed)
        End If
    End If
        
    ' Check if player stepped on sprite changing tile
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(index) = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            Call PlayerMsg(index, "You already have this sprite!", BrightRed)
            Exit Sub
        Else
            If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
                Call SendDataTo(index, "spritechange" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Else
                If Item(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(index, "This sprite will cost you " & map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3 & " " & Trim(Item(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", Yellow)
                Else
                    Call PlayerMsg(index, "This sprite will cost you a " & Trim(Item(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", Yellow)
                End If
                Call SendDataTo(index, "spritechange" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
        ' Check if player stepped on house buying tile

    ' Check if player stepped on house buying tile
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_HOUSE Then
        If Len(map(GetPlayerMap(index)).Owner) < 2 Then
        If GetPlayerName(index) = map(GetPlayerMap(index)).Owner Then
            Call PlayerMsg(index, "You already own this house!", BrightRed)
            Exit Sub
        Else
            If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 0 Then
                Call SendDataTo(index, "housebuy" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Else
                If Item(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(index, "This house will cost you " & map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 & " " & Trim(Item(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).Name) & "!", Yellow)
                Else
                    Call PlayerMsg(index, "This house will cost you a " & Trim(Item(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).Name) & "!", Yellow)
                End If
                Call SendDataTo(index, "housebuy" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            End If
        End If
            Else
    Call PlayerMsg(index, "This house is not for sale!", BrightRed)
    Exit Sub
    End If
    End If
    ' Check if player stepped on sprite changing tile
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_CLASS_CHANGE Then
        If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 > -1 Then
            If GetPlayerClass(index) <> map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
                Call PlayerMsg(index, "You arent the required class!", BrightRed)
                Exit Sub
            End If
        End If
        
        If GetPlayerClass(index) = map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            Call PlayerMsg(index, "You are already this class!", BrightRed)
        Else
            If Player(index).Char(Player(index).CharNum).Sex = 0 Then
                If GetPlayerSprite(index) = Class(GetPlayerClass(index)).MaleSprite Then
                    Call SetPlayerSprite(index, Class(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).MaleSprite)
                End If
            Else
                If GetPlayerSprite(index) = Class(GetPlayerClass(index)).FemaleSprite Then
                    Call SetPlayerSprite(index, Class(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).FemaleSprite)
                End If
            End If
            
            Call SetPlayerSTR(index, (Player(index).Char(Player(index).CharNum).STR - Class(GetPlayerClass(index)).STR))
            Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).DEF - Class(GetPlayerClass(index)).DEF))
            Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).Magi - Class(GetPlayerClass(index)).Magi))
            Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed - Class(GetPlayerClass(index)).Speed))
            
            Call SetPlayerClass(index, map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)

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
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_NOTICE Then
        If Trim(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) <> "" Then
            Call PlayerMsg(index, Trim(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), Black)
        End If
        If Trim(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2) <> "" Then
            Call PlayerMsg(index, Trim(map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2), Grey)
        End If
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String3 & SEP_CHAR & END_CHAR)
    End If
    
    ' Check if player stepped on sound tile
    If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1 & SEP_CHAR & END_CHAR)
    End If
    
    If Scripting = 1 Then
        If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & index & "," & map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        End If
    End If
    ' Check if player stepped on Bank tile
   If map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_BANK Then
       Call SendDataTo(index, "openbank" & SEP_CHAR & END_CHAR)
   End If

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
Dim I As Long, n As Long
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
                n = map(MapNum).Tile(x, y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum, MapNpcNum).x) And (GetPlayerY(I) = MapNpc(MapNum, MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next I
    
                If CanNPCMoveAttributeNPC(MapNum, MapNpcNum, DIR_UP) = False Then
                    CanNpcMove = False
                    Exit Function
                End If
                                
                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> MapNpcNum) And (MapNpc(MapNum, I).num > 0) And (MapNpc(MapNum, I).x = MapNpc(MapNum, MapNpcNum).x) And (MapNpc(MapNum, I).y = MapNpc(MapNum, MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next I
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If y < MAX_MAPY Then
                n = map(MapNum).Tile(x, y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum, MapNpcNum).x) And (GetPlayerY(I) = MapNpc(MapNum, MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next I
                
                If CanNPCMoveAttributeNPC(MapNum, MapNpcNum, DIR_DOWN) = False Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> MapNpcNum) And (MapNpc(MapNum, I).num > 0) And (MapNpc(MapNum, I).x = MapNpc(MapNum, MapNpcNum).x) And (MapNpc(MapNum, I).y = MapNpc(MapNum, MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next I
            Else
                CanNpcMove = False
            End If
                
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = map(MapNum).Tile(x - 1, y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum, MapNpcNum).x - 1) And (GetPlayerY(I) = MapNpc(MapNum, MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next I
                
                If CanNPCMoveAttributeNPC(MapNum, MapNpcNum, DIR_LEFT) = False Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> MapNpcNum) And (MapNpc(MapNum, I).num > 0) And (MapNpc(MapNum, I).x = MapNpc(MapNum, MapNpcNum).x - 1) And (MapNpc(MapNum, I).y = MapNpc(MapNum, MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next I
            Else
                CanNpcMove = False
            End If
                
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If x < MAX_MAPX Then
                n = map(MapNum).Tile(x + 1, y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If (GetPlayerMap(I) = MapNum) And (GetPlayerX(I) = MapNpc(MapNum, MapNpcNum).x + 1) And (GetPlayerY(I) = MapNpc(MapNum, MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next I
                
                If CanNPCMoveAttributeNPC(MapNum, MapNpcNum, DIR_RIGHT) = False Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> MapNpcNum) And (MapNpc(MapNum, I).num > 0) And (MapNpc(MapNum, I).x = MapNpc(MapNum, MapNpcNum).x + 1) And (MapNpc(MapNum, I).y = MapNpc(MapNum, MapNpcNum).y) Then
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
Dim x As Long
Dim y As Long
Dim I As Long

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
    
    ReDim Player(index).Party.Member(1 To MAX_PARTY_MEMBERS)
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendEmoticons(index)
    Call SendElements(index)
    Call SendArrows(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendInventory(index)
    Call SendBank(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendStats(index)
    Call SendWeatherTo(index)
    Call SendTimeTo(index)
    Call SendGameClockTo(index)
    Call SendOnlineList
    Call SendGameClockTo(index)
    Call SendNewsTo(index)
    Call DisabledTimeTo(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    Call SendPlayerData(index)
    
    Call SendWornEquipment(index)
    Call SendIndexWornEquipmentFromMap(index)
    
    If Scripting = 1 Then
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
                
            MOTD = STR(MOTD)
            Call SendDataTo(index, MOTD)
    
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
        
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "LeftGame " & index
        Else
            ' Check for boot map
            If map(GetPlayerMap(index)).BootMap > 0 Then
                Call SetPlayerX(index, map(GetPlayerMap(index)).BootX)
                Call SetPlayerY(index, map(GetPlayerMap(index)).BootY)
                Call SetPlayerMap(index, map(GetPlayerMap(index)).BootMap)
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
Dim I As Long, n As Long

    n = 0
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
            n = n + 1
        End If
    Next I
    
    GetTotalMapPlayers = n
End Function

Function GetNpcMaxhp(ByVal NpcNum As Long)

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxhp = 0
        Exit Function
    End If
    
    GetNpcMaxhp = Npc(NpcNum).MaxHp
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
Dim I As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
            GetPlayerHPRegen = 0
            Exit Function
        End If
        
        I = Int(GetPlayerDEF(index) / 2)
        If I < 2 Then I = 2
        
        GetPlayerHPRegen = I
    End If
End Function

Function GetPlayerMPRegen(ByVal index As Long)
Dim I As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
            GetPlayerMPRegen = 0
            Exit Function
        End If
        
        I = Int(GetPlayerMAGI(index) / 2)
        If I < 2 Then I = 2
        
        GetPlayerMPRegen = I
    End If
End Function

Function GetPlayerSPRegen(ByVal index As Long)
Dim I As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
            GetPlayerSPRegen = 0
            Exit Function
        End If
        
        I = Int(GetPlayerSPEED(index) / 2)
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
Sub CheckPlayerLevelUp(ByVal index As Long)
Dim I As Long
Dim d As Long
Dim C As Long
    C = 0
    
    If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
        If GetPlayerLevel(index) < MAX_LEVEL Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerLevelUp " & index
            Else
                Do Until GetPlayerExp(index) < GetPlayerNextLevel(index)
                    DoEvents
                    If GetPlayerLevel(index) < MAX_LEVEL Then
                        If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
                            d = GetPlayerExp(index) - GetPlayerNextLevel(index)
                            Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
                            I = Int(GetPlayerSPEED(index) / 10)
                            If I < 1 Then I = 1
                            If I > 3 Then I = 3
                                
                            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + I)
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
Dim SpellNum As Long, I As Long, n As Long, Damage As Long
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
    
    I = GetSpellReqLevel(index, SpellNum)

    ' Check if they have enough MP
    If GetPlayerMP(index) < Spell(SpellNum).MPCost Then
        Call BattleMsg(index, "Not enough mana!", BrightRed, 0)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If I > GetPlayerLevel(index) Then
        Call BattleMsg(index, "You must be level " & I & " to cast this spell.", BrightRed, 0)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < Player(index).AttackTimer + 1000 Then
        Exit Sub
    End If
    
    ' Check if the spell is scripted and do that instead of a stat modification
    If Spell(SpellNum).Type = SPELL_TYPE_SCRIPTED Then

        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedSpell " & index & "," & Spell(SpellNum).Data1
        
        Exit Sub
    End If
  '  End If
        
Dim x As Long, y As Long

If Spell(SpellNum).AE = 1 Then
    For y = GetPlayerY(index) - Spell(SpellNum).Range To GetPlayerY(index) + Spell(SpellNum).Range
        For x = GetPlayerX(index) - Spell(SpellNum).Range To GetPlayerX(index) + Spell(SpellNum).Range
            n = -1
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) = True Then
                    If GetPlayerMap(index) = GetPlayerMap(I) Then
                        If GetPlayerX(I) = x And GetPlayerY(I) = y Then
                            If I = index Then
                                If Spell(SpellNum).Type = SPELL_TYPE_ADDHP Or Spell(SpellNum).Type = SPELL_TYPE_ADDMP Or Spell(SpellNum).Type = SPELL_TYPE_ADDSP Then
                                    Player(index).Target = I
                                    Player(index).TargetType = TARGET_TYPE_PLAYER
                                    n = Player(index).Target
                                End If
                            Else
                                Player(index).Target = I
                                Player(index).TargetType = TARGET_TYPE_PLAYER
                                n = Player(index).Target
                            End If
                        End If
                    End If
                End If
            Next I
            
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(GetPlayerMap(index), I).num > 0 Then
                    If Npc(MapNpc(GetPlayerMap(index), I).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(index), I).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        If MapNpc(GetPlayerMap(index), I).x = x And MapNpc(GetPlayerMap(index), I).y = y Then
                            Player(index).Target = I
                            Player(index).TargetType = TARGET_TYPE_NPC
                            n = Player(index).Target
                        End If
                    End If
                End If
            Next I
                
        Casted = False
        If n > 0 Then
            If Player(index).TargetType = TARGET_TYPE_PLAYER Then
                If IsPlaying(n) Then
                    If n <> index Then
                        Player(index).TargetType = TARGET_TYPE_PLAYER
                        If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And (map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
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
                        If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And (map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
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
                                If Spell(SpellNum).Element <> 0 And Npc(MapNpc(GetPlayerMap(index), n).num).Element <> 0 Then
                                If Element(Spell(SpellNum).Element).Strong = Npc(MapNpc(GetPlayerMap(index), n).num).Element Or Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Weak = Spell(SpellNum).Element Then
                                    Call BattleMsg(index, " A deadly mix of elements harm the " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & "!", Blue, 0)
                                    Damage = Int(Damage * 1.25)
                                If Element(Spell(SpellNum).Element).Strong = Npc(MapNpc(GetPlayerMap(index), n).num).Element And Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Weak = Spell(SpellNum).Element Then Damage = Int(Damage * 1.2)
                                End If
                                
                                If Element(Spell(SpellNum).Element).Weak = Npc(MapNpc(GetPlayerMap(index), n).num).Element Or Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Strong = Spell(SpellNum).Element Then
                                    Call BattleMsg(index, " The " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & " aborbs much of the elemental damage!", Red, 0)
                                    Damage = Int(Damage * 0.75)
                                If Element(Spell(SpellNum).Element).Weak = Npc(MapNpc(GetPlayerMap(index), n).num).Element And Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Strong = Spell(SpellNum).Element Then Damage = Int(Damage * (2 / 3))
                                End If
                                End If
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
        Call SendDataToMap(GetPlayerMap(index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & Player(index).Target & SEP_CHAR & Player(index).CastedSpell & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & END_CHAR)
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
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
            If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And (map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
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
                                                    If Spell(SpellNum).Element <> 0 And Npc(MapNpc(GetPlayerMap(index), n).num).Element <> 0 Then
                                If Element(Spell(SpellNum).Element).Strong = Npc(MapNpc(GetPlayerMap(index), n).num).Element Or Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Weak = Spell(SpellNum).Element Then
                                    Call BattleMsg(index, " A deadly mix of elements harm the " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & "!", Blue, 0)
                                    Damage = Int(Damage * 1.25)
                                If Element(Spell(SpellNum).Element).Strong = Npc(MapNpc(GetPlayerMap(index), n).num).Element And Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Weak = Spell(SpellNum).Element Then Damage = Int(Damage * 1.2)
                                End If
                                
                                If Element(Spell(SpellNum).Element).Weak = Npc(MapNpc(GetPlayerMap(index), n).num).Element Or Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Strong = Spell(SpellNum).Element Then
                                    Call BattleMsg(index, " The " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & " aborbs much of the elemental damage!", Red, 0)
                                    Damage = Int(Damage * 0.75)
                                If Element(Spell(SpellNum).Element).Weak = Npc(MapNpc(GetPlayerMap(index), n).num).Element And Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Strong = Spell(SpellNum).Element Then Damage = Int(Damage * (2 / 3))
                                End If
                                End If
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
        Call SendDataToMap(GetPlayerMap(index), "spellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & Player(index).Target & SEP_CHAR & Player(index).CastedSpell & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & END_CHAR)
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
    End If
End Sub
Function GetSpellReqLevel(ByVal index As Long, ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).LevelReq ' - Int(GetClassMAGI(GetPlayerClass(index)) / 4)
End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
Dim I As Long, n As Long

    CanPlayerCriticalHit = False
    
    If GetPlayerWeaponSlot(index) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            I = Int(GetPlayerSTR(index) / 2) + Int(GetPlayerLevel(index) / 2)
    
            n = Int(Rnd * 100) + 1
            If n <= I Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
Dim I As Long, n As Long, ShieldSlot As Long

    CanPlayerBlockHit = False
    
    ShieldSlot = GetPlayerShieldSlot(index)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            I = Int(GetPlayerDEF(index) / 2) + Int(GetPlayerLevel(index) / 2)
        
            n = Int(Rnd * 100) + 1
            If n <= I Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Sub CheckEquippedItems(ByVal index As Long)
Dim slot As Long, ItemNum As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    slot = GetPlayerWeaponSlot(index)
    If slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
                Call SetPlayerWeaponSlot(index, 0)
            End If
        Else
            Call SetPlayerWeaponSlot(index, 0)
        End If
    End If

    slot = GetPlayerArmorSlot(index)
    If slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then
                Call SetPlayerArmorSlot(index, 0)
            End If
        Else
            Call SetPlayerArmorSlot(index, 0)
        End If
    End If

    slot = GetPlayerHelmetSlot(index)
    If slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then
                Call SetPlayerHelmetSlot(index, 0)
            End If
        Else
            Call SetPlayerHelmetSlot(index, 0)
        End If
    End If

    slot = GetPlayerShieldSlot(index)
    If slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then
                Call SetPlayerShieldSlot(index, 0)
            End If
        Else
            Call SetPlayerShieldSlot(index, 0)
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

Public Sub RemovePMember(ByVal index As Byte)
Dim I, b, q As Integer
I = 0

b = Player(index).Party.Leader

If Player(index).Party.Leader = index Then 'Change the party leader!
    For I = 1 To MAX_PARTY_MEMBERS
        If Player(index).Party.Member(I) > 0 And Player(index).Party.Member(I) <> index Then
        Call ChangePLeader(Player(index).Party.Member(I))
        Exit For
        End If
    Next I
End If

I = 0

For q = 1 To MAX_PARTY_MEMBERS ' find which member the player is
    If Player(index).Party.Member(q) = index Then
        Exit For
    End If
Next q

For I = 1 To MAX_PARTY_MEMBERS ' removes player from other members party
    If Player(index).Party.Member(I) > 0 Then Player(Player(index).Party.Member(I)).Party.Member(q) = 0
Next I

Player(index).Party.Leader = 0 'no leader
Player(index).InvitedBy = 0

For I = 1 To MAX_PARTY_MEMBERS ' clears player's party
    Player(index).Party.Member(I) = 0
Next I

Player(index).InParty = False 'not in party

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

Public Sub ChangePLeader(ByVal index As Byte)
Dim I As Integer

Player(index).Party.Leader = index

For I = 1 To MAX_PARTY_MEMBERS
    If Player(index).Party.Member(I) > 0 Then Player(Player(index).Party.Member(I)).Party.Leader = index
Next I

For I = 1 To MAX_PARTY_MEMBERS
    If Player(index).Party.Member(I) > 0 Then Call PlayerMsg(Player(index).Party.Member(I), "Leadership has been passed to " & GetPlayerName(index) & "!", Pink)
Next I

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
On Error Resume Next

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
Function CanAttackPlayerWithArrow(ByVal Attacker As Long, ByVal Victim As Long) As Boolean

CanAttackPlayerWithArrow = False

' Check To make sure that they dont have access
If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
    Call PlayerMsg(Attacker, "You cannot attack Any player For thou art an admin!", BrightBlue)
Else
' Check To make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITER Then
    Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
Else
' Check If map Is attackable
If map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
' Make sure they are high enough level
If GetPlayerLevel(Attacker) < 10 Then
    Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
Else
If GetPlayerLevel(Victim) < 10 Then
    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " Is below level 10, you cannot attack this player yet!", BrightRed)
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
End If
End If

End Function

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
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
 If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpcWithArrow = True
Else
If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(NpcNum).SpawnSecs
Else
Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
End If
End If

            Case DIR_DOWN
 If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpcWithArrow = True
Else
If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(NpcNum).SpawnSecs
Else
Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
End If
End If

            Case DIR_LEFT
 If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpcWithArrow = True
Else
If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(NpcNum).SpawnSecs
Else
Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
End If
End If

            Case DIR_RIGHT
 If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpcWithArrow = True
Else
If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(NpcNum).SpawnSecs
Else
Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
End If
End If
        End Select
    End If
End If
End Function

Sub SendIndexWornEquipment(ByVal index As Long)
Dim Packet As String
Dim Armor As Long
Dim Helmet As Long
Dim Shield As Long
Dim Weapon As Long
Dim Legs As Long
Dim Ring As Long
Dim Necklace As Long

    Armor = 0
    Helmet = 0
    Shield = 0
    Weapon = 0
    Legs = 0
    Ring = 0
    Necklace = 0

    If GetPlayerArmorSlot(index) > 0 Then Armor = GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))
    If GetPlayerHelmetSlot(index) > 0 Then Helmet = GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))
    If GetPlayerShieldSlot(index) > 0 Then Shield = GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))
    If GetPlayerWeaponSlot(index) > 0 Then Weapon = GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))
    If GetPlayerLegsSlot(index) > 0 Then Legs = GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))
    If GetPlayerRingSlot(index) > 0 Then Ring = GetPlayerInvItemNum(index, GetPlayerRingSlot(index))
    If GetPlayerNecklaceSlot(index) > 0 Then Necklace = GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))
   
    Packet = "itemworn" & SEP_CHAR & index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendIndexWornEquipmentFromMap(ByVal index As Long)
Dim Packet As String
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
            If GetPlayerMap(I) = GetPlayerMap(index) Then
                
                Armor = 0
                Helmet = 0
                Shield = 0
                Weapon = 0
                Legs = 0
                Ring = 0
                Necklace = 0
           
                If GetPlayerArmorSlot(I) > 0 Then Armor = GetPlayerInvItemNum(I, GetPlayerArmorSlot(I))
                If GetPlayerHelmetSlot(I) > 0 Then Helmet = GetPlayerInvItemNum(I, GetPlayerHelmetSlot(I))
                If GetPlayerShieldSlot(I) > 0 Then Shield = GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))
                If GetPlayerWeaponSlot(I) > 0 Then Weapon = GetPlayerInvItemNum(I, GetPlayerWeaponSlot(I))
                If GetPlayerLegsSlot(I) > 0 Then Legs = GetPlayerInvItemNum(I, GetPlayerLegsSlot(I))
                If GetPlayerRingSlot(I) > 0 Then Ring = GetPlayerInvItemNum(I, GetPlayerRingSlot(I))
                If GetPlayerNecklaceSlot(I) > 0 Then Necklace = GetPlayerInvItemNum(I, GetPlayerNecklaceSlot(I))
               
                Packet = "itemworn" & SEP_CHAR & I & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & SEP_CHAR & END_CHAR
                Call SendDataTo(index, Packet)
                'Call SendDataTo(I, Packet)
            End If
        End If
    Next I
End Sub

'ASGARD
Public Sub LoadWordfilter()
    Dim I
    ReDim Wordfilter(Val(GetVar(App.Path & "\wordfilter.ini", "WORDFILTER", "maxwords")))
    If FileExist("wordfilter.ini") Then
        WordList = Val(GetVar(App.Path & "\wordfilter.ini", "WORDFILTER", "maxwords"))
        If WordList >= 1 Then
            For I = 1 To WordList
                Wordfilter(I) = LCase(GetVar(App.Path & "\wordfilter.ini", "WORDFILTER", "word" & I))
            Next I
        End If
    Else
        Call MsgBox("Wordfilter.INI could not be found. Please make sure it exists.")
        WordList = 0
    End If
End Sub

Public Function SwearCheck(TextToSay As String) As String
    Dim I
    Dim J
    Dim Asterisk As String
    SwearCheck = TextToSay
    If WordList <= 0 Then Exit Function
    For I = 1 To WordList
    Asterisk = ""
    For J = 1 To Len(Wordfilter(I))
        Asterisk = Asterisk & "*"
    Next J
        SwearCheck = Replace(SwearCheck, Wordfilter(I), Asterisk, 1, -1, vbTextCompare)
    Next I
End Function
Sub AddNewTimer(ByVal Name As String, ByVal Interval As Long)
On Error Resume Next
  Dim TmpTimer As clsCTimers
  Set TmpTimer = New clsCTimers
  TmpTimer.Name = Name
  TmpTimer.Interval = Interval
  TmpTimer.tmrWait = GetTickCount + Interval
  CTimers.add TmpTimer, Name
  If Err.number > 0 Then
    Debug.Print "Err: " & Err.number
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
Sub ScriptSetTile(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal setx As Long, ByVal sety As Long, ByVal tileset As Long, ByVal layer As Long)
Dim Packet As String
Packet = "tilecheck" & SEP_CHAR & mapper & SEP_CHAR & CStr(x) & SEP_CHAR & CStr(y) & SEP_CHAR & CStr(layer) & SEP_CHAR

Select Case layer
Case 0
    map(mapper).Tile(x, y).Ground = sety * 14 + setx
    map(mapper).Tile(x, y).GroundSet = tileset
Packet = Packet & map(mapper).Tile(x, y).Ground & SEP_CHAR & map(mapper).Tile(x, y).GroundSet
Case 1
    map(mapper).Tile(x, y).Mask = sety * 14 + setx
    map(mapper).Tile(x, y).MaskSet = tileset
Packet = Packet & map(mapper).Tile(x, y).Mask & SEP_CHAR & map(mapper).Tile(x, y).MaskSet
Case 2
    map(mapper).Tile(x, y).Anim = sety * 14 + setx
    map(mapper).Tile(x, y).AnimSet = tileset
Packet = Packet & map(mapper).Tile(x, y).Anim & SEP_CHAR & map(mapper).Tile(x, y).AnimSet
Case 3
    map(mapper).Tile(x, y).Mask2 = sety * 14 + setx
    map(mapper).Tile(x, y).Mask2Set = tileset
Packet = Packet & map(mapper).Tile(x, y).Mask2 & SEP_CHAR & map(mapper).Tile(x, y).Mask2Set
Case 4
    map(mapper).Tile(x, y).M2Anim = sety * 14 + setx
    map(mapper).Tile(x, y).M2AnimSet = tileset
Packet = Packet & map(mapper).Tile(x, y).M2Anim & SEP_CHAR & map(mapper).Tile(x, y).M2AnimSet
Case 5
    map(mapper).Tile(x, y).Fringe = sety * 14 + setx
    map(mapper).Tile(x, y).FringeSet = tileset
Packet = Packet & map(mapper).Tile(x, y).Fringe & SEP_CHAR & map(mapper).Tile(x, y).FringeSet
Case 6
    map(mapper).Tile(x, y).FAnim = sety * 14 + setx
    map(mapper).Tile(x, y).FAnimSet = tileset
Packet = Packet & map(mapper).Tile(x, y).FAnim & SEP_CHAR & map(mapper).Tile(x, y).FAnimSet
Case 7
    map(mapper).Tile(x, y).Fringe2 = sety * 14 + setx
    map(mapper).Tile(x, y).Fringe2Set = tileset
Packet = Packet & map(mapper).Tile(x, y).Fringe2 & SEP_CHAR & map(mapper).Tile(x, y).Fringe2Set
Case 8
    map(mapper).Tile(x, y).F2Anim = sety * 14 + setx
    map(mapper).Tile(x, y).F2AnimSet = tileset
Packet = Packet & map(mapper).Tile(x, y).F2Anim & SEP_CHAR & map(mapper).Tile(x, y).F2AnimSet
End Select
Call SaveMap(mapper)
Call SendDataToAll(Packet & END_CHAR)
End Sub


Sub ScriptSetAttribute(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
Dim Packet As String
With map(mapper).Tile(x, y)
    .Type = Attrib
    .Data1 = Data1
    .Data2 = Data2
    .Data3 = Data3
    .String1 = String1
    .String2 = String2
    .String3 = String3
End With

Packet = "tilecheckattribute" & SEP_CHAR & mapper & SEP_CHAR & CStr(x) & SEP_CHAR & CStr(y) & SEP_CHAR
        With map(mapper).Tile(x, y)
            Packet = Packet & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR
        End With
Call SendDataToAll(Packet & END_CHAR)
End Sub
