Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal index As Long) As Long
Dim GetBaseDamage As Long
Dim GetMaxDamage As Long
On Error GoTo ErrorHandler
    GetPlayerDamage = 0
    GetBaseDamage = 0
    GetMaxDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    GetBaseDamage = Int(GetPlayerSTR(index) / 1.5)
    GetMaxDamage = Int(GetPlayerSTR(index) * 1.25)
    GetPlayerDamage = Int((GetMaxDamage - GetBaseDamage + 1) * Rnd + GetBaseDamage)
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetPlayerDamage", Err.Number, Err.Description
End Function

Function GetPlayerProtection(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
       GetPlayerProtection = GetPlayerDEF(index)

ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetPlayerProtection", Err.Number, Err.Description
End Function

Function GetNPCPlayerProtection(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetNPCPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    GetNPCPlayerProtection = GetPlayerDEF(index)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetNPCPlayerProtection", Err.Number, Err.Description
End Function

Function FindOpenPlayerSlot() As Long
On Error GoTo ErrorHandler
Dim i As Long

    FindOpenPlayerSlot = 0
    
    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindOpenPlayerSlot", Err.Number, Err.Description
End Function

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
On Error GoTo ErrorHandler
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
                If Trim(LCase(Item(GetPlayerInvItemNum(index, i)).Name)) = "gold" Then
                    FindOpenInvSlot = i
                Else
                    If GetPlayerInvItemValue(index, i) >= 100 Then
                        FindOpenInvSlot = 0
                    Else
                        FindOpenInvSlot = i
                    End If
                End If
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
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindOpenInvSlot", Err.Number, Err.Description
End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
On Error GoTo ErrorHandler
Dim i As Long
    
    FindOpenBankSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i
    End If
    
    For i = 1 To MAX_BANK
        ' Try to find an open free slot
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindOpenBankSlot", Err.Number, Err.Description
End Function


Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindOpenMapItemSlot", Err.Number, Err.Description
End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindOpenSpellSlot", Err.Number, Err.Description
End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
On Error GoTo ErrorHandler
Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "HasSpell", Err.Number, Err.Description
End Function
Function TotalOnlinePlayers() As Long
On Error GoTo ErrorHandler
Dim i As Long
    frmServer.LstPlayers.Clear
    frmServer.LstAccounts.Clear
    TotalOnlinePlayers = 0

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
            frmServer.LstPlayers.AddItem Trim(Player(i).Char(Player(i).CharNum).Name)
            frmServer.LstAccounts.AddItem Trim(Player(i).Login)
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "TotalOnlinePlayers", Err.Number, Err.Description
End Function

Function FindPlayer(ByVal Name As String) As Long
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindPlayer", Err.Number, Err.Description
End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "HasItem", Err.Number, Err.Description
End Function

Sub TakeItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
On Error GoTo ErrorHandler
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
                        
                        Case ITEM_TYPE_BOOTS
                        If GetPlayerBootsSlot(index) > 0 Then
                            If i = GetPlayerBootsSlot(index) Then
                                Call SetPlayerBootsSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerBootsSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                        Case ITEM_TYPE_GLOVES
                        If GetPlayerGlovesSlot(index) > 0 Then
                            If i = GetPlayerGlovesSlot(index) Then
                                Call SetPlayerGlovesSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerGlovesSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                        Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(index) > 0 Then
                            If i = GetPlayerRingSlot(index) Then
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
                        
                        Case ITEM_TYPE_AMULET
                        If GetPlayerAmuletSlot(index) > 0 Then
                            If i = GetPlayerAmuletSlot(index) Then
                                Call SetPlayerAmuletSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerAmuletSlot(index)) Then
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "TakeItem", Err.Number, Err.Description
End Sub

Sub GiveItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
On Error GoTo ErrorHandler
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
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_BOOTS) Or (Item(ItemNum).Type = ITEM_TYPE_GLOVES) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_AMULET) Then
            Call SetPlayerInvItemDur(index, i, Item(ItemNum).Data1)
        End If
        
        'Call BattleMsg(index, "You picked up", Yellow, index)
        Call SendInventoryUpdate(index, i)
    Else
        Call BattleMsg(index, "Your inventory is full.", Yellow, index)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "GiveItem", Err.Number, Err.Description
End Sub

Sub TakeBankItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
On Error GoTo ErrorHandler
Dim i As Long, n As Long
Dim TakeBankItem As Boolean

    TakeBankItem = False
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_BANK
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerBankItemValue(index, i) Then
                    TakeBankItem = True
                Else
                    Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) - ItemVal)
                    Call SendBankUpdate(index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerBankItemNum(index, i)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(index) > 0 Then
                            If i = GetPlayerWeaponSlot(index) Then
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
                            If i = GetPlayerArmorSlot(index) Then
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
                            If i = GetPlayerHelmetSlot(index) Then
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
                            If i = GetPlayerShieldSlot(index) Then
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
                        
                        Case ITEM_TYPE_BOOTS
                        If GetPlayerBootsSlot(index) > 0 Then
                            If i = GetPlayerBootsSlot(index) Then
                                Call SetPlayerBootsSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerBootsSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                        
                        Case ITEM_TYPE_GLOVES
                        If GetPlayerGlovesSlot(index) > 0 Then
                            If i = GetPlayerGlovesSlot(index) Then
                                Call SetPlayerGlovesSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerGlovesSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                        
                        Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(index) > 0 Then
                            If i = GetPlayerRingSlot(index) Then
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
                        
                        Case ITEM_TYPE_AMULET
                        If GetPlayerAmuletSlot(index) > 0 Then
                            If i = GetPlayerAmuletSlot(index) Then
                                Call SetPlayerAmuletSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerAmuletSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                End Select

                
                n = Item(GetPlayerBankItemNum(index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_BOOTS) And (n <> ITEM_TYPE_GLOVES) And (n <> ITEM_TYPE_RING) And (n <> ITEM_TYPE_AMULET) Then
                    TakeBankItem = True
                End If
            End If
                            
            If TakeBankItem = True Then
                Call SetPlayerBankItemNum(index, i, 0)
                Call SetPlayerBankItemValue(index, i, 0)
                Call SetPlayerBankItemDur(index, i, 0)
                
                ' Send the Bank update
                Call SendBankUpdate(index, i)
                Exit Sub
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "TakeBankItem", Err.Number, Err.Description
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal BankSlot As Long)
On Error GoTo ErrorHandler
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = BankSlot
    
    ' Check to see if Bankentory is full
    If i <> 0 Then
        Call SetPlayerBankItemNum(index, i, ItemNum)
        Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_BOOTS) Or (Item(ItemNum).Type = ITEM_TYPE_GLOVES) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_AMULET) Then
            Call SetPlayerBankItemDur(index, i, Item(ItemNum).Data1)
        End If
    Else
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Bank full!" & SEP_CHAR & END_CHAR)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "GiveBankItem", Err.Number, Err.Description
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, X, Y)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnItem", Err.Number, Err.Description
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler
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
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type <= ITEM_TYPE_BOOTS) And (Item(ItemNum).Type <= ITEM_TYPE_GLOVES) And (Item(ItemNum).Type <= ITEM_TYPE_RING) And (Item(ItemNum).Type <= ITEM_TYPE_AMULET) Then
                MapItem(MapNum, i).Dur = ItemDur
            Else
                MapItem(MapNum, i).Dur = 0
            End If
        Else
            MapItem(MapNum, i).Dur = 0
        End If
        
        MapItem(MapNum, i).X = X
        MapItem(MapNum, i).Y = Y
            
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnItemSlot", Err.Number, Err.Description
End Sub

Sub SpawnAllMapsItems()
On Error GoTo ErrorHandler
Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnAllMapsItems", Err.Number, Err.Description
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
On Error GoTo ErrorHandler
Dim X As Long
Dim Y As Long
Dim i As Long

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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnMapItems", Err.Number, Err.Description
End Sub

Sub PlayerMapGetItem(ByVal index As Long)
On Error GoTo ErrorHandler
Dim i As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String
Dim TempValue As Long, TempNum As Long

    If IsPlaying(index) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(index)
    
    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).X = GetPlayerX(index)) And (MapItem(MapNum, i).Y = GetPlayerY(index)) Then
                ' Find open slot
                n = FindOpenInvSlot(index, MapItem(MapNum, i).Num)
                
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(index, n, MapItem(MapNum, i).Num)
                    If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
                        If Trim(LCase(Item(GetPlayerInvItemNum(index, n)).Name)) = "gold" Then
                            TempValue = 0
                            TempNum = 0
                            Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).Value)
                        Else
                            If GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).Value > 100 Then
                                TempValue = 100 - GetPlayerInvItemValue(index, n)
                                TempNum = MapItem(MapNum, i).Num
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + TempValue)
                            Else
                                TempValue = 0
                                TempNum = 0
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).Value)
                            End If
                        End If
                        Msg = "You picked up " & MapItem(MapNum, i).Value & " " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(index, n, 0)
                        Msg = "You picked up a " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                    End If
                    Call SetPlayerInvItemDur(index, n, MapItem(MapNum, i).Dur)
                        
                    ' Erase item from the map
                    MapItem(MapNum, i).Num = TempNum
                    MapItem(MapNum, i).Value = TempValue
                    MapItem(MapNum, i).Dur = 0
                    MapItem(MapNum, i).X = 0
                    MapItem(MapNum, i).Y = 0
                        
                    Call SendInventoryUpdate(index, n)
                    Call SpawnItemSlot(i, TempNum, TempValue, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    Call BattleMsg(index, Msg, Yellow, index)
                    Exit Sub
                Else
                    Call BattleMsg(index, "Your inventory is full.", Yellow, index)
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    Call SendInventory(index)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "PlayerMapGetItem", Err.Number, Err.Description
End Sub

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
On Error GoTo ErrorHandler
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
                    
                Case ITEM_TYPE_BOOTS
                    If InvNum = GetPlayerBootsSlot(index) Then
                        Call SetPlayerBootsSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                
                Case ITEM_TYPE_GLOVES
                    If InvNum = GetPlayerGlovesSlot(index) Then
                        Call SetPlayerGlovesSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_RING
                    If InvNum = GetPlayerRingSlot(index) Then
                        Call SetPlayerRingSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_AMULET
                    If InvNum = GetPlayerAmuletSlot(index) Then
                        Call SetPlayerAmuletSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), i).X = GetPlayerX(index)
            MapItem(GetPlayerMap(index), i).Y = GetPlayerY(index)
                        
            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, InvNum)
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), i).Value = Amount
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Amount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(index), i).Value = 0
                If Item(GetPlayerInvItemNum(index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_SHIELD Or Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_BOOTS And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_GLOVES And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_RING And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_AMULET Then
                    If Item(GetPlayerInvItemNum(index, InvNum)).Data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Else
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    End If
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
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, Amount, MapItem(GetPlayerMap(index), i).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsg(index, "To many items already on the ground.", BrightRed)
        End If
    End If
    
    Call SendInventory(index)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "PlayerMapDropItem", Err.Number, Err.Description
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim NpcNum As Long
Dim i As Long, X As Long, Y As Long
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
        
        ' Well try 100 times to randomly place the sprite
        For i = 1 To 100
            X = Int(Rnd * MAX_MAPX)
            Y = Int(Rnd * MAX_MAPY)
            
            ' Check if the tile is walkable
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                MapNpc(MapNum, MapNpcNum).X = X
                MapNpc(MapNum, MapNpcNum).Y = Y
                Spawned = True
                Exit For
            End If
        Next i
        
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
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).Num).Big & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        End If
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnNpc", Err.Number, Err.Description
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnMapNpcs", Err.Number, Err.Description
End Sub

Sub SpawnAllMapNpcs()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnAllMapNpcs", Err.Number, Err.Description
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
On Error GoTo ErrorHandler
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
    If (GetPlayerMap(Attacker) = GetPlayerMap(Victim)) And (GetTickCount > Player(Attacker).AttackTimer + 1000) Then
        
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < 5 Then
                                Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
                            Else
                                If GetPlayerLevel(Victim) < 5 Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Attacker) - 10 <= GetPlayerLevel(Victim) Then
                                        If GetPlayerLevel(Attacker) + 11 <= GetPlayerLevel(Victim) Then
                                            Call PlayerMsg(Attacker, "You need to be within 10 levels of their level to attack them!", BrightRed)
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
                                    Else
                                        Call PlayerMsg(Attacker, "You need to be within 10 levels of their level to attack them!", BrightRed)
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
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < 5 Then
                                Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
                            Else
                                If GetPlayerLevel(Victim) < 5 Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Attacker) - 10 <= GetPlayerLevel(Victim) Then
                                        If GetPlayerLevel(Attacker) + 11 <= GetPlayerLevel(Victim) Then
                                            Call PlayerMsg(Attacker, "You need to be within 10 levels of their level to attack them!", BrightRed)
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
                                    Else
                                        Call PlayerMsg(Attacker, "You need to be within 10 levels of their level to attack them!", BrightRed)
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
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < 5 Then
                                Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
                            Else
                                If GetPlayerLevel(Victim) < 5 Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Attacker) - 10 <= GetPlayerLevel(Victim) Then
                                        If GetPlayerLevel(Attacker) + 11 <= GetPlayerLevel(Victim) Then
                                            Call PlayerMsg(Attacker, "You need to be within 10 levels of their level to attack them!", BrightRed)
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
                                    Else
                                        Call PlayerMsg(Attacker, "You need to be within 10 levels of their level to attack them!", BrightRed)
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
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < 5 Then
                                Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
                            Else
                                If GetPlayerLevel(Victim) < 5 Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Attacker) - 10 <= GetPlayerLevel(Victim) Then
                                        If GetPlayerLevel(Attacker) + 11 <= GetPlayerLevel(Victim) Then
                                            Call PlayerMsg(Attacker, "You need to be within 10 levels of their level to attack them!", BrightRed)
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
                                    Else
                                        Call PlayerMsg(Attacker, "You need to be within 10 levels of their level to attack them!", BrightRed)
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

ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "CanAttackPlayer", Err.Number, Err.Description
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
On Error GoTo ErrorHandler
Dim MapNum As Long, NpcNum As Long
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
NpcNum = MapNpc(MapNum, MapNpcNum).Num
 
' Make sure the npc isn't already dead
If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
Exit Function
End If
 
' Make sure they are on the same map
If IsPlaying(Attacker) Then
    If NpcNum > 0 And GetTickCount > Player(Attacker).AttackTimer + 1000 Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (MapNpc(MapNum, MapNpcNum).Y + 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
 
            Case DIR_DOWN
                If (MapNpc(MapNum, MapNpcNum).Y - 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
 
            Case DIR_LEFT
                If (MapNpc(MapNum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X + 1 = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
 
            Case DIR_RIGHT
                If (MapNpc(MapNum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).X - 1 = GetPlayerX(Attacker)) Then
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        CanAttackNpc = True
                    Else
                        Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If
                End If
        End Select
    End If
End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "CanAttackNpc", Err.Number, Err.Description
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
On Error GoTo ErrorHandler
Dim MapNum As Long, NpcNum As Long
    
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
    If Player(index).GettingMap = YES Then
        Exit Function
    End If
    
    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount
    
    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NpcNum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) + 1 = MapNpc(MapNum, MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) - 1 = MapNpc(MapNum, MapNpcNum).X) Then
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
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "CanNpcAttackPlayer", Err.Number, Err.Description
End Function

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
On Error GoTo ErrorHandler
Dim Exp As Long
Dim n As Long
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then Exit Sub
    
    If GetTickCount < Player(Attacker).AttackTimer + 1000 Then Exit Sub
    
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
        Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, Attacker)
        Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, Attacker)
    
        ' Player is dead
        ' Player is dead
        If Trim(GetPlayerGuild(Attacker)) <> "" Then
            Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker) & " {" & Trim(GetPlayerGuild(Attacker)) & "}", BrightRed)
        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        End If
        
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
                
                If GetPlayerBootsSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerBootsSlot(Victim), 0)
                End If
                
                If GetPlayerGlovesSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerGlovesSlot(Victim), 0)
                End If
                
                If GetPlayerRingSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerRingSlot(Victim), 0)
                End If
                
                If GetPlayerAmuletSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerAmuletSlot(Victim), 0)
                End If
            End If
            
            ' Calculate exp to give attacker
            Exp = Int(GetPlayerExp(Victim) / 8)
            
            ' Make sure we dont get less then 0
            If Exp < 0 Then
                Exp = 0
            End If
            
            If GetPlayerLevel(Victim) = MAX_LEVEL Then
                Call BattleMsg(Victim, "You cant lose any experience!", BrightRed, Attacker)
                Call BattleMsg(Attacker, GetPlayerName(Victim) & " is the max level!", BrightBlue, Attacker)
            Else
                If Exp = 0 Then
                    Call BattleMsg(Victim, "You lost no experience.", BrightRed, Attacker)
                    Call BattleMsg(Attacker, "You received no experience.", BrightBlue, Attacker)
                Else
                    Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                    Call BattleMsg(Victim, "You lost " & Exp & " experience.", BrightRed, Attacker)
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                    Call BattleMsg(Attacker, "You got " & Exp & " experience for killing " & GetPlayerName(Victim) & ".", BrightBlue, Attacker)
                End If
            End If
        End If
                
        ' Warp player away
        'MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
        Call PlayerWarp(Victim, Player(Victim).Char(Player(Victim).CharNum).Binding.Map, Player(Victim).Char(Player(Victim).CharNum).Binding.X, Player(Victim).Char(Player(Victim).CharNum).Binding.Y)
        Call PlaySound(Victim, "magic22.wav")
        
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
        Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, Attacker)
        Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, Attacker)
    End If
ElseIf Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, Attacker)
        Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, Attacker)
    
        ' Player is dead
        If Trim(GetPlayerGuild(Attacker)) <> "" Then
            Call GlobalMsg(GetPlayerName(Victim) & " has been killed in the arena by " & GetPlayerName(Attacker) & " {" & Trim(GetPlayerGuild(Attacker)) & "}", BrightRed)
        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has been killed in the arena by " & GetPlayerName(Attacker), BrightRed)
        End If
        
        ' Warp player away
        Call PlayerWarp(Victim, Player(Victim).Char(Player(Victim).CharNum).Binding.Map, Player(Victim).Char(Player(Victim).CharNum).Binding.X, Player(Victim).Char(Player(Victim).CharNum).Binding.Y)
        Call PlaySound(Victim, "magic22.wav")
        
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
        Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, Attacker)
        Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, Attacker)
    End If
End If
    
    ' Reset timer for attacking
    Player(Attacker).AttackTimer = GetTickCount
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "AttackPlayer", Err.Number, Err.Description
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
On Error GoTo ErrorHandler
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
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Victim)
    
    Name = Trim(Npc(MapNpc(MapNum, MapNpcNum).Num).Name)
    
    If Damage >= GetPlayerHP(Victim) Then
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
                
                If GetPlayerBootsSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerBootsSlot(Victim), 0)
                End If
                
                If GetPlayerGlovesSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerGlovesSlot(Victim), 0)
                End If
                
                If GetPlayerRingSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerRingSlot(Victim), 0)
                End If
                
                If GetPlayerAmuletSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerAmuletSlot(Victim), 0)
                End If
            End If
            
            ' Calculate exp to give attacker
            Exp = Int(GetPlayerExp(Victim) / 13)
            
            ' Make sure we dont get less then 0
            If Exp < 0 Then
                Exp = 0
            End If
            
            If Exp = 0 Then
                Call BattleMsg(Victim, "You lost no experience.", BrightRed, Victim)
            Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call BattleMsg(Victim, "You lost " & Exp & " experience.", BrightRed, Victim)
            End If
        End If
                
        Call PlaySound(Victim, "magic22.wav")
        Call PlayerWarp(Victim, Player(Victim).Char(Player(Victim).CharNum).Binding.Map, Player(Victim).Char(Player(Victim).CharNum).Binding.X, Player(Victim).Char(Player(Victim).CharNum).Binding.Y)
        
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
        
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        If GetPlayerPK(Victim) = YES Then
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
        End If
    Else
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
    End If
    
    Call BattleMsg(Victim, "The " & Name & " hit you for " & Damage & "!", BrightRed, Victim + 1)
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "NpcAttackPlayer", Err.Number, Err.Description
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
On Error GoTo ErrorHandler
Dim Name As String
Dim Exp As Long
Dim n As Long, i As Long
Dim STR As Long, DEF As Long, MapNum As Long, NpcNum As Long

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
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    Name = Trim(Npc(NpcNum).Name)
        
    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "You killed a " & Name, BrightRed, Attacker)
        
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
         If GetPlayerBootsSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerBootsSlot(Attacker))).AddEXP
        End If
         If GetPlayerGlovesSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerGlovesSlot(Attacker))).AddEXP
        End If
         If GetPlayerRingSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerRingSlot(Attacker))).AddEXP
        End If
         If GetPlayerAmuletSlot(Attacker) > 0 Then
            add = add + Item(GetPlayerInvItemNum(Attacker, GetPlayerAmuletSlot(Attacker))).AddEXP
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
        
        If GetPlayerLevel(Attacker) > 10 Then
            Exp = Exp - GetPlayerLevel(Attacker)
        Else
            Exp = Exp - (GetPlayerLevel(Attacker) / 2)
        End If
        
        ' Make sure we dont get less then 0
        If Exp <= 0 Then
            Exp = 0
        End If

        ' Check if in party, if so divide the exp up by the number of people in the group
        If Player(Attacker).Party.InParty = NO Then
            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "You cant gain anymore experience!", BrightBlue, Attacker)
            Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "You gain " & Exp & " experience from the " & Name & ".", BrightBlue, Attacker)
            End If
        Else
            n = 1
            If Player(Attacker).Party.Started = YES Then
                For i = 1 To MAX_PARTY_MEMS
                    If Player(Attacker).Party.PlayerNums(i) > 0 Then
                        n = n + 1
                    End If
                Next i
            Else
                For i = 1 To MAX_PARTY_MEMS
                    If Player(Player(Attacker).Party.PlayerNums(1)).Party.PlayerNums(i) > 0 Then
                        n = n + 1
                    End If
                Next i
            End If
            
            If Exp > 0 Then
                Exp = Exp / (n * 0.75)
            End If
            
            If Exp < 0 Then
                Exp = 1
            End If
            
            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "You cant gain anymore experience!", BrightBlue, Attacker)
            Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "You have gained " & Exp & " party experience points.", BrightBlue, Attacker)
            End If
            
            For i = 1 To MAX_PARTY_MEMS
                n = Player(Attacker).Party.PlayerNums(i)
                If n > 0 Then
                    If GetPlayerLevel(n) = MAX_LEVEL Then
                        Call SetPlayerExp(n, Experience(MAX_LEVEL))
                        Call PlayerMsg(n, "You cant gain anymore experience!", BrightBlue)
                    Else
                        Call SetPlayerExp(n, GetPlayerExp(n) + Exp)
                        Call PlayerMsg(n, "You have gained " & Exp & " party experience points.", BrightBlue)
                    End If
                End If
            Next i
        End If
                                
        For i = 1 To MAX_NPC_DROPS
            If Npc(NpcNum).ItemNPC(i).ItemNum > 0 Then
                n = Int(Rnd * Npc(NpcNum).ItemNPC(i).Chance) + 1
                If n = 1 Then
                    
                    Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
                End If
            End If
        Next i
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).Num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check for level up party member(s)
        If Player(Attacker).Party.InParty = YES Then
            For i = 1 To MAX_PARTY_MEMS
                Call CheckPlayerLevelUp(Player(Attacker).Party.PlayerNums(i))
            Next i
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
        Call BattleMsg(Attacker, "You hit a " & Name & " for " & Damage & " damage.", White, Attacker)
        
        ' Check if we should send a message
        If MapNpc(MapNum, MapNpcNum).Target = 0 And MapNpc(MapNum, MapNpcNum).Target <> Attacker Then
            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " : " & Trim(Npc(NpcNum).AttackSay) & "", SayColor)
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
    
    Call SendNpcHP(MapNum, MapNpcNum)
    
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "AttackNpc", Err.Number, Err.Description
End Sub

Sub SendNpcHP(ByVal MapNum As Long, ByVal MapNpcNum As Long)
On Error GoTo ErrorHandler
Dim NpcNum As Long

    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SendNpcHP", Err.Number, Err.Description
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler
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
    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)
                
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Player(index).GettingMap = YES
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "warp" & SEP_CHAR & END_CHAR)
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "PlayerWarp", Err.Number, Err.Description
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim MapNum As Long
Dim X As Long
Dim Y As Long
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

    If GetPlayerX(index) + 1 < MAX_MAPX + 1 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(index) + 1
            Y = GetPlayerY(index)
            
            If TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerX(index) - 1 > -1 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(index) - 1
            Y = GetPlayerY(index)
            
            If TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(index) - 1 > -1 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(index)
            Y = GetPlayerY(index) - 1
            
            If TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(index) + 1 < MAX_MAPY + 1 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(index)
            Y = GetPlayerY(index) + 1
            
            If TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                                
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
            
    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        X = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        Y = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3
                        
        Call PlayerWarp(index, MapNum, X, Y)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KEYOPEN Then
        X = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        Y = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        
        If Map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
            TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
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
    
    ' They tried to hack
    'If Moved = NO Then
        'Call HackingAttempt(Index, "Position Modification")
    'End If
    
    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(index) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            Call PlayerMsg(index, "You already have this sprite!", BrightRed)
            Exit Sub
        Else
            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
                Call PlayerMsg(index, "This sprite is free!", Yellow)
            Else
                If Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(index, "This sprite will cost you " & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3 & " " & Trim(Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", Yellow)
                Else
                    Call PlayerMsg(index, "This sprite will cost you a " & Trim(Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", Yellow)
                End If
            End If
            Call SendDataTo(index, "spritechange" & SEP_CHAR & END_CHAR)
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
            
            'Call SetPlayerSTR(index, (Player(index).Char(Player(index).CharNum).STR - Class(GetPlayerClass(index)).STR))
            'Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).DEF - Class(GetPlayerClass(index)).DEF))
            'Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).MAGI - Class(GetPlayerClass(index)).MAGI))
            'Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).SPEED - Class(GetPlayerClass(index)).SPEED))
            
            Call SetPlayerClass(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)

            Call SetPlayerSTR(index, (GetPlayerBaseSTR(index) + Class(GetPlayerClass(index)).STR))
            Call SetPlayerDEF(index, (GetPlayerBaseDEF(index) + Class(GetPlayerClass(index)).DEF))
            Call SetPlayerMAGI(index, (GetPlayerBaseMAGI(index) + Class(GetPlayerClass(index)).MAGI))
            Call SetPlayerSPEED(index, (GetPlayerBaseSPEED(index) + Class(GetPlayerClass(index)).SPEED))
            Call SetPlayerVIT(index, (GetPlayerBaseVIT(index) + Class(GetPlayerClass(index)).VIT))
            
            Call PlayerMsg(index, "Your new class is a " & Trim(GetPlayerClassName(GetPlayerClass(index))) & "!", BrightGreen)
            
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
    
    If Scripting = 1 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & index & "," & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        End If
    End If
    
    'healing tiles code
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SendHP(index)
        Call PlayerMsg(index, "You feel completely refreshed!", BrightGreen)
    End If
    
    'Check for kill tile, and if so kill them
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KILL Then
        Call SetPlayerHP(index, 0)
        Call PlayerMsg(index, "The hand of death has awakened!", BrightRed)
        
        'MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & index
        Call PlayerWarp(index, Player(index).Char(Player(index).CharNum).Binding.Map, Player(index).Char(Player(index).CharNum).Binding.X, Player(index).Char(Player(index).CharNum).Binding.Y)
        Call PlaySound(index, "magic22.wav")
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SetPlayerSP(index, GetPlayerMaxSP(index))
        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
        Moved = YES
    End If
    
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_BANK Then
        Call SendDataTo(index, "openbank" & SEP_CHAR & END_CHAR)
    End If

ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "PlayerMove", Err.Number, Err.Description
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
On Error GoTo ErrorHandler
Dim i As Long, n As Long
Dim X As Long, Y As Long

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
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X) And (MapNpc(MapNum, i).Y = MapNpc(MapNum, MapNpcNum).Y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
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
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X) And (MapNpc(MapNum, i).Y = MapNpc(MapNum, MapNpcNum).Y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
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
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X - 1) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X - 1) And (MapNpc(MapNum, i).Y = MapNpc(MapNum, MapNpcNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
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
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum, MapNpcNum).X + 1) And (GetPlayerY(i) = MapNpc(MapNum, MapNpcNum).Y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next i
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum, i).Num > 0) And (MapNpc(MapNum, i).X = MapNpc(MapNum, MapNpcNum).X + 1) And (MapNpc(MapNum, i).Y = MapNpc(MapNum, MapNpcNum).Y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next i
            Else
                CanNpcMove = False
            End If
    End Select
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "CanNpcMove", Err.Number, Err.Description
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim X As Long
Dim Y As Long
Dim i As Long

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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "NpcMove", Err.Number, Err.Description
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    Packet = "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "NpcDir", Err.Number, Err.Description
End Sub

Sub JoinGame(ByVal index As Long)
On Error GoTo ErrorHandler
Dim MOTD As String
Dim f As Long

    ' Set the flag so we know the person is in the game
    Player(index).InGame = True
           
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(index, "LOGINOK" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendEmoticons(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendInventory(index)
    Call SendBank(index)
    Call SendWornEquipment(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendStats(index)
    Call SendWeatherTo(index)
    Call SendTimeTo(index)
    Call SendOnlineList
    Call PlayerClass(index)
    Call PlayerPoints(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    Call SendPlayerData(index)

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
        Call PlayerMsg(index, "Welcome to " & GAME_NAME & "!", 15)
        
        ' Send motd
        If Trim(MOTD) <> "" Then
            Call PlayerMsg(index, "MOTD: " & MOTD, 11)
        End If
    End If
    
    ' Send whos online
    Call SendWhosOnline(index)

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(index, "INGAME" & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "JoinGame", Err.Number, Err.Description
End Sub

Sub LeftGame(ByVal index As Long)
On Error GoTo ErrorHandler
Dim i As Long
Dim n As Long
Dim f As Byte

    If Player(index).InGame = True Then
        Player(index).InGame = False
        
        Call SavePlayer(index, Player(index).CharNum)
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) = 0 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If

        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If Player(index).Party.InParty = YES Then
            If Player(index).Party.Started = YES Then
                For i = 1 To MAX_PARTY_MEMS
                    n = Player(index).Party.PlayerNums(i)
                    If n > 0 Then
                        Call PlayerMsg(n, GetPlayerName(index) & " has left " & GAME_NAME & ". The party has been disbanded.", 13)
                        Player(n).Party.InParty = NO
                        Player(n).Party.PlayerNums(1) = 0
                    End If
                Next i
            Else
                n = Player(index).Party.PlayerNums(1)
                Call PlayerMsg(n, GetPlayerName(index) & " has left the party.", 13)
                For i = 1 To MAX_PARTY_MEMS
                    If Player(n).Party.PlayerNums(i) = index Then
                        Player(n).Party.PlayerNums(i) = 0
                    ElseIf Player(n).Party.PlayerNums(i) > 0 Then
                        Call PlayerMsg(Player(n).Party.PlayerNums(i), GetPlayerName(index) & " has left the party.", 13)
                    End If
                Next i
                f = 0
                For i = 1 To MAX_PARTY_MEMS
                    If Player(n).Party.PlayerNums(i) > 0 Then
                        f = f + 1
                        Exit For
                    End If
                Next i
                If f = 0 Then
                    Player(n).Party.InParty = NO
                    Player(n).Party.Started = NO
                    Call PlayerMsg(n, "The party has been disbanded.", 13)
                End If
            End If
        End If
        SendParty (index)
        MyScript.ExecuteStatement "Scripts\Main.txt", "LeftGame " & index
        
        Call TextAdd(frmServer.txtText, GetPlayerName(index) & " has disconnected from " & GAME_NAME & ".", True)
        Call SendLeftGame(index)
    End If
    
    Call ClearPlayer(index)
    Call SendOnlineList
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "LeftGame", Err.Number, Err.Description
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
On Error GoTo ErrorHandler
Dim i As Long, n As Long

    n = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If
    Next i
    
    GetTotalMapPlayers = n
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetTotalMapPlayers", Err.Number, Err.Description
End Function

Function GetNpcMaxHP(ByVal NpcNum As Long)
On Error GoTo ErrorHandler

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxHP = 0
        Exit Function
    End If
    
    GetNpcMaxHP = Npc(NpcNum).MaxHp
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetNpcMaxHP", Err.Number, Err.Description
End Function

Function GetNpcMaxMP(ByVal NpcNum As Long)
On Error GoTo ErrorHandler
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxMP = 0
        Exit Function
    End If
        
    GetNpcMaxMP = Npc(NpcNum).MAGI * 2
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetNpcMaxMP", Err.Number, Err.Description
End Function

Function GetNpcMaxSP(ByVal NpcNum As Long)
On Error GoTo ErrorHandler
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxSP = 0
        Exit Function
    End If
        
    GetNpcMaxSP = Npc(NpcNum).SPEED * 2
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetNpcMaxSP", Err.Number, Err.Description
End Function

Function GetPlayerHPRegen(ByVal index As Long)
On Error GoTo ErrorHandler
    If GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
            GetPlayerHPRegen = 0
            Exit Function
        End If

        GetPlayerHPRegen = 1
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetPlayerHPRegen", Err.Number, Err.Description
End Function

Function GetPlayerMPRegen(ByVal index As Long)
On Error GoTo ErrorHandler
    If GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
            GetPlayerMPRegen = 0
            Exit Function
        End If
        
        GetPlayerMPRegen = 1
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetPlayerMPRegen", Err.Number, Err.Description
End Function

Function GetPlayerSPRegen(ByVal index As Long)
On Error GoTo ErrorHandler
    If GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen") = 1 Then
        ' Prevent subscript out of range
        If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
            GetPlayerSPRegen = 0
            Exit Function
        End If
        
        GetPlayerSPRegen = 1
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetPlayerSPRegen", Err.Number, Err.Description
End Function

Sub CheckPlayerLevelUp(ByVal index As Long)
On Error GoTo ErrorHandler
Dim i As Long
Dim d As Long
Dim c As Long
    c = 0
    If GetPlayerLevel(index) < MAX_LEVEL Then
        If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
            Do Until GetPlayerExp(index) < GetPlayerNextLevel(index)
                DoEvents
                If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
                    d = GetPlayerExp(index) - GetPlayerNextLevel(index)
                    Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
                    If GetPlayerLevel(index) = 5 Then
                        Call SetPlayerClass(index, 1)
                        i = 4
                    ElseIf GetPlayerLevel(index) = 15 Then
                        Call SetPlayerClass(index, 2)
                        i = 4
                    ElseIf GetPlayerLevel(index) = 25 Then
                        Call SetPlayerClass(index, 3)
                        i = 4
                    ElseIf GetPlayerLevel(index) = 40 Then
                        Call SetPlayerClass(index, 4)
                        i = 4
                    ElseIf GetPlayerLevel(index) = 60 Then
                        i = 4
                    ElseIf GetPlayerLevel(index) = 80 Then
                        i = 4
                    Else
                        i = 2
                    End If
                    
                    Call PlayerMsg(index, "You have gained " & i & " stat points and 1 vitality and 1 stamina point.", Yellow)
                    Call SetPlayerVIT(index, GetPlayerBaseVIT(index) + 1)
                    Call SetPlayerSPEED(index, GetPlayerBaseSPEED(index) + 1)
                    Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)
                    
                    Call SetPlayerExp(index, d)
                    c = c + 1
                End If
                Call PlayerPoints(index)
            Loop
            If c > 1 Then
                Call BattleMsg(index, "You leveled " & c & " times!", 6, index)
            Else
                Call BattleMsg(index, "You leveled!", 6, index)
            End If
            Call PlaySound(index, "magic32.wav")
            Call SendDataToMap(GetPlayerMap(index), "levelup" & SEP_CHAR & index & SEP_CHAR & END_CHAR)

            Call SetPlayerHP(index, GetPlayerMaxHP(index))
            Call SetPlayerMP(index, GetPlayerMaxMP(index))
            Call SetPlayerSP(index, GetPlayerMaxSP(index))

            Call PlayerClass(index)
        End If
    End If
    
    If GetPlayerLevel(index) = MAX_LEVEL Then
        Call SetPlayerExp(index, Experience(MAX_LEVEL))
    End If
    
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendStats(index)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "CheckPlayerLevelUp", Err.Number, Err.Description
End Sub

Sub CastSpell(ByVal index As Long, ByVal SpellSlot As Long)
On Error GoTo ErrorHandler
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
        Call PlayerMsg(index, "You do not have this spell!", BrightRed)
        Exit Sub
    End If
    
    i = GetSpellReqLevel(index, SpellNum)

    ' Check if they have enough MP
    If GetPlayerMP(index) < Spell(SpellNum).MPCost Then
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
            Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
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
            If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
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
                Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
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
                            Call SetPlayerSP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                            Call SendSP(n)
                    End Select
                    
                    ' Take away the mana points
                    Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
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
            Call SetPlayerMP(index, GetPlayerMP(index) - Spell(SpellNum).MPCost)
            Call SendMP(index)
            Casted = True
        Else
            Call PlayerMsg(index, "Could not cast spell!", BrightRed)
        End If
    End If

    If Casted = True Then
        Player(index).AttackTimer = GetTickCount
        Player(index).CastedSpell = YES
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "CastSpell", Err.Number, Err.Description
End Sub
Function GetSpellReqLevel(ByVal index As Long, ByVal SpellNum As Long)
On Error GoTo ErrorHandler
    GetSpellReqLevel = Spell(SpellNum).LevelReq ' - Int(GetClassMAGI(GetPlayerClass(index)) / 4)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "GetSpellReqLevel", Err.Number, Err.Description
End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "CanPlayerCriticalHit", Err.Number, Err.Description
End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "CanPlayerBlockHit", Err.Number, Err.Description
End Function

Sub CheckEquippedItems(ByVal index As Long)
On Error GoTo ErrorHandler

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
    
        Slot = GetPlayerBootsSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_BOOTS Then
                Call SetPlayerBootsSlot(index, 0)
            End If
        Else
            Call SetPlayerBootsSlot(index, 0)
        End If
    End If
    
        Slot = GetPlayerGlovesSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_GLOVES Then
                Call SetPlayerGlovesSlot(index, 0)
            End If
        Else
            Call SetPlayerGlovesSlot(index, 0)
        End If
    End If
    
        Slot = GetPlayerRingSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_RING Then
                Call SetPlayerRingSlot(index, 0)
            End If
        Else
            Call SetPlayerRingSlot(index, 0)
        End If
    End If
    
        Slot = GetPlayerAmuletSlot(index)
    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(index, Slot)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_AMULET Then
                Call SetPlayerAmuletSlot(index, 0)
            End If
        Else
            Call SetPlayerAmuletSlot(index, 0)
        End If
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "CheckEquippedItems", Err.Number, Err.Description
End Sub
