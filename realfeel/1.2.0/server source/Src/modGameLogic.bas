Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
Dim WeaponSlot As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    GetPlayerDamage = GetPlayerSTR(Index)
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(Index)
        
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2
        
        If Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data5 = UNBREAKABLE Then Exit Function
        
        Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) - 1)
        
        If GetPlayerInvItemDur(Index, WeaponSlot) <= 0 Then
            Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " has broken.", Yellow)
            Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
        Else
            If GetPlayerInvItemDur(Index, WeaponSlot) <= 5 Then
                Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " is about to break!", Yellow)
            End If
        End If
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerDamage", Err.Number, Err.Description)
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
Dim ArmorSlot As Long, HelmSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ArmorSlot = GetPlayerArmorSlot(Index)
    HelmSlot = GetPlayerHelmetSlot(Index)
    GetPlayerProtection = Int(GetPlayerDEF(Index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data2
        
        If Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data5 = UNBREAKABLE Then Exit Function
        
        Call SetPlayerInvItemDur(Index, ArmorSlot, GetPlayerInvItemDur(Index, ArmorSlot) - 1)
        
        If GetPlayerInvItemDur(Index, ArmorSlot) <= 0 Then
            Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " has broken.", Yellow)
            Call TakeItem(Index, GetPlayerInvItemNum(Index, ArmorSlot), 0)
        Else
            If GetPlayerInvItemDur(Index, ArmorSlot) <= 5 Then
                Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " is about to break!", Yellow)
            End If
        End If
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, HelmSlot)).Data2
        
        If Item(GetPlayerInvItemNum(Index, HelmSlot)).Data5 = UNBREAKABLE Then Exit Function
        
        Call SetPlayerInvItemDur(Index, HelmSlot, GetPlayerInvItemDur(Index, HelmSlot) - 1)
        
        If GetPlayerInvItemDur(Index, HelmSlot) <= 0 Then
            Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " has broken.", Yellow)
            Call TakeItem(Index, GetPlayerInvItemNum(Index, HelmSlot), 0)
        Else
            If GetPlayerInvItemDur(Index, HelmSlot) <= 5 Then
                Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " is about to break!", Yellow)
            End If
        End If
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerProtection", Err.Number, Err.Description)
End Function

Function FindOpenPlayerSlot() As Long
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modGameLogic.bas", "FindOpenPlayerSlot", Err.Number, Err.Description)
End Function

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
'On Error GoTo errorhandler:
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "FindOpenInvSlot", Err.Number, Err.Description)
End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
'On Error GoTo errorhandler:
Dim i As Long
    
    FindOpenBankSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_BANK_ITEMS
            If GetPlayerBankItemNum(Index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i
    End If
    
    For i = 1 To MAX_BANK_ITEMS
        ' Try to find an open free slot
        If GetPlayerBankItemNum(Index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "FindOpenBankSlot", Err.Number, Err.Description)
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modGameLogic.bas", "FindOpenMapItemSlot", Err.Number, Err.Description)
End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "FindOpenSpellSlot", Err.Number, Err.Description)
End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
'On Error GoTo errorhandler:
Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "HasSpell", Err.Number, Err.Description)
End Function

Function TotalOnlinePlayers() As Long
'On Error GoTo errorhandler:
Dim i As Long

    TotalOnlinePlayers = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "TotalOnlinePlayers", Err.Number, Err.Description)
End Function

Function FindPlayer(ByVal Name As String) As Long
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "FindPlayer", Err.Number, Err.Description)
End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
'On Error GoTo errorhandler:
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
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "HasItem", Err.Number, Err.Description)
End Function

Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
'On Error GoTo errorhandler:
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
                End Select

                
                n = Item(GetPlayerInvItemNum(Index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) Then
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "TakeItem", Err.Number, Err.Description)
End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
'On Error GoTo errorhandler:
Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_BANK_ITEMS
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                'Debug.Print "ItemValue: " & ItemVal & " - BankItemValue: " & GetPlayerBankItemValue(Index, i)
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerBankItemValue(Index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerBankItemValue(Index, i, GetPlayerBankItemValue(Index, i) - ItemVal)
                    Call SendUpdateBankItemTo(Index, i)
                    Exit Sub
                End If
            Else
                TakeItem = True
            End If
                            
            If TakeItem = True Then
                Call SetPlayerBankItemNum(Index, i, 0)
                Call SetPlayerBankItemValue(Index, i, 0)
                Call SetPlayerBankItemDur(Index, i, 0)
                
                ' Send the inventory update
                Call SendUpdateBankItemTo(Index, i)
                'Call KillData(Index, i, "BANK")
                'Debug.Print "Item taken!"
                'Debug.Print "ItemValue: " & ItemVal & " - BankItemValue: " & GetPlayerBankItemValue(Index, i)
                Exit Sub
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "TakeBankItem", Err.Number, Err.Description)
End Sub

Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
'On Error GoTo errorhandler:
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
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "GiveItem", Err.Number, Err.Description)
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
'On Error GoTo errorhandler:
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = FindOpenBankSlot(Index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerBankItemNum(Index, i, ItemNum)
        Call SetPlayerBankItemValue(Index, i, GetPlayerBankItemValue(Index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerBankItemDur(Index, i, Item(ItemNum).Data1)
        End If
        
        Call SendUpdateBankItemTo(Index, i)
    Else
        Call PlayerMsg(Index, "Your bank is full.", Red)
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "GiveBankItem", Err.Number, Err.Description)
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
'On Error GoTo errorhandler:
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, x, y)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SpawnItem", Err.Number, Err.Description)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
'On Error GoTo errorhandler:
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
            
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SpawnItemSlot", Err.Number, Err.Description)
End Sub

Sub SpawnAllMapsItems()
'On Error GoTo errorhandler:
Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SpawnAllMapsItems", Err.Number, Err.Description)
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
'On Error GoTo errorhandler:
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
            If (Map(MapNum).Tile(x, y).Item = True) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(x, y).ItemNum).Type = ITEM_TYPE_CURRENCY Then
                    If Map(MapNum).Tile(x, y).ItemValue <= 0 Then
                        Call SpawnItem(Map(MapNum).Tile(x, y).ItemNum, 1, MapNum, x, y)
                    Else
                        Call SpawnItem(Map(MapNum).Tile(x, y).ItemNum, Map(MapNum).Tile(x, y).ItemValue, MapNum, x, y)
                    End If
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).ItemNum, Map(MapNum).Tile(x, y).ItemValue, MapNum, x, y)
                End If
            End If
        Next x
    Next y
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SpawnMapItems", Err.Number, Err.Description)
End Sub

Sub PlayerMapGetItem(ByVal Index As Long)
'On Error GoTo errorhandler:
Dim i As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String

    If IsPlaying(Index) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Index)
    
    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).x = GetPlayerX(Index)) And (MapItem(MapNum, i).y = GetPlayerY(Index)) Then
                ' Find open slot
                n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
                
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
                    If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Then
                        Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, i).Value)
                        Msg = "You picked up " & MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(Index, n, 0)
                        Msg = "You picked up a " & Trim$(Item(GetPlayerInvItemNum(Index, n)).Name) & "."
                    End If
                    Call SetPlayerInvItemDur(Index, n, MapItem(MapNum, i).Dur)
                        
                    ' Erase item from the map
                    MapItem(MapNum, i).Num = 0
                    MapItem(MapNum, i).Value = 0
                    MapItem(MapNum, i).Dur = 0
                    MapItem(MapNum, i).x = 0
                    MapItem(MapNum, i).y = 0

                    Call SendInventoryUpdate(Index, n)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                    Call PlayerMsg(Index, Msg, Yellow)
                    Exit Sub
                Else
                    Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                    Exit Sub
                End If
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "PlayerMapGetItem", Err.Number, Err.Description)
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Ammount As Long)
'On Error GoTo errorhandler:
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(Index))
        
        If i <> 0 Then
            MapItem(GetPlayerMap(Index), i).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(Index) Then
                        Call SetPlayerArmorSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(Index) Then
                        Call SetPlayerWeaponSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(Index) Then
                        Call SetPlayerHelmetSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(Index) Then
                        Call SetPlayerShieldSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
            MapItem(GetPlayerMap(Index), i).x = GetPlayerX(Index)
            MapItem(GetPlayerMap(Index), i).y = GetPlayerY(Index)
                        
            If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Ammount >= GetPlayerInvItemValue(Index, InvNum) Then
                    MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(Index), i).Value = Ammount
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Ammount & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Ammount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(Index), i).Value = 0
                If Item(GetPlayerInvItemNum(Index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " " & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Data1 & ".", Yellow)
                Else
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(Index, InvNum, 0)
                Call SetPlayerInvItemValue(Index, InvNum, 0)
                Call SetPlayerInvItemDur(Index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(Index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Ammount, MapItem(GetPlayerMap(Index), i).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
        Else
            Call PlayerMsg(Index, "To many items already on the ground.", BrightRed)
        End If
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "PlayerMapDropItem", Err.Number, Err.Description)
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
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
        MapNpc(MapNum, MapNpcNum).MaxHP = GetNpcMaxHP(NpcNum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)
        
        'Set the map npc's default behavior
        MapNpc(MapNum, MapNpcNum).Behavior = Npc(NpcNum).Behavior
                
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)
        
        ' Well try 100 times to randomly place the sprite
        For i = 1 To 100
            x = Int(Rnd * MAX_MAPX)
            y = Int(Rnd * MAX_MAPY)
            
            ' Check if the tile is walkable
            If Map(MapNum).Tile(x, y).Walkable = True Then
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
                    If Map(MapNum).Tile(x, y).Walkable = True Then
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SpawnNpc", Err.Number, Err.Description)
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SpawnMapNpcs", Err.Number, Err.Description)
End Sub

Sub SpawnAllMapNpcs()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SpawnAllMapNpcs", Err.Number, Err.Description)
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
'On Error GoTo errorhandler:
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
                                If GetPlayerLevel(Attacker) < 5 Then
                                    Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 5 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
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
                                If GetPlayerLevel(Attacker) < 5 Then
                                    Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 5 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
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
                                If GetPlayerLevel(Attacker) < 5 Then
                                    Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 5 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
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
                                If GetPlayerLevel(Attacker) < 5 Then
                                    Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
                                Else
                                    If GetPlayerLevel(Victim) < 5 Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
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
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "CanAttackPlayer", Err.Number, Err.Description)
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
'On Error GoTo errorhandler:
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
        If NpcNum > 0 And GetTickCount > Player(Attacker).AttackTimer + 950 Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If (MapNpc(MapNum, MapNpcNum).y + 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x = GetPlayerX(Attacker)) Then
                        If MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & " says: " & Npc(NpcNum).AttackSay, SayColor)
                        End If
                    End If
                
                Case DIR_DOWN
                    If (MapNpc(MapNum, MapNpcNum).y - 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x = GetPlayerX(Attacker)) Then
                        If MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & " says: " & Npc(NpcNum).AttackSay, SayColor)
                        End If
                    End If
                
                Case DIR_LEFT
                    If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x + 1 = GetPlayerX(Attacker)) Then
                        If MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & " says: " & Npc(NpcNum).AttackSay, SayColor)
                        End If
                    End If
                
                Case DIR_RIGHT
                    If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x - 1 = GetPlayerX(Attacker)) Then
                        If MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & " says: " & Npc(NpcNum).AttackSay, SayColor)
                        End If
                    End If
            End Select
        End If
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "CanAttackNpc", Err.Number, Err.Description)
End Function

Function CanNpcAttackNpc(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal MapNum As Long) As Boolean
'On Error GoTo errorhandler:
Dim NpcNum As Long, VictimNum As Long
    CanNpcAttackNpc = False
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(MapNum, MapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    VictimNum = MapNpc(MapNum, Victim).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    
    'Make sure both of the npcs are valid
    If NpcNum > 0 And VictimNum > 0 Then
            ' Check if at same coordinates
            If (MapNpc(MapNum, Victim).y + 1 = MapNpc(MapNum, MapNpcNum).y) And (MapNpc(MapNum, Victim).x = MapNpc(MapNum, MapNpcNum).x) Then
                CanNpcAttackNpc = True
            Else
                If (MapNpc(MapNum, Victim).y - 1 = MapNpc(MapNum, MapNpcNum).y) And (MapNpc(MapNum, Victim).x = MapNpc(MapNum, MapNpcNum).x) Then
                    CanNpcAttackNpc = True
                Else
                    If (MapNpc(MapNum, Victim).y = MapNpc(MapNum, MapNpcNum).y) And (MapNpc(MapNum, Victim).x + 1 = MapNpc(MapNum, MapNpcNum).x) Then
                        CanNpcAttackNpc = True
                    Else
                        If (MapNpc(MapNum, Victim).y = MapNpc(MapNum, MapNpcNum).y) And (MapNpc(MapNum, Victim).x - 1 = MapNpc(MapNum, MapNpcNum).x) Then
                            CanNpcAttackNpc = True
                        End If
                    End If
                End If
            End If
    End If
    
    'Debug.Print CanNpcAttackNpc
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "CanNpcAttackNpc", Err.Number, Err.Description)
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
'On Error GoTo errorhandler:
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
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "CanNpcAttackPlayer", Err.Number, Err.Description)
End Function

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
'On Error GoTo errorhandler:
Dim EXP As Long
Dim START_MAP As Long, START_X As Long, START_Y As Long
Dim n As Long
Dim i As Long

    START_MAP = Class(Player(Victim).Char(Player(Victim).CharNum).Class).Map
    START_X = Class(Player(Victim).Char(Player(Victim).CharNum).Class).x
    START_Y = Class(Player(Victim).Char(Player(Victim).CharNum).Class).y

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
    ' Send weapon sound message
    If GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) <> 0 Then Call SendDataToMap(GetPlayerMap(Attacker), "PLAYSOUND" & SEP_CHAR & Trim$(Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).Sound) & SEP_CHAR & END_CHAR)
        
    If Damage >= GetPlayerHP(Victim) Then
        ' Set HP to nothing
        Call SetPlayerHP(Victim, 0)
        
        ' Send the player death animation
        Call SendDataToMap(GetPlayerMap(Victim), "KILLPLAYER" & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
        
        ' Check for a weapon and say damage
        'If n = 0 Then
        '    Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
        '    Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        'Else
        '    Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
        '    Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        'End If
        
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
        EXP = Int(GetPlayerExp(Victim) / 10)
        
        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 0
        End If
        
        If EXP = 0 Then
            Call PlayerMsg(Victim, "You lost no experience points.", BrightRed)
            Call PlayerMsg(Attacker, "You received no experience points from that weak insignificant player.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - EXP)
            Call PlayerMsg(Victim, "You lost " & EXP & " experience points.", BrightRed)
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
            Call PlayerMsg(Attacker, "You got " & EXP & " experience points for killing " & GetPlayerName(Victim) & ".", BrightBlue)
            Call SendDataTo(Victim, "experience" & SEP_CHAR & GetPlayerExp(Victim) & SEP_CHAR & (GetPlayerNextLevel(Victim) - GetPlayerExp(Victim)) & SEP_CHAR & END_CHAR)
            Call SendDataTo(Attacker, "experience" & SEP_CHAR & GetPlayerExp(Attacker) & SEP_CHAR & (GetPlayerNextLevel(Attacker) - GetPlayerExp(Attacker)) & SEP_CHAR & END_CHAR)
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
        
        ' Send packet to display damage text
        Call SendDataToMap(GetPlayerMap(Victim), "ATTACKPLAYER" & SEP_CHAR & Victim & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
        
        ' Check for a weapon and say damage
        'If n = 0 Then
        '    Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
        '    Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        'Else
        '    Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
        '    Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        'End If
    End If
    
    ' Reset timer for attacking
    Player(Attacker).AttackTimer = GetTickCount
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "AttackPlayer", Err.Number, Err.Description)
End Sub

Sub NpcAttackNpc(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
'Debug.Print "NPC ATTACKING (" & Npc(MapNpc(MapNum, MapNpcNum).Num).Name & ")"
'Debug.Print "NPC DEFENDING (" & Npc(MapNpc(MapNum, Victim).Num).Name & ")"
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Victim <= 0 Or Victim > MAX_MAP_NPCS Or Damage < 0 Then
        'Debug.Print "No damage!"
        Exit Sub
    End If
    
    ' Check for subscript out of range
    If MapNpc(MapNum, MapNpcNum).Num <= 0 Or MapNpc(MapNum, Victim).Num <= 0 Then
        'Debug.Print "Problem with mapnpc numbers!"
        Exit Sub
    End If
    
    ' Send this packet if player is on map so they can see the npc attacking
    Call SendDataToMap(MapNum, "NPCATTACK" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    
    If Damage >= MapNpc(MapNum, Victim).HP Then
        'Kill Victim
        MapNpc(MapNum, Victim).Num = 0
        MapNpc(MapNum, Victim).SpawnWait = GetTickCount
        MapNpc(MapNum, Victim).HP = 0
        'Execute NpcDeath Sub
        MyScript.ExecuteStatement "main.txt", "NpcDeath " & MapNpcNum & "," & Victim & "," & MapNpcNum & "," & MapNum & "," & False
        
        'Send map data
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    'Debug.Print "Killed npc!"
        ' Set NPC target to 0
        MapNpc(MapNum, MapNpcNum).Target = 0
    'Debug.Print "Attacker's target reset!"
    Else
        ' Npc not dead, just do the damage
        MapNpc(MapNum, Victim).HP = MapNpc(MapNum, Victim).HP - Damage
        
        ' Send npc attack packet - displays damage
        Call SendDataToMap(GetPlayerMap(MapNum), "ATTACKNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
        
    'Debug.Print "Did damage!"
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "NpcAttackNpc", Err.Number, Err.Description)
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
'On Error GoTo errorhandler:
Dim Name As String
Dim EXP As Long
Dim MapNum As Long
Dim START_MAP As Long, START_X As Long, START_Y As Long

    START_MAP = Class(Player(Victim).Char(Player(Victim).CharNum).Class).Map
    START_X = Class(Player(Victim).Char(Player(Victim).CharNum).Class).x
    START_Y = Class(Player(Victim).Char(Player(Victim).CharNum).Class).y

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
    Name = Trim$(Npc(MapNpc(MapNum, MapNpcNum).Num).Name)
    
    If Damage >= GetPlayerHP(Victim) Then
        ' Send the player death animation
        Call SendDataToMap(GetPlayerMap(Victim), "KILLPLAYER" & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    
        ' Say damage
        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        
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
        EXP = Int(GetPlayerExp(Victim) / 3)
        
        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 0
        End If
        
        If EXP = 0 Then
            Call PlayerMsg(Victim, "You lost no experience points.", BrightRed)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - EXP)
            Call PlayerMsg(Victim, "You lost " & EXP & " experience points.", BrightRed)
            Call SendDataTo(Victim, "experience" & SEP_CHAR & GetPlayerExp(Victim) & SEP_CHAR & (GetPlayerNextLevel(Victim) - GetPlayerExp(Victim)) & SEP_CHAR & END_CHAR)
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
        
        ' Send packet to display damage text
        Call SendDataToMap(GetPlayerMap(Victim), "ATTACKPLAYER" & SEP_CHAR & Victim & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
        
        ' Say damage
        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "NpcAttackPlayer", Err.Number, Err.Description)
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
'On Error GoTo errorhandler:
Dim Name As String
Dim EXP As Long
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
    
    ' Send message to play sound
    If n <> 0 Then Call SendDataToMap(GetPlayerMap(Attacker), "PLAYSOUND" & SEP_CHAR & Trim$(Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).Sound) & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    Name = Trim$(Npc(NpcNum).Name)
        
    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        'If n = 0 Then
        '    Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points, killing it.", BrightRed)
        'Else
        '    Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points, killing it.", BrightRed)
        'End If
                        
        EXP = Npc(NpcNum).EXP
        
        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 1
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If Player(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
            Call PlayerMsg(Attacker, "You have gained " & EXP & " experience points.", BrightBlue)
            Call SendDataTo(Attacker, "experience" & SEP_CHAR & GetPlayerExp(Attacker) & SEP_CHAR & (GetPlayerNextLevel(Attacker) - GetPlayerExp(Attacker)) & SEP_CHAR & END_CHAR)
        Else
            EXP = EXP / 2
            
            If EXP < 0 Then
                EXP = 1
            End If
            
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
            Call PlayerMsg(Attacker, "You have gained " & EXP & " party experience points.", BrightBlue)
            Call SendDataTo(Attacker, "experience" & SEP_CHAR & GetPlayerExp(Attacker) & SEP_CHAR & (GetPlayerNextLevel(Attacker) - GetPlayerExp(Attacker)) & SEP_CHAR & END_CHAR)
            
            n = Player(Attacker).PartyPlayer
            If n > 0 Then
                Call SetPlayerExp(n, GetPlayerExp(n) + EXP)
                Call PlayerMsg(n, "You have gained " & EXP & " party experience points.", BrightBlue)
                Call SendDataTo(Attacker, "experience" & SEP_CHAR & GetPlayerExp(n) & SEP_CHAR & (GetPlayerNextLevel(n) - GetPlayerExp(n)) & SEP_CHAR & END_CHAR)
            End If
        End If
                                
        ' Drop the goods if they get it
        n = Int(Rnd * Npc(NpcNum).DropChance) + 1
        If n = 1 Then
            Call SpawnItem(Npc(NpcNum).DropItem, Npc(NpcNum).DropItemValue, MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).Num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        'Debug.Print MapNpcNum & "," & NpcNum & "," & MapNum
        
        'Execute NpcDeath Sub
        MyScript.ExecuteStatement "main.txt", "NpcDeath " & Attacker & "," & NpcNum & "," & MapNpcNum & "," & MapNum & "," & True
        
        'Send map data
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
        
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
        
        ' Send packet to display damage text
        Call SendDataToMap(GetPlayerMap(Attacker), "ATTACKNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
        
        ' Check for a weapon and say damage
        'If n = 0 Then
        '    Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        'Else
        '    Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
        'End If
        
        ' Check if we should send a message
        If MapNpc(MapNum, MapNpcNum).Target = 0 And MapNpc(MapNum, MapNpcNum).Target <> Attacker Then
            If Trim$(Npc(NpcNum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
            End If
        End If
        
        If MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_TRAITOR And MapNpc(MapNum, MapNpcNum).Behavior <> NPC_BEHAVIOR_FOLLOW Then
            ' Set the NPC target to the player
            MapNpc(MapNum, MapNpcNum).Target = Attacker
        End If
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If MapNpc(MapNum, MapNpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum, i).Num = MapNpc(MapNum, MapNpcNum).Num Then
                    MapNpc(MapNum, i).Target = Attacker
                End If
            Next i
        End If
    End If
       
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "AttackNpc", Err.Number, Err.Description)
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim ShopNum As Long, OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    ShopNum = Map(GetPlayerMap(Index)).Shop
    If ShopNum > 0 Then
        If Trim$(Shop(ShopNum).LeaveSay) <> "" Then
            Call PlayerMsg(Index, Trim$(Shop(ShopNum).Name) & " says, '" & Trim$(Shop(ShopNum).LeaveSay) & "'", SayColor)
        End If
    End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    Call SendLeaveMap(Index, OldMap)
    
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    
    ' Check if there is an npc on the map and say hello if so
    ShopNum = Map(GetPlayerMap(Index)).Shop
    If ShopNum > 0 Then
        If Trim$(Shop(ShopNum).JoinSay) <> "" Then
            Call PlayerMsg(Index, Trim$(Shop(ShopNum).Name) & " says, '" & Trim$(Shop(ShopNum).JoinSay) & "'", SayColor)
        End If
    End If
            
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    
    Player(Index).GettingMap = YES
    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "PlayerWarp", Err.Number, Err.Description)
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim MapNum As Long
Dim x As Long
Dim y As Long
Dim i As Long
Dim Moved As Byte
Dim EXP As Long
Dim START_MAP As Long, START_X As Long, START_Y As Long

    START_MAP = Class(Player(Index).Char(Player(Index).CharNum).Class).Map
    START_X = Class(Player(Index).Char(Player(Index).CharNum).Class).x
    START_Y = Class(Player(Index).Char(Player(Index).CharNum).Class).y

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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Blocked = False Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).South = False Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).North = False Then
                    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Key = False Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Key = True And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                                Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                        
                                Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                                Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        
                                'If a healing tile, heal the player -smchronos
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Heal = True Then
                                    Call SetPlayerHP(Index, GetPlayerHP(Index) + Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).HealValue)
                            
                                    Call PlayerMsg(Index, "You have been healed " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).HealValue & " health!", Cyan)
                            
                                    'check to see if the added health exceeds the total hp of the player, if so, set the player's hp to his max hp
                                    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                                    Call SendHP(Index)
                                End If
                                
                                'If a damage tile, hurt the player -smchronos
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Damage = True Then
                                    Call SetPlayerHP(Index, GetPlayerHP(Index) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DamageValue)
                                    
                                    Call PlayerMsg(Index, "You have been hurt " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DamageValue & " health!", Red)
                            
                                    'check to see if the health removed sets the player's hp to 0 or less, if so, run death scenario
                                    If GetPlayerHP(Index) <= 0 Then
                                        ' Set HP to nothing
                                        Call SetPlayerHP(Index, 0)
        
                                        ' Player is dead
                                        Call GlobalMsg(GetPlayerName(Index) & " has died...", BrightRed)
        
                                        ' Drop all worn items by player
                                        If GetPlayerWeaponSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerWeaponSlot(Index), 0)
                                        End If
                                        If GetPlayerArmorSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerArmorSlot(Index), 0)
                                        End If
                                        If GetPlayerHelmetSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerHelmetSlot(Index), 0)
                                        End If
                                        If GetPlayerShieldSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerShieldSlot(Index), 0)
                                        End If

                                        ' Calculate exp to take away
                                        EXP = Int(GetPlayerExp(Index) / 10)
        
                                        ' Make sure we dont get less then 0
                                        If EXP < 0 Then
                                            EXP = 0
                                        End If
        
                                        If EXP = 0 Then
                                            Call PlayerMsg(Index, "You lost no experience points.", BrightRed)
                                        Else
                                            Call SetPlayerExp(Index, GetPlayerExp(Index) - EXP)
                                            Call PlayerMsg(Index, "You lost " & EXP & " experience points.", BrightRed)
                                            Call SendDataTo(Index, "experience" & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & (GetPlayerNextLevel(Index) - GetPlayerExp(Index)) & SEP_CHAR & END_CHAR)
                                        End If
                
                                        ' Warp player away
                                        Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        
                                        ' Restore vitals
                                        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                                        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
                                        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
                                        Call SendHP(Index)
                                        Call SendMP(Index)
                                        Call SendSP(Index)
                                    End If
                                    If GetPlayerHP(Index) > 0 Then Call SendHP(Index)
                                End If
                        
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Shop = True Then
                                    'Send the shop data to the player
                                    Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).ShopNum)
                                End If
                        
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Bank = True Then
                                    'Send the player bank data
                                    Call SendBankInv(Index)
                                End If
                        
                                Moved = YES
                            End If
                        End If
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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Blocked = False Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).North = False Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).South = False Then
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Key = False Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Key = True And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                                Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                        
                                Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                                Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        
                                'If a healing tile, heal the player -smchronos
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Heal = True Then
                                    Call SetPlayerHP(Index, GetPlayerHP(Index) + Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).HealValue)
                            
                                    Call PlayerMsg(Index, "You have been healed " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).HealValue & " health!", Cyan)
                            
                                    'check to see if the added health exceeds the total hp of the player, if so, set the player's hp to his max hp
                                    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                                    Call SendHP(Index)
                                End If
                                
                                'If a damage tile, hurt the player -smchronos
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Damage = True Then
                                    Call SetPlayerHP(Index, GetPlayerHP(Index) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DamageValue)
                            
                                    Call PlayerMsg(Index, "You have been hurt " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DamageValue & " health!", Red)
                            
                                    'check to see if the health removed sets the player's hp to 0 or less, if so, run death scenario
                                    If GetPlayerHP(Index) <= 0 Then
                                        ' Set HP to nothing
                                        Call SetPlayerHP(Index, 0)
        
                                        ' Player is dead
                                        Call GlobalMsg(GetPlayerName(Index) & " has died...", BrightRed)
        
                                        ' Drop all worn items by player
                                        If GetPlayerWeaponSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerWeaponSlot(Index), 0)
                                        End If
                                        If GetPlayerArmorSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerArmorSlot(Index), 0)
                                        End If
                                        If GetPlayerHelmetSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerHelmetSlot(Index), 0)
                                        End If
                                        If GetPlayerShieldSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerShieldSlot(Index), 0)
                                        End If

                                        ' Calculate exp to take away
                                        EXP = Int(GetPlayerExp(Index) / 10)
        
                                        ' Make sure we dont get less then 0
                                        If EXP < 0 Then
                                            EXP = 0
                                        End If
        
                                        If EXP = 0 Then
                                            Call PlayerMsg(Index, "You lost no experience points.", BrightRed)
                                        Else
                                            Call SetPlayerExp(Index, GetPlayerExp(Index) - EXP)
                                            Call PlayerMsg(Index, "You lost " & EXP & " experience points.", BrightRed)
                                            Call SendDataTo(Index, "experience" & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & (GetPlayerNextLevel(Index) - GetPlayerExp(Index)) & SEP_CHAR & END_CHAR)
                                        End If
                
                                        ' Warp player away
                                        Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        
                                        ' Restore vitals
                                        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                                        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
                                        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
                                        Call SendHP(Index)
                                        Call SendMP(Index)
                                        Call SendSP(Index)
                                    End If
                                    If GetPlayerHP(Index) > 0 Then Call SendHP(Index)
                                End If
                        
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Shop = True Then
                                    'Send the shop data to the player
                                    Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).ShopNum)
                                End If
                        
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Bank = True Then
                                    'Send the player bank data
                                    Call SendBankInv(Index)
                                End If
                        
                                Moved = YES
                            End If
                        End If
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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Blocked = False Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).East = False Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).West = False Then
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Key = False Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Key = True And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                                Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                        
                                Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                                Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        
                                'If a healing tile, heal the player -smchronos
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Heal = True Then
                                    Call SetPlayerHP(Index, GetPlayerHP(Index) + Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).HealValue)
                            
                                    Call PlayerMsg(Index, "You have been healed " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).HealValue & " health!", Cyan)
                            
                                    'check to see if the added health exceeds the total hp of the player, if so, set the player's hp to his max hp
                                    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                                    Call SendHP(Index)
                                End If
                        
                                'If a damage tile, hurt the player -smchronos
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Damage = True Then
                                    Call SetPlayerHP(Index, GetPlayerHP(Index) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DamageValue)
                            
                                    Call PlayerMsg(Index, "You have been hurt " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DamageValue & " health!", Red)
                            
                                    'check to see if the health removed sets the player's hp to 0 or less, if so, run death scenario
                                    If GetPlayerHP(Index) <= 0 Then
                                        ' Set HP to nothing
                                        Call SetPlayerHP(Index, 0)
        
                                        ' Player is dead
                                        Call GlobalMsg(GetPlayerName(Index) & " has died...", BrightRed)
        
                                        ' Drop all worn items by player
                                        If GetPlayerWeaponSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerWeaponSlot(Index), 0)
                                        End If
                                        If GetPlayerArmorSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerArmorSlot(Index), 0)
                                        End If
                                        If GetPlayerHelmetSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerHelmetSlot(Index), 0)
                                        End If
                                        If GetPlayerShieldSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerShieldSlot(Index), 0)
                                        End If

                                        ' Calculate exp to take away
                                        EXP = Int(GetPlayerExp(Index) / 10)
        
                                        ' Make sure we dont get less then 0
                                        If EXP < 0 Then
                                            EXP = 0
                                        End If
        
                                        If EXP = 0 Then
                                            Call PlayerMsg(Index, "You lost no experience points.", BrightRed)
                                        Else
                                            Call SetPlayerExp(Index, GetPlayerExp(Index) - EXP)
                                            Call PlayerMsg(Index, "You lost " & EXP & " experience points.", BrightRed)
                                            Call SendDataTo(Index, "experience" & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & (GetPlayerNextLevel(Index) - GetPlayerExp(Index)) & SEP_CHAR & END_CHAR)
                                        End If
                
                                        ' Warp player away
                                        Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        
                                        ' Restore vitals
                                        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                                        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
                                        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
                                        Call SendHP(Index)
                                        Call SendMP(Index)
                                        Call SendSP(Index)
                                    End If
                                    If GetPlayerHP(Index) > 0 Then Call SendHP(Index)
                                End If
                        
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Shop = True Then
                                    'Send the shop data to the player
                                    Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).ShopNum)
                                End If
                        
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Bank = True Then
                                    'Send the player bank data
                                    Call SendBankInv(Index)
                                End If
                        
                                Moved = YES
                            End If
                        End If
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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Blocked = False Then
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).West = False Then
                        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).East = False Then
                    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Key = False Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Key = True And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                                Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                        
                                Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                                Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        
                                'If a healing tile, heal the player -smchronos
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Heal = True Then
                                    Call SetPlayerHP(Index, GetPlayerHP(Index) + Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).HealValue)
                            
                                    Call PlayerMsg(Index, "You have been healed " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).HealValue & " health!", Cyan)
                            
                                    'check to see if the added health exceeds the total hp of the player, if so, set the player's hp to his max hp
                                    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                                    Call SendHP(Index)
                                End If
                        
                                'If a damage tile, hurt the player -smchronos
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Damage = True Then
                                    Call SetPlayerHP(Index, GetPlayerHP(Index) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DamageValue)
                            
                                    Call PlayerMsg(Index, "You have been hurt " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DamageValue & " health!", Red)
                            
                                    'check to see if the health removed sets the player's hp to 0 or less, if so, run death scenario
                                    If GetPlayerHP(Index) <= 0 Then
                                        ' Set HP to nothing
                                        Call SetPlayerHP(Index, 0)
        
                                        ' Player is dead
                                        Call GlobalMsg(GetPlayerName(Index) & " has died...", BrightRed)
        
                                        ' Drop all worn items by player
                                        If GetPlayerWeaponSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerWeaponSlot(Index), 0)
                                        End If
                                        If GetPlayerArmorSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerArmorSlot(Index), 0)
                                        End If
                                        If GetPlayerHelmetSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerHelmetSlot(Index), 0)
                                        End If
                                        If GetPlayerShieldSlot(Index) > 0 Then
                                            Call PlayerMapDropItem(Index, GetPlayerShieldSlot(Index), 0)
                                        End If

                                        ' Calculate exp to take away
                                        EXP = Int(GetPlayerExp(Index) / 10)
        
                                        ' Make sure we dont get less then 0
                                        If EXP < 0 Then
                                            EXP = 0
                                        End If
        
                                        If EXP = 0 Then
                                            Call PlayerMsg(Index, "You lost no experience points.", BrightRed)
                                        Else
                                            Call SetPlayerExp(Index, GetPlayerExp(Index) - EXP)
                                            Call PlayerMsg(Index, "You lost " & EXP & " experience points.", BrightRed)
                                            Call SendDataTo(Index, "experience" & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & (GetPlayerNextLevel(Index) - GetPlayerExp(Index)) & SEP_CHAR & END_CHAR)
                                        End If
                
                                        ' Warp player away
                                        Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        
                                        ' Restore vitals
                                        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                                        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
                                        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
                                        Call SendHP(Index)
                                        Call SendMP(Index)
                                        Call SendSP(Index)
                                    End If
                                    If GetPlayerHP(Index) > 0 Then Call SendHP(Index)
                                End If
                        
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Shop = True Then
                                    'Send the shop data to the player
                                    Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).ShopNum)
                                End If
                        
                                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Bank = True Then
                                    'Send the player bank data
                                    Call SendBankInv(Index)
                                End If
                        
                                Moved = YES
                            End If
                        End If
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
        
    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Warp = True Then
        MapNum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).WarpMap
        x = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).WarpX
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).WarpY
                        
        Call PlayerWarp(Index, MapNum, x, y)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).KeyOpen = True Then
        x = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).KeyOpenX
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).KeyOpenY
        
        If Map(GetPlayerMap(Index)).Tile(x, y).Key = True Then
            If TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
            End If
        End If
    End If
    
    ' They tried to hack
    If Moved = NO Then
        Call HackingAttempt(Index, "Position Modification")
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "PlayerMove", Err.Number, Err.Description)
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
'On Error GoTo errorhandler:
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
                
                
                ' Check to make sure that the tile is walkable
                If Map(MapNum).Tile(x, y - 1).Walkable = False Then
                    If Map(MapNum).Tile(x, y - 1).Item = False Then
                        CanNpcMove = False
                        Exit Function
                    End If
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
                
                ' Check to make sure that the tile is walkable
                If Map(MapNum).Tile(x, y + 1).Walkable = False Then
                    If Map(MapNum).Tile(x, y + 1).Item = False Then
                        CanNpcMove = False
                        Exit Function
                    End If
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
                
                ' Check to make sure that the tile is walkable
                If Map(MapNum).Tile(x - 1, y).Walkable = False Then
                    If Map(MapNum).Tile(x - 1, y).Item = False Then
                        CanNpcMove = False
                        Exit Function
                    End If
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
                
                ' Check to make sure that the tile is walkable
                If Map(MapNum).Tile(x + 1, y).Walkable = False Then
                    If Map(MapNum).Tile(x + 1, y).Item = False Then
                        CanNpcMove = False
                        Exit Function
                    End If
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
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "CanNpcMove", Err.Number, Err.Description)
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
'On Error GoTo errorhandler:
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "NpcMove", Err.Number, Err.Description)
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modGameLogic.bas", "NpcDir", Err.Number, Err.Description)
End Sub

Sub JoinGame(ByVal Index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
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
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    'Call SendBankInv(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendStats(Index)
    Call SendPlayers(Index)
    Call SendFriends(Index)
    Call SendWeatherTo(Index)
    Call SendTimeTo(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            
    'Run Public Script
    MyScript.SControl.Modules("main.txt").Run "JoinGame", Index
    
    ' Send welcome messages
    Call SendWelcome(Index)

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(Index, "INGAME" & SEP_CHAR & END_CHAR)
    
    'send equip data
    'Call SendEquipDataTo(Index)
    
    Packet = "PLAYERJOIN" & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & END_CHAR
    Call SendDataToAllBut(Index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "JoinGame", Err.Number, Err.Description)
End Sub

Sub LeftGame(ByVal Index As Long)
'On Error GoTo errorhandler:
On Error Resume Next
Dim Packet As String
Dim n As Long
Dim p As Integer

    If Player(Index).InGame = True Then
        Player(Index).InGame = False
        
        ' Check if player was the only player on the map and stop npc processing if so
        ' Fixed? -smchronos
        If GetTotalMapPlayers(GetPlayerMap(Index)) = 0 Then
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
        'If GetPlayerAccess(Index) <= ADMIN_MONITER Then
        '    Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", JoinLeftColor)
        'Else
        '    Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", White)
        'End If
        
        '.ExecuteStatement "main.txt", "LeaveGame"
        'Run Public Script
        MyScript.SControl.Modules("main.txt").Run "LeaveGame", Index
        Call SendPlayerLeave(Index)
    
        Call TextAdd(frmServer.txtText, GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".", True)
        
        'Remove name from the Player List
        For p = 0 To frmServer.lstPlayers.ListCount
            If frmServer.lstPlayers.List(p) = (GetPlayerLogin(Index) & "/" & GetPlayerName(Index)) Then
                frmServer.lstPlayers.RemoveItem (p)
            End If
        Next p
        
        Call SendLeftGame(Index)
    End If
    
    Call ClearPlayer(Index)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "LeftGame", Err.Number, Err.Description)
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modGameLogic.bas", "GetTotalMapPlayers", Err.Number, Err.Description)
End Function

Function GetNpcMaxHP(ByVal NpcNum As Long)
'On Error GoTo errorhandler:
Dim x As Long, y As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxHP = 0
        Exit Function
    End If
    
    GetNpcMaxHP = Npc(NpcNum).HP
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetNpcMaxHP", Err.Number, Err.Description)
End Function

Function GetNpcMaxMP(ByVal NpcNum As Long)
'On Error GoTo errorhandler:
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxMP = 0
        Exit Function
    End If
        
    GetNpcMaxMP = Npc(NpcNum).MAGI * 2
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetNpcMaxMP", Err.Number, Err.Description)
End Function

Function GetNpcMaxSP(ByVal NpcNum As Long)
'On Error GoTo errorhandler:
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxSP = 0
        Exit Function
    End If
        
    GetNpcMaxSP = Npc(NpcNum).SPEED * 2
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetNpcMaxSP", Err.Number, Err.Description)
End Function

Function GetPlayerHPRegen(ByVal Index As Long)
'On Error GoTo errorhandler:
Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerHPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerDEF(Index) / 2)
    If i < 2 Then i = 2
    
    GetPlayerHPRegen = i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerHPRegen", Err.Number, Err.Description)
End Function

Function GetPlayerMPRegen(ByVal Index As Long)
'On Error GoTo errorhandler:
Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerMPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerMAGI(Index) / 2)
    If i < 2 Then i = 2
    
    GetPlayerMPRegen = i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerMPRegen", Err.Number, Err.Description)
End Function

Function GetPlayerSPRegen(ByVal Index As Long)
'On Error GoTo errorhandler:
Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerSPRegen = 0
        Exit Function
    End If
    
    i = Int(GetPlayerSPEED(Index) / 2)
    If i < 2 Then i = 2
    
    GetPlayerSPRegen = i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerSPRegen", Err.Number, Err.Description)
End Function

Function GetNpcHPRegen(ByVal NpcNum As Long)
'On Error GoTo errorhandler:
Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcHPRegen = 0
        Exit Function
    End If
    
    i = Int(Npc(NpcNum).DEF / 3)
    If i < 1 Then i = 1
    
    GetNpcHPRegen = i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetNpcHPRegen", Err.Number, Err.Description)
End Function

Sub SetAllMapNpcBehavior(ByVal NpcNum As Long)
'On Error GoTo errorhandler:
'Add this to update all map npc behaviors
'-smchronos
Dim i As Long, y As Long, x As Byte
    For i = 1 To MAX_NPCS
        For y = 1 To MAX_MAPS
            For x = 1 To MAX_MAP_NPCS
                If MapNpc(y, x).Num = NpcNum Then
                    If MapNpc(y, x).Behavior <> Npc(NpcNum).Behavior Then
                        MapNpc(y, x).Behavior = Npc(NpcNum).Behavior
                    End If
                End If
            Next x
        Next y
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetAllMapNpcBehavior", Err.Number, Err.Description)
End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
'On Error GoTo errorhandler:
Dim i As Long
Dim Packet As String

    ' Check if attacker got a level up
    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        'Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                    
        'Get the ammount of skill points to add
        'i = Int(GetPlayerLevel(Index) / 10)
        'If i < 1 Then i = 1
        'If i > 3 Then i = 3
            
        'Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + i)
        'Call SetPlayerExp(Index, 0)
        'Call GlobalMsg(GetPlayerName(Index) & " has gained a level!", Brown)
        'Call PlayerMsg(Index, "You have gained a level!  You now have " & GetPlayerPOINTS(Index) & " stat points to distribute.", BrightBlue)
        MyScript.SControl.Modules("main.txt").Run "PlayerLevelUp", Index
    
        Packet = "levelup" & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR
        Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerTNL(Index) & SEP_CHAR & END_CHAR
    
        Call SendDataTo(Index, Packet)
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "CheckPlayerLevelUp", Err.Number, Err.Description)
End Sub

Sub CastSpell(ByVal Index As Long, ByVal SpellSlot As Long)
'On Error GoTo errorhandler:
Dim SpellNum As Long, MPReq As Long, i As Long, n As Long, Damage As Long
Dim Casted As Boolean

    Casted = False
    
    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then
        Call PlayerMsg(Index, "You do not have this spell!", BrightRed)
        Exit Sub
    End If
    
    i = GetSpellReqLevel(Index, SpellNum)
    MPReq = (i + Spell(SpellNum).Data1 + Spell(SpellNum).Data2 + Spell(SpellNum).Data3)
    
    ' Check if they have enough MP
    If GetPlayerMP(Index) < MPReq Then
        Call PlayerMsg(Index, "Not enough mana points!", BrightRed)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If i > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & i & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < Player(Index).AttackTimer + 1000 Then
        Exit Sub
    End If
    
    ' Check if the spell is a give item and do that instead of a stat modification
    If Spell(SpellNum).Type = SPELL_TYPE_GIVEITEM Then
        n = FindOpenInvSlot(Index, Spell(SpellNum).Data1)
        
        If n > 0 Then
            Call GiveItem(Index, Spell(SpellNum).Data1, Spell(SpellNum).Data2)
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & ".", BrightBlue)
            
            ' Take away the mana points
            Call SetPlayerMP(Index, GetPlayerMP(Index) - MPReq)
            Call SendMP(Index)
            Casted = True
        Else
            Call PlayerMsg(Index, "Your inventory is full!", BrightRed)
        End If
        
        Exit Sub
    End If
        
    n = Player(Index).Target
    
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
            If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(n) >= 10 And Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
'                If GetPlayerLevel(n) + 5 >= GetPlayerLevel(Index) Then
'                    If GetPlayerLevel(n) - 5 <= GetPlayerLevel(Index) Then
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                
                        Select Case Spell(SpellNum).Type
                            Case SPELL_TYPE_SUBHP
                        
                                Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(n)
                                If Damage > 0 Then
                                    Call AttackPlayer(Index, n, Damage)
                                Else
                                    Call PlayerMsg(Index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BrightRed)
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
                Call SetPlayerMP(Index, GetPlayerMP(Index) - MPReq)
                Call SendMP(Index)
                Casted = True
            Else
                If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then
                    Select Case Spell(SpellNum).Type
                    
                        Case SPELL_TYPE_ADDHP
                            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerHP(n, GetPlayerHP(n) + Spell(SpellNum).Data1)
                            Call SendHP(n)
                                    
                        Case SPELL_TYPE_ADDMP
                            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerMP(n, GetPlayerMP(n) + Spell(SpellNum).Data1)
                            Call SendMP(n)
                    
                        Case SPELL_TYPE_ADDSP
                            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                            Call SetPlayerSP(n, GetPlayerSP(n) + Spell(SpellNum).Data1)
                            Call SendSP(n)
                    End Select
                    
                    ' Take away the mana points
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - MPReq)
                    Call SendMP(Index)
                    Casted = True
                Else
                    Call PlayerMsg(Index, "Could not cast spell!", BrightRed)
                End If
            End If
        Else
            Call PlayerMsg(Index, "Could not cast spell!", BrightRed)
        End If
    Else
        If MapNpc(GetPlayerMap(Index), n).Behavior <> NPC_BEHAVIOR_FRIENDLY And MapNpc(GetPlayerMap(Index), n).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(Index), n).Num).Name) & ".", BrightBlue)
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_ADDHP
                    MapNpc(GetPlayerMap(Index), n).HP = MapNpc(GetPlayerMap(Index), n).HP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBHP
                    
                    Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(Index), n).Num).DEF / 2)
                    If Damage > 0 Then
                        Call AttackNpc(Index, n, Damage)
                    Else
                        Call PlayerMsg(Index, "The spell was to weak to hurt " & Trim$(Npc(MapNpc(GetPlayerMap(Index), n).Num).Name) & "!", BrightRed)
                    End If
                    
                Case SPELL_TYPE_ADDMP
                    MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBMP
                    MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP - Spell(SpellNum).Data1
            
                Case SPELL_TYPE_ADDSP
                    MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP + Spell(SpellNum).Data1
                
                Case SPELL_TYPE_SUBSP
                    MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP - Spell(SpellNum).Data1
            End Select
        
            ' Take away the mana points
            Call SetPlayerMP(Index, GetPlayerMP(Index) - MPReq)
            Call SendMP(Index)
            Casted = True
        Else
            Call PlayerMsg(Index, "Could not cast spell!", BrightRed)
        End If
    End If

    If Casted = True Then
        Player(Index).AttackTimer = GetTickCount
        Player(Index).CastedSpell = YES
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "CastSpell", Err.Number, Err.Description)
End Sub

Function GetSpellReqLevel(ByVal Index As Long, ByVal SpellNum As Long)
'On Error GoTo errorhandler:
    GetSpellReqLevel = Spell(SpellNum).LevelReq
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetSpellReqLevel", Err.Number, Err.Description)
End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
'On Error GoTo errorhandler:
Dim i As Long, n As Long

    CanPlayerCriticalHit = False
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerSTR(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
    
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "CanPlayerCriticalHit", Err.Number, Err.Description)
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
'On Error GoTo errorhandler:
Dim i As Long, n As Long, ShieldSlot As Long

    CanPlayerBlockHit = False
    
    ShieldSlot = GetPlayerShieldSlot(Index)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
        
            n = Int(Rnd * 100) + 1
            If n <= i Then
                If Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).Data5 <> UNBREAKABLE Then
                    Call SetPlayerInvItemDur(Index, GetPlayerShieldSlot(Index), GetPlayerInvItemDur(Index, GetPlayerShieldSlot(Index)) - 1)
                End If
                CanPlayerBlockHit = True
            End If
        End If
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "CanPlayerBlockHit", Err.Number, Err.Description)
End Function

Function WillArrowSnap(ByVal Index As Long) As Boolean
'On Error GoTo errorhandler:
Dim Chance As Long
    WillArrowSnap = False
    
    If Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).Data4 = SHIELD_TYPE_ARROW Then
        If ItemUnbreakable(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))) Then Exit Function
    Else
        Exit Function
    End If

    'Gets a random number from 1 to 100
    Chance = Int((80 - 1 + 1) * Rnd + 1)
    
    'Checks to see if the chance is higher than the player's speed, if so, break the arrow
    If Chance > GetPlayerSPEED(Index) Then
        WillArrowSnap = True
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "WillArrowSnap", Err.Number, Err.Description)
End Function

Function ItemUnbreakable(ByVal ItemNum As Long) As Boolean
'On Error GoTo errorhandler:
    ItemUnbreakable = False
    If Item(ItemNum).Data5 = UNBREAKABLE Then ItemUnbreakable = True
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "ItemUnbreakable", Err.Number, Err.Description)
End Function

Sub CheckEquippedItems(ByVal Index As Long)
'On Error GoTo errorhandler:
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "CheckEquippedItems", Err.Number, Err.Description)
End Sub

Sub KillData(ByVal Index As Long, ByVal Number As Long, sType As String)
'On Error GoTo errorhandler:
    Call SendDataTo(Index, "KILLDATA" & SEP_CHAR & sType & SEP_CHAR & Number & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "KillData", Err.Number, Err.Description)
End Sub

Sub ClearTempTile()
'On Error GoTo errorhandler:
Dim i As Long, y As Long, x As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                TempTile(i).DoorOpen(x, y) = NO
            Next x
        Next y
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearTempTile", Err.Number, Err.Description)
End Sub

Sub ClearClasses()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 0 To Max_Classes
        Class(i).Name = ""
        Class(i).HP = 0
        Class(i).MP = 0
        Class(i).SP = 0
        Class(i).STR = 0
        Class(i).DEF = 0
        Class(i).SPEED = 0
        Class(i).MAGI = 0
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearClasses", Err.Number, Err.Description)
End Sub

Sub ClearPlayer(ByVal Index As Long)
'On Error GoTo errorhandler:
Dim i As Long
Dim n As Long

    Player(Index).Login = ""
    Player(Index).Password = ""
    
    For i = 1 To MAX_CHARS
        Player(Index).Char(i).Name = ""
        Player(Index).Char(i).Class = 0
        Player(Index).Char(i).Level = 0
        Player(Index).Char(i).Sprite = 0
        Player(Index).Char(i).EXP = 0
        Player(Index).Char(i).Access = 0
        Player(Index).Char(i).PK = NO
        Player(Index).Char(i).POINTS = 0
        Player(Index).Char(i).Guild = 0
        
        Player(Index).Char(i).HP = 0
        Player(Index).Char(i).MP = 0
        Player(Index).Char(i).SP = 0
        
        Player(Index).Char(i).STR = 0
        Player(Index).Char(i).DEF = 0
        Player(Index).Char(i).SPEED = 0
        Player(Index).Char(i).MAGI = 0
        
        Player(Index).Char(i).Text = ""
        
        For n = 1 To MAX_FRIENDS
            Player(Index).Char(i).Friends(n) = ""
        Next n
        
        For n = 1 To MAX_INV
            Player(Index).Char(i).Inv(n).Num = 0
            Player(Index).Char(i).Inv(n).Value = 0
            Player(Index).Char(i).Inv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_BANK_ITEMS
            Player(Index).Char(i).BankInv(n).Num = 0
            Player(Index).Char(i).BankInv(n).Value = 0
            Player(Index).Char(i).BankInv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(i).Spell(n) = 0
        Next n
        
        Player(Index).Char(i).ArmorSlot = 0
        Player(Index).Char(i).WeaponSlot = 0
        Player(Index).Char(i).HelmetSlot = 0
        Player(Index).Char(i).ShieldSlot = 0
        
        Player(Index).Char(i).Map = 0
        Player(Index).Char(i).x = 0
        Player(Index).Char(i).y = 0
        Player(Index).Char(i).Dir = 0
        
        ' Temporary vars
        Player(Index).Buffer = ""
        Player(Index).IncBuffer = ""
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
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearPlayer", Err.Number, Err.Description)
End Sub

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
'On Error GoTo errorhandler:
Dim n As Long
    
    Player(Index).Char(CharNum).Name = ""
    Player(Index).Char(CharNum).Class = 0
    Player(Index).Char(CharNum).Sprite = 0
    Player(Index).Char(CharNum).Level = 0
    Player(Index).Char(CharNum).EXP = 0
    Player(Index).Char(CharNum).Access = 0
    Player(Index).Char(CharNum).PK = NO
    Player(Index).Char(CharNum).POINTS = 0
    Player(Index).Char(CharNum).Guild = 0
    
    Player(Index).Char(CharNum).HP = 0
    Player(Index).Char(CharNum).MP = 0
    Player(Index).Char(CharNum).SP = 0
    
    Player(Index).Char(CharNum).STR = 0
    Player(Index).Char(CharNum).DEF = 0
    Player(Index).Char(CharNum).SPEED = 0
    Player(Index).Char(CharNum).MAGI = 0
    
    For n = 1 To MAX_INV
        Player(Index).Char(CharNum).Inv(n).Num = 0
        Player(Index).Char(CharNum).Inv(n).Value = 0
        Player(Index).Char(CharNum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_BANK_ITEMS
        Player(Index).Char(CharNum).BankInv(n).Num = 0
        Player(Index).Char(CharNum).BankInv(n).Value = 0
        Player(Index).Char(CharNum).BankInv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SPELLS
        Player(Index).Char(CharNum).Spell(n) = 0
    Next n
    
    Player(Index).Char(CharNum).ArmorSlot = 0
    Player(Index).Char(CharNum).WeaponSlot = 0
    Player(Index).Char(CharNum).HelmetSlot = 0
    Player(Index).Char(CharNum).ShieldSlot = 0
    
    Player(Index).Char(CharNum).Map = 0
    Player(Index).Char(CharNum).x = 0
    Player(Index).Char(CharNum).y = 0
    Player(Index).Char(CharNum).Dir = 0
    
    'Debug.Print Player(Index).Char(CharNum).Name
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearChar", Err.Number, Err.Description)
End Sub
    
Sub ClearItem(ByVal Index As Long)
'On Error GoTo errorhandler:
    Item(Index).Name = ""
    
    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearItem", Err.Number, Err.Description)
End Sub

Sub ClearItems()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearItems", Err.Number, Err.Description)
End Sub

Sub ClearNpc(ByVal Index As Long)
'On Error GoTo errorhandler:
    Npc(Index).Name = ""
    Npc(Index).AttackSay = ""
    Npc(Index).Sprite = 0
    Npc(Index).SpawnSecs = 0
    Npc(Index).Behavior = 0
    Npc(Index).Range = 0
    Npc(Index).DropChance = 0
    Npc(Index).DropItem = 0
    Npc(Index).DropItemValue = 0
    Npc(Index).STR = 0
    Npc(Index).DEF = 0
    Npc(Index).SPEED = 0
    Npc(Index).MAGI = 0
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearNpc", Err.Number, Err.Description)
End Sub

Sub ClearNpcs()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearNpcs", Err.Number, Err.Description)
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
    MapItem(MapNum, Index).Num = 0
    MapItem(MapNum, Index).Value = 0
    MapItem(MapNum, Index).Dur = 0
    MapItem(MapNum, Index).x = 0
    MapItem(MapNum, Index).y = 0
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearMapItem", Err.Number, Err.Description)
End Sub

Sub ClearMapItems()
'On Error GoTo errorhandler:
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next x
    Next y
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearMapItems", Err.Number, Err.Description)
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearMapNpc", Err.Number, Err.Description)
End Sub

Sub ClearMapNpcs()
'On Error GoTo errorhandler:
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next x
    Next y
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearMapNpcs", Err.Number, Err.Description)
End Sub

Sub ClearMap(ByVal MapNum As Long)
'On Error GoTo errorhandler:
Dim i As Long
Dim x As Long
Dim y As Long

    Map(MapNum).Name = ""
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).Left = 0
    Map(MapNum).Right = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map(MapNum).Tile(x, y).Ground = 0
            Map(MapNum).Tile(x, y).Mask = 0
            Map(MapNum).Tile(x, y).Mask2 = 0
            Map(MapNum).Tile(x, y).Anim = 0
            Map(MapNum).Tile(x, y).Anim2 = 0
            Map(MapNum).Tile(x, y).Fringe = 0
            Map(MapNum).Tile(x, y).FringeAnim = 0
            Map(MapNum).Tile(x, y).Fringe2 = 0
            Map(MapNum).Tile(x, y).Walkable = True
            Map(MapNum).Tile(x, y).Blocked = False
            Map(MapNum).Tile(x, y).North = False
            Map(MapNum).Tile(x, y).West = False
            Map(MapNum).Tile(x, y).East = False
            Map(MapNum).Tile(x, y).South = False
            Map(MapNum).Tile(x, y).Warp = False
            Map(MapNum).Tile(x, y).WarpMap = 0
            Map(MapNum).Tile(x, y).WarpX = 0
            Map(MapNum).Tile(x, y).WarpY = 0
            Map(MapNum).Tile(x, y).Item = False
            Map(MapNum).Tile(x, y).ItemNum = 0
            Map(MapNum).Tile(x, y).ItemValue = 0
            Map(MapNum).Tile(x, y).NpcAvoid = False
            Map(MapNum).Tile(x, y).Key = False
            Map(MapNum).Tile(x, y).KeyNum = 0
            Map(MapNum).Tile(x, y).KeyTake = 0
            Map(MapNum).Tile(x, y).KeyOpen = False
            Map(MapNum).Tile(x, y).KeyOpenX = 0
            Map(MapNum).Tile(x, y).KeyOpenY = 0
            Map(MapNum).Tile(x, y).Bank = False
            Map(MapNum).Tile(x, y).Shop = False
            Map(MapNum).Tile(x, y).ShopNum = 0
            Map(MapNum).Tile(x, y).Heal = False
            Map(MapNum).Tile(x, y).HealValue = 0
            Map(MapNum).Tile(x, y).Damage = False
            Map(MapNum).Tile(x, y).DamageValue = 0
        Next x
    Next y
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearMap", Err.Number, Err.Description)
End Sub

Sub ClearMaps()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearMaps", Err.Number, Err.Description)
End Sub

Sub ClearShop(ByVal Index As Long)
'On Error GoTo errorhandler:
Dim i As Long

    Shop(Index).Name = ""
    Shop(Index).JoinSay = ""
    Shop(Index).LeaveSay = ""
    
    For i = 1 To MAX_TRADES
        Shop(Index).TradeItem(i).GiveItem = 0
        Shop(Index).TradeItem(i).GiveValue = 0
        Shop(Index).TradeItem(i).GetItem = 0
        Shop(Index).TradeItem(i).GetValue = 0
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearShop", Err.Number, Err.Description)
End Sub

Sub ClearShops()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearShops", Err.Number, Err.Description)
End Sub

Sub ClearSpell(ByVal Index As Long)
'On Error GoTo errorhandler:
    Spell(Index).Name = ""
    Spell(Index).ClassReq = 0
    Spell(Index).LevelReq = 0
    Spell(Index).Type = 0
    Spell(Index).Data1 = 0
    Spell(Index).Data2 = 0
    Spell(Index).Data3 = 0
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearSpell", Err.Number, Err.Description)
End Sub

Sub ClearSpells()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "ClearSpells", Err.Number, Err.Description)
End Sub




' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal Index As Long) As String
'On Error GoTo errorhandler:
    GetPlayerLogin = Trim$(Player(Index).Login)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerLogin", Err.Number, Err.Description)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
'On Error GoTo errorhandler:
    Player(Index).Login = Login
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerLogin", Err.Number, Err.Description)
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
'On Error GoTo errorhandler:
    GetPlayerPassword = Trim$(Player(Index).Password)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerPassword", Err.Number, Err.Description)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
'On Error GoTo errorhandler:
    Player(Index).Password = Password
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerPassword", Err.Number, Err.Description)
End Sub

Function GetPlayerName(ByVal Index As Long) As String
'On Error GoTo errorhandler:
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerName", Err.Number, Err.Description)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Name = Name
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerName", Err.Number, Err.Description)
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerClass", Err.Number, Err.Description)
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerClass", Err.Number, Err.Description)
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerSprite", Err.Number, Err.Description)
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerSprite", Err.Number, Err.Description)
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).Level
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerLevel", Err.Number, Err.Description)
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Level = Level
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerLevel", Err.Number, Err.Description)
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    If GetPlayerLevel(Index) > MAX_EXPERIENCE Then
        GetPlayerNextLevel = Experience(MAX_EXPERIENCE)
    Else
        GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerNextLevel", Err.Number, Err.Description)
End Function

Function GetPlayerTNL(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerTNL = GetPlayerNextLevel(Index) - GetPlayerExp(Index)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerTNL", Err.Number, Err.Description)
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).EXP
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerExp", Err.Number, Err.Description)
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).EXP = EXP
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerExp", Err.Number, Err.Description)
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerAccess", Err.Number, Err.Description)
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Access = Access
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerAccess", Err.Number, Err.Description)
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerPK", Err.Number, Err.Description)
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).PK = PK
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerPK", Err.Number, Err.Description)
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerHP", Err.Number, Err.Description)
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).HP = HP
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If
    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerHP", Err.Number, Err.Description)
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerMP", Err.Number, Err.Description)
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If
    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerMP", Err.Number, Err.Description)
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerSP", Err.Number, Err.Description)
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If
    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerSP", Err.Number, Err.Description)
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
Dim CharNum As Long
Dim i As Long

    CharNum = Player(Index).CharNum
    GetPlayerMaxHP = (Player(Index).Char(CharNum).Level + Int(GetPlayerSTR(Index) / 2) + Class(Player(Index).Char(CharNum).Class).HP)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerMaxHP", Err.Number, Err.Description)
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
Dim CharNum As Long

    CharNum = Player(Index).CharNum
    GetPlayerMaxMP = (Player(Index).Char(CharNum).Level + Int(GetPlayerMAGI(Index) / 2) + Class(Player(Index).Char(CharNum).Class).MP)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerMaxMP", Err.Number, Err.Description)
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
Dim CharNum As Long

    CharNum = Player(Index).CharNum
    GetPlayerMaxSP = (Player(Index).Char(CharNum).Level + Int(GetPlayerSPEED(Index) / 2) + Class(Player(Index).Char(CharNum).Class).SP)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerMaxSP", Err.Number, Err.Description)
End Function

Function GetClassName(ByVal ClassNum As Long) As String
'On Error GoTo errorhandler:
    GetClassName = Trim$(Class(ClassNum).Name)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerClassName", Err.Number, Err.Description)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
'On Error GoTo errorhandler:
    'GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
    GetClassMaxHP = Class(ClassNum).HP
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetClassMaxHP", Err.Number, Err.Description)
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
'On Error GoTo errorhandler:
    'GetClassMaxMP = (1 + Int(Class(ClassNum).MAGI / 2) + Class(ClassNum).MAGI) * 2
    GetClassMaxMP = Class(ClassNum).MP
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetClassMaxMP", Err.Number, Err.Description)
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
'On Error GoTo errorhandler:
    'GetClassMaxSP = (1 + Int(Class(ClassNum).SPEED / 2) + Class(ClassNum).SPEED) * 2
    GetClassMaxSP = Class(ClassNum).SP
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetClassMaxSP", Err.Number, Err.Description)
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
'On Error GoTo errorhandler:
    GetClassSTR = Class(ClassNum).STR
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetClassSTR", Err.Number, Err.Description)
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
'On Error GoTo errorhandler:
    GetClassDEF = Class(ClassNum).DEF
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetClassDEF", Err.Number, Err.Description)
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
'On Error GoTo errorhandler:
    GetClassSPEED = Class(ClassNum).SPEED
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetClassSPEED", Err.Number, Err.Description)
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
'On Error GoTo errorhandler:
    GetClassMAGI = Class(ClassNum).MAGI
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetClassMAGI", Err.Number, Err.Description)
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerSTR", Err.Number, Err.Description)
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).STR = STR
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerSTR", Err.Number, Err.Description)
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerDEF", Err.Number, Err.Description)
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerDEF", Err.Number, Err.Description)
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).SPEED
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerSPEED", Err.Number, Err.Description)
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal SPEED As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).SPEED = SPEED
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerSPEED", Err.Number, Err.Description)
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).MAGI
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerMAGI", Err.Number, Err.Description)
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).MAGI = MAGI
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerMAGI", Err.Number, Err.Description)
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerPOINTS", Err.Number, Err.Description)
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerPOINTS", Err.Number, Err.Description)
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerMap", Err.Number, Err.Description)
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = MapNum
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerMap", Err.Number, Err.Description)
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).x
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerX", Err.Number, Err.Description)
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).x = x
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerX", Err.Number, Err.Description)
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).y
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerY", Err.Number, Err.Description)
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).y = y
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerY", Err.Number, Err.Description)
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerDir", Err.Number, Err.Description)
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerDir", Err.Number, Err.Description)
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
'On Error GoTo errorhandler:
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerIP", Err.Number, Err.Description)
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerInvItemNum", Err.Number, Err.Description)
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num = ItemNum
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerInvItemNum", Err.Number, Err.Description)
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerInvItemValue", Err.Number, Err.Description)
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerInvItemValue", Err.Number, Err.Description)
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerInvItemDur", Err.Number, Err.Description)
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerInvItemDur", Err.Number, Err.Description)
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerBankItemNum = Player(Index).Char(Player(Index).CharNum).BankInv(InvSlot).Num
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerBankItemNum", Err.Number, Err.Description)
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).BankInv(InvSlot).Num = ItemNum
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerBankItemNum", Err.Number, Err.Description)
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerBankItemValue = Player(Index).Char(Player(Index).CharNum).BankInv(InvSlot).Value
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerBankItemValue", Err.Number, Err.Description)
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).BankInv(InvSlot).Value = ItemValue
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerBankItemValue", Err.Number, Err.Description)
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerBankItemDur = Player(Index).Char(Player(Index).CharNum).BankInv(InvSlot).Dur
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerBankItemDur", Err.Number, Err.Description)
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).BankInv(InvSlot).Dur = ItemDur
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerBankItemDur", Err.Number, Err.Description)
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerSpell", Err.Number, Err.Description)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = SpellNum
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerSpell", Err.Number, Err.Description)
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerArmorSlot", Err.Number, Err.Description)
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerArmorSlot", Err.Number, Err.Description)
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerWeaponSlot", Err.Number, Err.Description)
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerWeaponSlot", Err.Number, Err.Description)
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerHelmetSlot", Err.Number, Err.Description)
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerHelmetSlot", Err.Number, Err.Description)
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
'On Error GoTo errorhandler:
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetPlayerShieldSlot", Err.Number, Err.Description)
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
'On Error GoTo errorhandler:
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetPlayerShieldSlot", Err.Number, Err.Description)
End Sub

Sub SetBehavior(ByVal Map As Long, ByVal MapNpcNum As Byte, ByVal Behavior As Byte)
'On Error GoTo errorhandler:
'Sets the behavior of one map npc
'-smchronos
    MapNpc(Map, MapNpcNum).Behavior = Behavior
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetBehavior", Err.Number, Err.Description)
End Sub

Function GetBehavior(ByVal Map As Long, ByVal MapNpcNum As Byte)
'On Error GoTo errorhandler:
'Gets the behavior of one map npc
'-smchronos
    GetBehavior = MapNpc(Map, MapNpcNum).Behavior
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetBehavior", Err.Number, Err.Description)
End Function

Function GetNpcTarget(ByVal Map As Long, ByVal MapNpcNum As Byte)
'On Error GoTo errorhandler:
'Gets the target of one map npc
'-smchronos
    GetNpcTarget = MapNpc(Map, MapNpcNum).Target
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modGameLogic.bas", "GetNpcTarget", Err.Number, Err.Description)
End Function

Sub SetNpcTarget(ByVal Map As Long, ByVal MapNpcNum As Byte, ByVal Target As Long)
'On Error GoTo errorhandler:
'Sets the target of one map npc
'-smchronos
    MapNpc(Map, MapNpcNum).Target = Target
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGameLogic.bas", "SetNpcTarget", Err.Number, Err.Description)
End Sub

