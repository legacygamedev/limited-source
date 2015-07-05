Attribute VB_Name = "modPlayerLogic"
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


Function GetPlayerDamage(ByVal Index As Long) As Long
Dim WeaponSlot As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 = WEAPON_SUBTYPE_BOW Then
            GetPlayerDamage = Int(GetPlayerDEX(Index) / 3) + Int(GetPlayerSTR(Index) / 5)
        Else
            GetPlayerDamage = Int(GetPlayerSTR(Index) / 2)
        End If
    Else
        GetPlayerDamage = Int(GetPlayerSTR(Index) / 2)
    End If
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(Index)
        
        'GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2
        
        If Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data3 <> WEAPON_SUBTYPE_BOW Then
            GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2
            Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) - 1)
            Call SendInventoryUpdate(Index, WeaponSlot)
            Call AttributeSkillsExp(Index, SKILL_TYPE_ATTRIBUTE, SKILL_ATTRIBUTE_STR, 1, True)
        Else
            'Call AttributeSkillsExp(Index, SKILL_TYPE_ATTRIBUTE, SKILL_ATTRIBUTE_STR, 1, True)
            GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index))).Data2
            Call AttributeSkillsExp(Index, SKILL_TYPE_ATTRIBUTE, SKILL_ATTRIBUTE_DEX, 1, False)
            Call AttributeSkillsExp(Index, SKILL_TYPE_CHANCE, SKILL_CHANCE_ACCU, 1, False)
        End If
        
        If GetPlayerInvItemDur(Index, WeaponSlot) <= 0 Then
            Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " has Broken" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
            Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
        Else
            If GetPlayerInvItemDur(Index, WeaponSlot) < 5 Then
                Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " Breaking!" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
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
        Call SetPlayerInvItemDur(Index, ArmorSlot, GetPlayerInvItemDur(Index, ArmorSlot) - 1)
        Call SendInventoryUpdate(Index, ArmorSlot)
        Call AttributeSkillsExp(Index, SKILL_TYPE_ATTRIBUTE, SKILL_ATTRIBUTE_DEF, 1, False)
        
        If GetPlayerInvItemDur(Index, ArmorSlot) <= 0 Then
            Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " has Broken" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
            Call TakeItem(Index, GetPlayerInvItemNum(Index, ArmorSlot), 0)
        Else
            If GetPlayerInvItemDur(Index, ArmorSlot) <= 5 Then
                Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " Breaking!" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, HelmSlot)).Data2
        Call SetPlayerInvItemDur(Index, HelmSlot, GetPlayerInvItemDur(Index, HelmSlot) - 1)
        Call SendInventoryUpdate(Index, HelmSlot)
        Call AttributeSkillsExp(Index, SKILL_TYPE_ATTRIBUTE, SKILL_ATTRIBUTE_DEF, 1, False)
        
        If GetPlayerInvItemDur(Index, HelmSlot) <= 0 Then
            Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " has Broken" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
            Call TakeItem(Index, GetPlayerInvItemNum(Index, HelmSlot), 0)
        Else
            If GetPlayerInvItemDur(Index, HelmSlot) <= 5 Then
                Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " Breaking!" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
End Function

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
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
End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenSkillSlot(ByVal Index As Long) As Long
Dim i As Long

    FindOpenSkillSlot = 0
    
    For i = 1 To MAX_PLAYER_SKILLS
        If GetPlayerSkill(Index, i) = 0 Then
            FindOpenSkillSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenQuestSlot(ByVal Index As Long) As Long
Dim i As Long

    FindOpenQuestSlot = 0
    
    For i = 1 To MAX_PLAYER_QUESTS
        If GetPlayerQuest(Index, i) = 0 Then
            FindOpenQuestSlot = i
            Exit Function
        End If
    Next i
End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next i
End Function

Function HasSkill(ByVal Index As Long, ByVal SkillNum As Long) As Boolean
Dim i As Long

    HasSkill = False
    
    For i = 1 To MAX_PLAYER_SKILLS
        If GetPlayerSkill(Index, i) = SkillNum Then
            HasSkill = True
            Exit Function
        End If
    Next i
End Function

Sub AttributeSkillsExp(ByVal Index As Long, ByVal SkillType As Long, ByVal SubType As Long, ByVal EXP As Long, Optional Weapon As Boolean = False)
Dim i As Long
Dim SkillNum As Long
Dim WeaponType As Long

    If Weapon = False Then
        For i = 1 To MAX_PLAYER_SKILLS
            SkillNum = GetPlayerSkill(Index, i)
            If SkillNum > 0 Then
                If Skill(SkillNum).Type = SkillType And Skill(SkillNum).Data1 = SubType Then
                    Call SetPlayerSkillExp(Index, i, (GetPlayerSkillExp(Index, i) + EXP))
                    'Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Trim(Skill(SkillNum).Name) & " Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                    'Call SendPlayerSkillsExp(Index, I)
                    Call CheckSkillLevelUp(Index, i)
                End If
            End If
        Next i
    Else
        WeaponType = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3
        For i = 1 To MAX_PLAYER_SKILLS
            SkillNum = GetPlayerSkill(Index, i)
            If SkillNum > 0 Then
                If Skill(SkillNum).Type = SkillType And Skill(SkillNum).Data1 = SubType Then
                    If Skill(SkillNum).Data3 = WeaponType Then
                        Call SetPlayerSkillExp(Index, i, (GetPlayerSkillExp(Index, i) + EXP))
                        'Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Skill(SkillNum).Name) & " Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                        'Call SendPlayerSkillsExp(Index, I)
                        Call CheckSkillLevelUp(Index, i)
                    End If
                End If
            End If
        Next i
    End If
End Sub

Sub PlayerMapGetItem(ByVal Index As Long)
Dim i As Long
Dim n As Long
Dim f As Long
Dim MapNum As Long
Dim Msg As String

    If IsPlaying(Index) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Index)
    
    For i = MAX_MAP_ITEMS To 1 Step -1
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).x = GetPlayerX(Index)) And (MapItem(MapNum, i).y = GetPlayerY(Index)) Then
                ' Find open slot
                n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
                
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventory
                    Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
                    If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Then
                        Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, i).Value)
                        Msg = "Picked up " & MapItem(MapNum, i).Value & " " & Trim(Item(GetPlayerInvItemNum(Index, n)).Name)
                    Else
                        Call SetPlayerInvItemValue(Index, n, 0)
                        Msg = "Picked up: " & Trim(Item(GetPlayerInvItemNum(Index, n)).Name)
                    End If
                    Call SetPlayerInvItemDur(Index, n, MapItem(MapNum, i).Dur)
                    
                    ' Check for quest item
                    For f = 1 To MAX_PLAYER_QUESTS
                        If Player(Index).Char(Player(Index).CharNum).Quests(f).Num > 0 Then
                            If Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data1 = MapItem(MapNum, i).Num Then
                                If Player(Index).Char(Player(Index).CharNum).Quests(f).Count < Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2 Then
                                    If Item(MapItem(MapNum, i).Num).Type <> ITEM_TYPE_CURRENCY Then
                                        Player(Index).Char(Player(Index).CharNum).Quests(f).Count = (Player(Index).Char(Player(Index).CharNum).Quests(f).Count + 1)
                                    Else
                                        Player(Index).Char(Player(Index).CharNum).Quests(f).Count = GetPlayerInvItemValue(Index, n)
                                    End If
                                    If Player(Index).Char(Player(Index).CharNum).Quests(f).Count > Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2 Then
                                        Player(Index).Char(Player(Index).CharNum).Quests(f).Count = Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2
                                    End If
                                    Call UpdatePlayerQuest(Index, f)
                                End If
                            End If
                        End If
                    Next f
                        
                    ' Erase item from the map
                    MapItem(MapNum, i).Num = 0
                    MapItem(MapNum, i).Value = 0
                    MapItem(MapNum, i).Dur = 0
                    MapItem(MapNum, i).x = 0
                    MapItem(MapNum, i).y = 0
                        
                    Call SendInventoryUpdate(Index, n)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                    Exit Sub
                Else
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Inventory Full" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                   Exit Sub
                End If
            End If
        End If
    Next i
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Ammount As Long)
Dim MapNum As Long
Dim i As Long
Dim f As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Index)
    
    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(MapNum)
        
        If i <> 0 Then
            MapItem(MapNum, i).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(Index) Then
                        Call SetPlayerArmorSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(MapNum, i).Dur = GetPlayerInvItemDur(Index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(Index) Then
                        Call SetPlayerWeaponSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(MapNum, i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(Index) Then
                        Call SetPlayerHelmetSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(MapNum, i).Dur = GetPlayerInvItemDur(Index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(Index) Then
                        Call SetPlayerShieldSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(MapNum, i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_TOOL
                    If InvNum = GetPlayerWeaponSlot(Index) Then
                        Call SetPlayerWeaponSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(MapNum, i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_AMULET
                    If InvNum = GetPlayerAmuletSlot(Index) Then
                        Call SetPlayerAmuletSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    'MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_RING
                    If InvNum = GetPlayerRingSlot(Index) Then
                        Call SetPlayerRingSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    'MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_ARROW
                    If InvNum = GetPlayerArrowSlot(Index) Then
                        Call SetPlayerArrowSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(MapNum, i).Dur = GetPlayerInvItemDur(Index, InvNum)
            End Select
                                
            MapItem(MapNum, i).Num = GetPlayerInvItemNum(Index, InvNum)
            MapItem(MapNum, i).x = GetPlayerX(Index)
            MapItem(MapNum, i).y = GetPlayerY(Index)
                        
            If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Ammount >= GetPlayerInvItemValue(Index, InvNum) Then
                    MapItem(MapNum, i).Value = GetPlayerInvItemValue(Index, InvNum)
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                    ' Check for quest utem dropped
                    For f = 1 To MAX_PLAYER_QUESTS
                        If Player(Index).Char(Player(Index).CharNum).Quests(f).Num > 0 Then
                            If Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data1 = MapItem(MapNum, i).Num Then
                                Player(Index).Char(Player(Index).CharNum).Quests(f).Count = GetPlayerInvItemValue(Index, InvNum)
                                Call UpdatePlayerQuest(Index, f)
                            End If
                        End If
                    Next f
                Else
                    MapItem(MapNum, i).Value = Ammount
                    Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Ammount)
                    ' Check for quest item reduction
                    For f = 1 To MAX_PLAYER_QUESTS
                        If Player(Index).Char(Player(Index).CharNum).Quests(f).Num > 0 Then
                            If Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data1 = GetPlayerInvItemNum(Index, InvNum) Then
                                If Player(Index).Char(Player(Index).CharNum).Quests(f).Count <= Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2 Then
                                    Player(Index).Char(Player(Index).CharNum).Quests(f).Count = GetPlayerInvItemValue(Index, InvNum)
                                    If Player(Index).Char(Player(Index).CharNum).Quests(f).Count > Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2 Then
                                        Player(Index).Char(Player(Index).CharNum).Quests(f).Count = Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2
                                    End If
                                    Call UpdatePlayerQuest(Index, f)
                                End If
                            End If
                        End If
                    Next f
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(MapNum, i).Value = 0
                
                Call SetPlayerInvItemNum(Index, InvNum, 0)
                Call SetPlayerInvItemValue(Index, InvNum, 0)
                Call SetPlayerInvItemDur(Index, InvNum, 0)
                ' Check for quest item dropped
                For f = 1 To MAX_PLAYER_QUESTS
                    If Player(Index).Char(Player(Index).CharNum).Quests(f).Num > 0 Then
                        If Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data1 = MapItem(MapNum, i).Num Then
                            Player(Index).Char(Player(Index).CharNum).Quests(f).Count = GetPlayerInvItemValue(Index, InvNum)
                            Call UpdatePlayerQuest(Index, f)
                        End If
                    End If
                Next f
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(Index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(MapNum, i).Num, Ammount, MapItem(MapNum, i).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            
            ' Check for quest item - needs adding to item type check
            'For f = 1 To MAX_PLAYER_QUESTS
                'If Player(Index).Char(Player(Index).CharNum).Quests(f).Num > 0 Then
                    'If Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data1 = GetPlayerInvItemNum(Index, InvNum) Then
                        'If Player(Index).Char(Player(Index).CharNum).Quests(f).Count < Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2 Then
                            'Player(Index).Char(Player(Index).CharNum).Quests(f).Count = GetPlayerInvItemValue(Index, InvNum)
                            'If Player(Index).Char(Player(Index).CharNum).Quests(f).Count > Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2 Then
                                'Player(Index).Char(Player(Index).CharNum).Quests(f).Count = Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2
                            'End If
                            'Call UpdatePlayerQuest(Index, f)
                        'End If
                    'End If
                'End If
            'Next f
        Else
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "No More Room" & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
        End If
    End If
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
                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Admin Don't PK!" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Cannot Attack Admin" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_PLAYER Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_GUILD Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Must Be Level 10" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Player Protected" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Safe Zone" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                            End If
                        End If
                    End If
                End If
            
            Case DIR_DOWN
                If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Admin Don't PK!" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Cannot Attack Admin" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_PLAYER Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_GUILD Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Must Be Level 10" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Player Protected" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Safe Zone" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                            End If
                        End If
                    End If
                End If
        
            Case DIR_LEFT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Admin Don't PK!" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Cannot Attack Admin" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_PLAYER Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_GUILD Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Must Be Level 10" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Player Protected" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Safe Zone" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                            End If
                        End If
                    End If
                End If
            
            Case DIR_RIGHT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                    ' Check to make sure that they dont have access
                    If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Admin Don't PK!" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Else
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Cannot Attack Admin" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_PLAYER Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_GUILD Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < 10 Then
                                    Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Must Be Level 10" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Else
                                    If GetPlayerLevel(Victim) < 10 Then
                                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Player Protected" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            Else
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Safe Zone" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                            End If
                        End If
                    End If
                End If
        End Select
    End If
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
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
                        If Npc(NpcNum).HitOnlyWith > 0 Then
                            If GetPlayerWeaponSlot(Attacker) = 0 Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Equip Weapon" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                            If GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) <> Npc(NpcNum).HitOnlyWith Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Weapon" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                        End If
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                                Call SendTrade(Attacker, Npc(NpcNum).ShopLink)
                            Else
                                Call SendNpcQuests(Attacker, NpcNum)
                                'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
                            End If
                        End If
                    End If
                
                Case DIR_DOWN
                    If (MapNpc(MapNum, MapNpcNum).y - 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).HitOnlyWith > 0 Then
                            If GetPlayerWeaponSlot(Attacker) = 0 Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Equip Weapon" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                            If GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) <> Npc(NpcNum).HitOnlyWith Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Weapon" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                        End If
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                                Call SendTrade(Attacker, Npc(NpcNum).ShopLink)
                            Else
                                Call SendNpcQuests(Attacker, NpcNum)
                                'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
                            End If
                        End If
                    End If
                
                Case DIR_LEFT
                    If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x + 1 = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).HitOnlyWith > 0 Then
                            If GetPlayerWeaponSlot(Attacker) = 0 Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Equip Weapon" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                            If GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) <> Npc(NpcNum).HitOnlyWith Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Weapon" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                        End If
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                                Call SendTrade(Attacker, Npc(NpcNum).ShopLink)
                            Else
                                Call SendNpcQuests(Attacker, NpcNum)
                                'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
                            End If
                        End If
                    End If
                
                Case DIR_RIGHT
                    If (MapNpc(MapNum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNpc(MapNum, MapNpcNum).x - 1 = GetPlayerX(Attacker)) Then
                        If Npc(NpcNum).HitOnlyWith > 0 Then
                            If GetPlayerWeaponSlot(Attacker) = 0 Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Equip Weapon" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                            If GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) <> Npc(NpcNum).HitOnlyWith Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Weapon" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                        End If
                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            CanAttackNpc = True
                        Else
                            If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                                Call SendTrade(Attacker, Npc(NpcNum).ShopLink)
                            Else
                                Call SendNpcQuests(Attacker, NpcNum)
                                'Call PlayerMsg(Attacker, "You cannot attack a " & Trim(Npc(NpcNum).Name) & "!", BrightBlue)
                            End If
                        End If
                    End If
            End Select
        End If
    End If
End Function

Function CanAttackResource(ByVal Attacker As Long, ByVal MapResourceNum As Long) As Boolean
Dim MapNum As Long, ResourceNum As Long

    CanAttackResource = False
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapResourceNum <= 0 Or MapResourceNum > MAX_MAP_RESOURCES Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapResource(GetPlayerMap(Attacker), MapResourceNum).Num <= 0 Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Attacker)
    ResourceNum = MapResource(MapNum, MapResourceNum).Num
    
    ' Make sure the npc isn't already dead
    If MapResource(MapNum, MapResourceNum).HP <= 0 Then
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If ResourceNum > 0 And GetTickCount > Player(Attacker).AttackTimer + 950 Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If (MapResource(MapNum, MapResourceNum).y + 1 = GetPlayerY(Attacker)) And (MapResource(MapNum, MapResourceNum).x = GetPlayerX(Attacker)) Then
                        If Npc(ResourceNum).HitOnlyWith > 0 Then
                            If GetPlayerWeaponSlot(Attacker) = 0 Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Equip Tool" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                            If GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) <> Npc(ResourceNum).HitOnlyWith Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Tool" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                        End If
                    CanAttackResource = True
                    End If
                
                Case DIR_DOWN
                    If (MapResource(MapNum, MapResourceNum).y - 1 = GetPlayerY(Attacker)) And (MapResource(MapNum, MapResourceNum).x = GetPlayerX(Attacker)) Then
                        If Npc(ResourceNum).HitOnlyWith > 0 Then
                            If GetPlayerWeaponSlot(Attacker) = 0 Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Equip Tool" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                            If GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) <> Npc(ResourceNum).HitOnlyWith Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Tool" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                        End If
                    CanAttackResource = True
                    End If
                
                Case DIR_LEFT
                    If (MapResource(MapNum, MapResourceNum).y = GetPlayerY(Attacker)) And (MapResource(MapNum, MapResourceNum).x + 1 = GetPlayerX(Attacker)) Then
                        If Npc(ResourceNum).HitOnlyWith > 0 Then
                            If GetPlayerWeaponSlot(Attacker) = 0 Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Equip Tool" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                            If GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) <> Npc(ResourceNum).HitOnlyWith Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Tool" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                        End If
                    CanAttackResource = True
                    End If
                
                Case DIR_RIGHT
                    If (MapResource(MapNum, MapResourceNum).y = GetPlayerY(Attacker)) And (MapResource(MapNum, MapResourceNum).x - 1 = GetPlayerX(Attacker)) Then
                        If Npc(ResourceNum).HitOnlyWith > 0 Then
                            If GetPlayerWeaponSlot(Attacker) = 0 Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Equip Tool" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                            If GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker)) <> Npc(ResourceNum).HitOnlyWith Then
                                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Tool" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                Exit Function
                            End If
                        End If
                    CanAttackResource = True
                    End If
            End Select
        End If
    End If
End Function

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
Dim Name As String
Dim EXP As Long
Dim ExpType As Long
Dim ExpSentAttacker As Byte
Dim ExpSentParty As Byte
Dim ExpSentPartyPlayer As Byte
Dim n As Long, i As Long, J As Long, K As Long
Dim STR As Long, DEF As Long, MapNum As Long, NpcNum As Long
Dim AmuletSlot As Long
Dim RingSlot As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    Name = Trim(Npc(NpcNum).Name)
    ExpType = Npc(NpcNum).ExpType
    ExpSentAttacker = 0
    ExpSentParty = 0
    ExpSentPartyPlayer = 0
        
    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        Call SendDataTo(Attacker, "BLITWARNMSG" & SEP_CHAR & " Battle Won" & SEP_CHAR & Brown & SEP_CHAR & END_CHAR)
                        
        ' Calculate exp to give attacker
        STR = Npc(NpcNum).STR
        DEF = Npc(NpcNum).DEF
        EXP = Npc(NpcNum).EXP
        
        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 1
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If Player(Attacker).InParty = NO Then
            ' Check for non standard experience
            If ExpType <= 0 Then
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "  Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                Call SendStats(Attacker)
                ExpSentAttacker = 1
            Else
                ' Determine the skill type to give EXP to
                For i = 1 To MAX_PLAYER_SKILLS
                    If GetPlayerSkill(Attacker, i) = ExpType Then
                        Call SetPlayerSkillExp(Attacker, i, GetPlayerSkillExp(Attacker, i) + EXP)
                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & Trim(Skill(GetPlayerSkill(Attacker, i)).Name) & " Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                        'Call SendPlayerSkillsExp(Attacker, I)
                        Call CheckSkillLevelUp(Attacker, i)
                        ExpSentAttacker = 1
                    End If
                Next i
            End If
            If ExpSentAttacker = 0 Then
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "  Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                Call SendStats(Attacker)
                ExpSentAttacker = 1
            End If
        Else
            EXP = EXP / 2
            
            If EXP < 0 Then
                EXP = 1
            End If
            
            ' Check for non standard experience
            If ExpType <= 0 Then
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Party Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                Call SendStats(Attacker)
                ExpSentParty = 1
            Else
                ' Determine the skill type to give EXP to
                For i = 1 To MAX_PLAYER_SKILLS
                    If GetPlayerSkill(Attacker, i) = ExpType Then
                        Call SetPlayerSkillExp(Attacker, i, GetPlayerSkillExp(Attacker, i) + EXP)
                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & Trim(Skill(GetPlayerSkill(Attacker, i)).Name) & " Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                        'Call SendPlayerSkillsExp(Attacker, I)
                        Call CheckSkillLevelUp(Attacker, i)
                        ExpSentParty = 1
                    End If
                Next i
            End If
            If ExpSentParty = 0 Then
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Party Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                Call SendStats(Attacker)
                ExpSentParty = 1
            End If
            
            n = Player(Attacker).PartyPlayer
            If n > 0 Then
                ' Check for non standard experience
                If ExpType <= 0 Then
                    Call SetPlayerExp(n, GetPlayerExp(n) + EXP)
                    Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & "Party Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                    Call SendStats(n)
                    ExpSentPartyPlayer = 1
                Else
                    ' Determine the skill type to give EXP to
                    For i = 1 To MAX_PLAYER_SKILLS
                        If GetPlayerSkill(n, i) = ExpType Then
                            Call SetPlayerSkillExp(n, i, GetPlayerSkillExp(n, i) + EXP)
                            Call SendDataTo(n, "BLITWARNMSG" & SEP_CHAR & "Party Experience" & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                            Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & Trim(Skill(GetPlayerSkill(n, i)).Name) & " Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                            'Call SendPlayerSkillsExp(n, I)
                            Call CheckSkillLevelUp(n, i)
                            ExpSentPartyPlayer = 1
                        End If
                    Next i
                End If
                If ExpSentPartyPlayer = 0 Then
                    Call SetPlayerExp(n, GetPlayerExp(n) + EXP)
                    Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & "Party Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                    Call SendStats(n)
                    ExpSentPartyPlayer = 1
                End If
            End If
        End If
        
        For i = 1 To MAX_PLAYER_QUESTS
            If Player(Attacker).Char(Player(Attacker).CharNum).Quests(i).Num > 0 Then
                If Quest(Player(Attacker).Char(Player(Attacker).CharNum).Quests(i).Num).Data1 = NpcNum Then
                    If Player(Attacker).Char(Player(Attacker).CharNum).Quests(i).Count < Quest(Player(Attacker).Char(Player(Attacker).CharNum).Quests(i).Num).Data2 Then
                        Player(Attacker).Char(Player(Attacker).CharNum).Quests(i).Count = Player(Attacker).Char(Player(Attacker).CharNum).Quests(i).Count + 1
                        Call UpdatePlayerQuest(Attacker, i)
                    End If
                End If
            End If
        Next i
        
        For i = 1 To MAX_NPC_DROPS
            If Npc(NpcNum).ItemNPC(i).ItemNum > 0 Then
                ' Drop the goods if they get it
                J = Npc(NpcNum).ItemNPC(i).Chance
                AmuletSlot = GetPlayerAmuletSlot(Attacker)
                RingSlot = GetPlayerRingSlot(Attacker)
                For K = 1 To MAX_PLAYER_SKILLS
                    If GetPlayerSkill(Attacker, K) > 0 Then
                        If (Skill(GetPlayerSkill(Attacker, K)).Type = SKILL_TYPE_CHANCE) And (Skill(GetPlayerSkill(Attacker, K)).Data1 = SKILL_CHANCE_DROP) Then
                            J = J - Int((J / 100) * (Skill(GetPlayerSkill(Attacker, K)).Data2 * GetPlayerSkillLevel(Attacker, K)))
                        End If
                    End If
                Next K
                If AmuletSlot > 0 Then
                    If Item(GetPlayerInvItemNum(Attacker, AmuletSlot)).Data1 = CHARM_TYPE_ADDDROP Then
                        J = J - Int((J / 100) * Item(GetPlayerInvItemNum(Attacker, AmuletSlot)).Data2)
                    End If
                End If
                If RingSlot > 0 Then
                    If Item(GetPlayerInvItemNum(Attacker, RingSlot)).Data1 = CHARM_TYPE_ADDDROP Then
                        J = J - Int((J / 100) * Item(GetPlayerInvItemNum(Attacker, RingSlot)).Data2)
                    End If
                End If
                n = Int(Rnd * J) + 1
                If n = 1 Then
                    Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
                    Call AttributeSkillsExp(Attacker, SKILL_TYPE_CHANCE, SKILL_CHANCE_DROP, Npc(NpcNum).ItemNPC(i).Chance, False)
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

Sub AttackResource(ByVal Attacker As Long, ByVal MapResourceNum As Long, ByVal Damage As Long)
Dim Name As String
Dim EXP As Long
Dim ExpType As Long
Dim ExpSentAttacker As Byte
Dim ExpSentParty As Byte
Dim ExpSentPartyPlayer As Byte
Dim n As Long, i As Long, J As Long, K As Long
Dim STR As Long, DEF As Long, MapNum As Long, ResourceNum As Long
Dim AmuletSlot As Long
Dim RingSlot As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapResourceNum <= 0 Or MapResourceNum > MAX_MAP_RESOURCES Or Damage < 0 Then
        Exit Sub
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)
    
    MapNum = GetPlayerMap(Attacker)
    ResourceNum = MapResource(MapNum, MapResourceNum).Num
    Name = Trim(Npc(ResourceNum).Name)
    ExpSentAttacker = 0
    ExpSentParty = 0
    ExpSentPartyPlayer = 0
        
    If Damage >= MapResource(MapNum, MapResourceNum).HP Then
        Call SendDataTo(Attacker, "BLITWARNMSG" & SEP_CHAR & "Resource Processed" & SEP_CHAR & Brown & SEP_CHAR & END_CHAR)
                        
        ' Calculate exp to give attacker
        STR = Npc(ResourceNum).STR
        DEF = Npc(ResourceNum).DEF
        EXP = Npc(ResourceNum).EXP
        ExpType = Npc(ResourceNum).ExpType
        
        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 1
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If Player(Attacker).InParty = NO Then
            ' Check for non standard experience
            If ExpType <= 0 Then
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "  Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                Call SendStats(Attacker)
                ExpSentAttacker = 1
            Else
                ' Determine the skill type to give EXP to
                For i = 1 To MAX_PLAYER_SKILLS
                    If GetPlayerSkill(Attacker, i) = ExpType Then
                        Call SetPlayerSkillExp(Attacker, i, GetPlayerSkillExp(Attacker, i) + EXP)
                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & Trim(Skill(GetPlayerSkill(Attacker, i)).Name) & " Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                        'Call SendPlayerSkillsExp(Attacker, I)
                        Call CheckSkillLevelUp(Attacker, i)
                        ExpSentAttacker = 1
                    End If
                Next i
            End If
            If ExpSentAttacker = 0 Then
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "  Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                Call SendStats(Attacker)
                ExpSentAttacker = 1
            End If
        Else
            EXP = EXP / 2
            
            If EXP < 0 Then
                EXP = 1
            End If
            
            ' Check for non standard experience
            If ExpType <= 0 Then
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Party Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                Call SendStats(Attacker)
                ExpSentParty = 1
            Else
                ' Determine the skill type to give EXP to
                For i = 1 To MAX_PLAYER_SKILLS
                    If GetPlayerSkill(Attacker, i) = ExpType Then
                        Call SetPlayerSkillExp(Attacker, i, GetPlayerSkillExp(Attacker, i) + EXP)
                        Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & Trim(Skill(GetPlayerSkill(Attacker, i)).Name) & " Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                        'Call SendPlayerSkillsExp(Attacker, I)
                        Call CheckSkillLevelUp(Attacker, i)
                        ExpSentParty = 1
                    End If
                Next i
            End If
            If ExpSentParty = 0 Then
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                Call SendDataTo(Attacker, "BLITPLAYERMSG" & SEP_CHAR & "Party Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                Call SendStats(Attacker)
                ExpSentParty = 1
            End If
            
            n = Player(Attacker).PartyPlayer
            If n > 0 Then
                ' Check for non standard experience
                If ExpType <= 0 Then
                    Call SetPlayerExp(n, GetPlayerExp(n) + EXP)
                    Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & "Party Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                    Call SendStats(n)
                    ExpSentPartyPlayer = 1
                Else
                    ' Determine the skill type to give EXP to
                    For i = 1 To MAX_PLAYER_SKILLS
                        If GetPlayerSkill(n, i) = ExpType Then
                            Call SetPlayerSkillExp(n, i, GetPlayerSkillExp(n, i) + EXP)
                            Call SendDataTo(n, "BLITWARNMSG" & SEP_CHAR & "Party Experience" & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                            Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & Trim(Skill(GetPlayerSkill(n, i)).Name) & " Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                            'Call SendPlayerSkillsExp(n, I)
                            Call CheckSkillLevelUp(n, i)
                            ExpSentPartyPlayer = 1
                        End If
                    Next i
                End If
                If ExpSentPartyPlayer = 0 Then
                    Call SetPlayerExp(n, GetPlayerExp(n) + EXP)
                    Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & "Party Exp: " & EXP & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                    Call SendStats(n)
                    ExpSentPartyPlayer = 1
                End If
            End If
        End If
                                
        For i = 1 To MAX_NPC_DROPS
            ' Drop the goods if they get it
            J = Npc(ResourceNum).ItemNPC(i).Chance
            AmuletSlot = GetPlayerAmuletSlot(Attacker)
            RingSlot = GetPlayerRingSlot(Attacker)
            For K = 1 To MAX_PLAYER_SKILLS
                If GetPlayerSkill(Attacker, K) > 0 Then
                    If (Skill(GetPlayerSkill(Attacker, K)).Type = SKILL_TYPE_CHANCE) And (Skill(GetPlayerSkill(Attacker, K)).Data1 = SKILL_CHANCE_DROP) Then
                        J = J - Int((J / 100) * (Skill(GetPlayerSkill(Attacker, K)).Data2 * GetPlayerSkillLevel(Attacker, K)))
                    End If
                End If
            Next K
            If AmuletSlot > 0 Then
                If Item(GetPlayerInvItemNum(Attacker, AmuletSlot)).Data1 = CHARM_TYPE_ADDDROP Then
                    J = J - Int((J / 100) * Item(GetPlayerInvItemNum(Attacker, AmuletSlot)).Data2)
                End If
            End If
            If RingSlot > 0 Then
                If Item(GetPlayerInvItemNum(Attacker, RingSlot)).Data1 = CHARM_TYPE_ADDDROP Then
                    J = J - Int((J / 100) * Item(GetPlayerInvItemNum(Attacker, RingSlot)).Data2)
                End If
            End If
            n = Int(Rnd * J) + 1
            If n = 1 Then
                Call SpawnItem(Npc(ResourceNum).ItemNPC(i).ItemNum, Npc(ResourceNum).ItemNPC(i).ItemValue, MapNum, MapResource(MapNum, MapResourceNum).x, MapResource(MapNum, MapResourceNum).y)
            End If
        Next i
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapResource(MapNum, MapResourceNum).Num = 0
        MapResource(MapNum, MapResourceNum).SpawnWait = GetTickCount
        MapResource(MapNum, MapResourceNum).HP = 0
        Call SendDataToMap(MapNum, "RESOURCEDEAD" & SEP_CHAR & MapResourceNum & SEP_CHAR & END_CHAR)
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check for level up party member
        If Player(Attacker).InParty = YES Then
            Call CheckPlayerLevelUp(Player(Attacker).PartyPlayer)
        End If
    
        ' Check if target is resource that died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_RESOURCE And Player(Attacker).Target = MapResourceNum Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
    Else
        ' NPC not dead, just do the damage
        MapResource(MapNum, MapResourceNum).HP = MapResource(MapNum, MapResourceNum).HP - Damage
    End If
    
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim Packet As String
Dim ShopNum As Long, OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    Call SendLeaveMap(Index, OldMap)
    
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    
    Player(Index).GettingMap = YES
    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim MapNum As Long
Dim x As Long
Dim y As Long
Dim i As Long
Dim Moved As Byte

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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).WalkUp <> 1 Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).WalkDown <> 1 Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).WalkLeft <> 1 Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
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
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).WalkRight <> 1 Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
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
        
    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        x = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3
                        
        Call PlayerWarp(Index, MapNum, x, y)
        Moved = YES
    End If
    
    ' Check for key trigger open
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KEYOPEN Then
        x = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        
        If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = NO Then
            TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Door Opened" & SEP_CHAR & Brown & SEP_CHAR & END_CHAR)
        End If
    End If
    
    ' Check for pushblock trigger
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_PUSHBLOCK Then
        x = GetPlayerX(Index)
        y = GetPlayerY(Index)
        
        If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_PUSHBLOCK And PushTile(GetPlayerMap(Index)).Pushed(x, y) = NO Then
            PushTile(GetPlayerMap(Index)).Pushed(x, y) = YES
            PushTile(GetPlayerMap(Index)).PushedTimer = GetTickCount
            
            Call SendDataToMap(GetPlayerMap(Index), "PUSHBLOCK" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR)
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Block Pushed" & SEP_CHAR & Brown & SEP_CHAR & END_CHAR)
        End If
    End If
    
    ' They tried to hack
    If Moved = NO Then
        Call HackingAttempt(Index, "Position Modification")
    End If
End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
Dim i As Long

    ' Check if attacker got a level up
    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                    
        ' Get the ammount of skill points to add
        i = Int(GetPlayerSPEED(Index) / 10)
        If i < 1 Then i = 1
        If i > 3 Then i = 3
            
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + i)
        Call SetPlayerExp(Index, 0)
        'Call GlobalMsg(GetPlayerName(Index) & " has gained a level!", Brown)
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Level Up!" & SEP_CHAR & BrightGreen & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub CheckSkillLevelUp(ByVal Index As Long, ByVal SkillSlot As Long)
Dim i As Long

    ' Check if attacker got a level up
    If GetPlayerSkillExp(Index, SkillSlot) >= GetSkillNextLevel(Index, SkillSlot) Then
        Call SetPlayerSkillLevel(Index, SkillSlot, GetPlayerSkillLevel(Index, SkillSlot) + 1)
                    
        '' Get the ammount of skill points to add
        'I = Int(GetPlayerSPEED(Index) / 10)
        'If I < 1 Then I = 1
        'If I > 3 Then I = 3
            
        'Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + I)
        Call SetPlayerSkillExp(Index, SkillSlot, 0)
        'Call GlobalMsg(GetPlayerName(Index) & " has gained a level!", Brown)
        Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Skill Up!" & SEP_CHAR & BrightGreen & SEP_CHAR & END_CHAR)
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Trim(Skill(GetPlayerSkill(Index, SkillSlot)).Name) & SEP_CHAR & BrightGreen & SEP_CHAR & END_CHAR)
        Call SendPlayerSkillsLevel(Index, SkillSlot)
        Call SendPlayerSkillsExp(Index, SkillSlot)
    End If
    Call SendPlayerSkillsExp(Index, SkillSlot)
End Sub

Sub CheckSpellLevelUp(ByVal Index As Long, ByVal SpellSlot As Long)
Dim i As Long

    ' Check if attacker got a level up
    If GetPlayerSpellExp(Index, SpellSlot) >= GetSpellNextLevel(Index, SpellSlot) Then
        Call SetPlayerSpellLevel(Index, SpellSlot, GetPlayerSpellLevel(Index, SpellSlot) + 1)
                    
        '' Get the ammount of spell points to add
        'I = Int(GetPlayerSPEED(Index) / 10)
        'If I < 1 Then I = 1
        'If I > 3 Then I = 3
            
        'Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + I)
        Call SetPlayerSpellExp(Index, SpellSlot, 0)
        'Call GlobalMsg(GetPlayerName(Index) & " has gained a level!", Brown)
        Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Spell Up!" & SEP_CHAR & BrightGreen & SEP_CHAR & END_CHAR)
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Trim(Spell(GetPlayerSpell(Index, SpellSlot)).Name) & SEP_CHAR & BrightGreen & SEP_CHAR & END_CHAR)
        Call SendPlayerSpellsLevel(Index, SpellSlot)
        Call SendPlayerSpellsExp(Index, SpellSlot)
    End If
    Call SendPlayerSpellsExp(Index, SpellSlot)
End Sub

Sub CastSpell(ByVal Index As Long, ByVal SpellSlot As Long)
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
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Spell Not Available" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    i = GetSpellReqLevel(Index, SpellNum)
    MPReq = Int((GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) / 2)
    If MPReq <= 0 Then
        MPReq = 1
    End If
    
    ' Check if they have enough MP
    If GetPlayerMP(Index) < MPReq Then
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Insufficient Mana" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If i > GetPlayerLevel(Index) Then
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Level Req: " & i & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
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
            'Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim(Spell(SpellNum).Name) & ".", BrightBlue)
            
            ' Take away the mana points
            Call SetPlayerMP(Index, GetPlayerMP(Index) - MPReq - 5)
            Call SendMP(Index)
            ' Deal with exp for spell
            Call SetPlayerSpellExp(Index, SpellSlot, GetPlayerSpellExp(Index, SpellSlot) + MPReq)
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Trim(Spell(GetPlayerSpell(Index, SpellSlot)).Name) & " Exp: " & MPReq & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
            Casted = True
        Else
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Inventory Full" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        End If
        
        Call CheckSpellLevelUp(Index, SpellSlot)
        Exit Sub
    End If
        
    n = Player(Index).Target
    If n = 0 Then
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "No Target" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        If IsPlaying(n) Then
            If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
            'If GetPlayerHP(n) > 0 And GetPlayerMap(Index) = GetPlayerMap(n) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(n) >= 10 And Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(n) <= 0 Then
                If GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Data1 >= SPELL_STAT_SUBHP And Spell(SpellNum).Data1 <= SPELL_STAT_SUBSP Then
                    'If GetPlayerLevel(n) + 5 >= GetPlayerLevel(Index) Then
                        'If GetPlayerLevel(n) - 5 <= GetPlayerLevel(Index) Then
                            'Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                
                            Select Case Spell(SpellNum).Type
                                Case SPELL_TYPE_STAT
                                    Select Case Spell(SpellNum).Data1
                                        Case SPELL_STAT_SUBHP
                        
                                            Damage = (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                                            If Damage > 0 Then
                                                Call AttackPlayer(Index, n, Damage)
                                            Else
                                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Spell Fizzled" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                            End If
                    
                                        Case SPELL_STAT_SUBMP
                                            If GetPlayerMP(n) > 0 Then
                                                Call SetPlayerMP(n, GetPlayerMP(n) - (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2))
                                                Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & "MP Lost: +" & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                                Call SendMP(n)
                                            Else
                                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Target MP At 0" & SEP_CHAR & Red & SEP_CHAR & END_CHAR)
                                                'Call SetPlayerMP(n, 0)
                                                Exit Sub
                                            End If
                
                                        Case SPELL_STAT_SUBSP
                                            If GetPlayerSP(n) > 0 Then
                                                Call SetPlayerSP(n, GetPlayerSP(n) - (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2))
                                                Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & "SP Lost: +" & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                                Call SendSP(n)
                                            Else
                                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Target SP At 0" & SEP_CHAR & Red & SEP_CHAR & END_CHAR)
                                                'Call SetPlayerSP(n, 0)
                                                Exit Sub
                                            End If
                                    End Select
                            End Select
                        'Else
                            'Call PlayerMsg(Index, GetPlayerName(n) & " is far to powerful to even consider attacking.", BrightBlue)
                        'End If
                    'Else
                        'Call PlayerMsg(Index, GetPlayerName(n) & " is to weak to even bother with.", BrightBlue)
                    'End If
            
                    ' Take away the mana points
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - MPReq)
                    Call SendMP(Index)
                    ' Deal with exp for spell
                    Call SetPlayerSpellExp(Index, SpellSlot, GetPlayerSpellExp(Index, SpellSlot) + MPReq)
                    Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Spell(GetPlayerSpell(Index, SpellSlot)).Name) & " Exp: " & MPReq & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                    Casted = True
                ElseIf GetPlayerMap(Index) = GetPlayerMap(n) And Spell(SpellNum).Data1 >= SPELL_STAT_ADDHP And Spell(SpellNum).Data1 <= SPELL_STAT_ADDSP Then
                    Select Case Spell(SpellNum).Type
                        Case SPELL_TYPE_STAT
                            Select Case Spell(SpellNum).Data1
                                Case SPELL_STAT_ADDHP
                                    'Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                    If GetPlayerHP(n) < GetPlayerMaxHP(n) Then
                                        Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & "HP  + " & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                                        Call SetPlayerHP(n, GetPlayerHP(n) + (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2))
                                        Call SendHP(n)
                                    Else
                                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Target HP Full" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                        Exit Sub
                                    End If
                                    
                                Case SPELL_STAT_ADDMP
                                    'Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                    If GetPlayerMP(n) < GetPlayerMaxMP(n) Then
                                        Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & "MP  + " & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                                        Call SetPlayerMP(n, GetPlayerMP(n) + (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2))
                                        Call SendMP(n)
                                    Else
                                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Target MP Full" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                        Exit Sub
                                    End If
                    
                                Case SPELL_STAT_ADDSP
                                    'Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                    If GetPlayerSP(n) < GetPlayerMaxSP(n) Then
                                        Call SendDataTo(n, "BLITPLAYERMSG" & SEP_CHAR & "SP  + " & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                                        Call SetPlayerMP(n, GetPlayerSP(n) + (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2))
                                        Call SendSP(n)
                                    Else
                                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Target SP Full" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                        Exit Sub
                                    End If
                            End Select
                    End Select
                    
                    ' Take away the mana points
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - MPReq)
                    Call SendMP(Index)
                    ' Deal with exp for spell and check level up
                    Call SetPlayerSpellExp(Index, SpellSlot, GetPlayerSpellExp(Index, SpellSlot) + MPReq)
                    Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Spell(GetPlayerSpell(Index, SpellSlot)).Name) & " Exp: " & MPReq & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                    Casted = True
                Else
                    ' Add elseif for additional spell types
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Spell Failed" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                End If
            End If
        Else
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Spell Failed" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        End If
    ElseIf Player(Index).TargetType = TARGET_TYPE_NPC Then
        If Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then 'And Npc(MapResource(GetPlayerMap(Index), n).Num).Behavior <> NPC_BEHAVIOR_RESOURCE Then
            'Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim(Spell(SpellNum).Name) & " on a " & Trim(Npc(MapNpc(GetPlayerMap(Index), n).Num).Name) & ".", BrightBlue)
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_STAT
                    Select Case Spell(SpellNum).Data1
                        Case SPELL_STAT_ADDHP
                            MapNpc(GetPlayerMap(Index), n).HP = MapNpc(GetPlayerMap(Index), n).HP + (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                            Call SendDataTo(Index, "BLITNPCMSG" & SEP_CHAR & "  HP +" & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & n & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                
                        Case SPELL_STAT_SUBHP
                            Damage = (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                            If Damage > 0 Then
                                Call AttackNpc(Index, n, Damage)
                                Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & n & SEP_CHAR & White & SEP_CHAR & END_CHAR)
                            Else
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Spell Fizzled" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                            End If
                    
                        Case SPELL_STAT_ADDMP
                            MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP + (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                            Call SendDataTo(Index, "BLITNPCMSG" & SEP_CHAR & "  MP +" & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & n & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                
                        Case SPELL_STAT_SUBMP
                            If MapNpc(GetPlayerMap(Index), n).MP > 0 Then
                                MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP - (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                                Call SendDataTo(Index, "BLITNPCMSG" & SEP_CHAR & "  MP -" & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & n & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                            Else
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Target MP At 0" & SEP_CHAR & Red & SEP_CHAR & END_CHAR)
                                MapNpc(GetPlayerMap(Index), n).MP = 0
                                Exit Sub
                            End If
            
                        Case SPELL_STAT_ADDSP
                            MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP + (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                            Call SendDataTo(Index, "BLITNPCMSG" & SEP_CHAR & "  SP +" & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & n & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                
                        Case SPELL_STAT_SUBSP
                            If MapNpc(GetPlayerMap(Index), n).SP > 0 Then
                                MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP - (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                                Call SendDataTo(Index, "BLITNPCMSG" & SEP_CHAR & "  SP -" & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & n & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                            Else
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Target SP At 0" & SEP_CHAR & Red & SEP_CHAR & END_CHAR)
                                MapNpc(GetPlayerMap(Index), n).SP = 0
                                Exit Sub
                            End If
                End Select
            End Select
        
            ' Take away the mana points
            Call SetPlayerMP(Index, GetPlayerMP(Index) - MPReq)
            Call SendMP(Index)
            Call SetPlayerSpellExp(Index, SpellSlot, GetPlayerSpellExp(Index, SpellSlot) + MPReq)
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Trim(Spell(GetPlayerSpell(Index, SpellSlot)).Name) & " Exp: " & MPReq & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
            Casted = True
        Else
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Cannot Cast" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        End If
    ElseIf Player(Index).TargetType = TARGET_TYPE_RESOURCE Then
        If Npc(MapResource(GetPlayerMap(Index), n).Num).HitOnlyWith <= 0 Then
            'Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim(Spell(SpellNum).Name) & " on a " & Trim(Npc(MapNpc(GetPlayerMap(Index), n).Num).Name) & ".", BrightBlue)
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_STAT
                    Select Case Spell(SpellNum).Data1
                        Case SPELL_STAT_ADDHP
                            MapResource(GetPlayerMap(Index), n).HP = MapResource(GetPlayerMap(Index), n).HP + (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                            Call SendDataTo(Index, "BLITRESOURCEMSG" & SEP_CHAR & n & SEP_CHAR & "+ " & (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2) & SEP_CHAR & i & SEP_CHAR & Red & SEP_CHAR & END_CHAR)
                
                        Case SPELL_STAT_SUBHP
                    
                            Damage = (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                            If Damage > 0 Then
                                Call AttackResource(Index, n, Damage)
                                Call SendDataTo(Index, "BLITRESOURCEDMG" & SEP_CHAR & Damage & SEP_CHAR & Player(Index).Target & SEP_CHAR & White & SEP_CHAR & END_CHAR)
                            Else
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Spell Fizzled" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                            End If
                    
                        'Case SPELL_STAT_ADDMP
                            'MapResource(GetPlayerMap(Index), n).MP = MapResource(GetPlayerMap(Index), n).MP + (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                
                        'Case SPELL_STAT_SUBMP
                            'MapResource(GetPlayerMap(Index), n).MP = MapResource(GetPlayerMap(Index), n).MP - (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
            
                        'Case SPELL_STAT_ADDSP
                            'MapResource(GetPlayerMap(Index), n).SP = MapResource(GetPlayerMap(Index), n).SP + (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                
                        'Case SPELL_STAT_SUBSP
                            'MapResource(GetPlayerMap(Index), n).SP = MapResource(GetPlayerMap(Index), n).SP - (GetPlayerSpellLevel(Index, SpellSlot) * Spell(SpellNum).Data2)
                    End Select
            End Select
        
            ' Take away the mana points
            Call SetPlayerMP(Index, GetPlayerMP(Index) - MPReq)
            Call SendMP(Index)
            Call SetPlayerSpellExp(Index, SpellSlot, GetPlayerSpellExp(Index, SpellSlot) + MPReq)
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Trim(Spell(GetPlayerSpell(Index, SpellSlot)).Name) & " Exp: " & MPReq & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
            Casted = True
        Else
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Cannot Cast" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        End If
    Else
        ' Space for spells that can be used on items
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Cannot Cast" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
    End If

    Call CheckSpellLevelUp(Index, SpellSlot)
    If Casted = True Then
        Player(Index).AttackTimer = GetTickCount
        Player(Index).CastedSpell = YES
    End If
End Sub

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
Dim i As Long, J As Long, n As Long
Dim AmuletSlot As Long
Dim RingSlot As Long

    CanPlayerCriticalHit = False
    AmuletSlot = GetPlayerAmuletSlot(Index)
    RingSlot = GetPlayerRingSlot(Index)
    
    If GetPlayerWeaponSlot(Index) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerSTR(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
    
            n = Int(Rnd * 100) + 1
            For J = 1 To MAX_PLAYER_SKILLS
                If GetPlayerSkill(Index, J) > 0 Then
                    If (Skill(GetPlayerSkill(Index, J)).Type = SKILL_TYPE_CHANCE) And (Skill(GetPlayerSkill(Index, J)).Data1 = SKILL_CHANCE_CRIT) Then
                        n = n - (Skill(GetPlayerSkill(Index, J)).Data2 * GetPlayerSkillLevel(Index, J))
                    End If
                End If
            Next J
            If AmuletSlot > 0 Then
                If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDCRIT Then
                    n = n - Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
                End If
            End If
            If RingSlot > 0 Then
                If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDCRIT Then
                    n = n - Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
                End If
            End If
            If n <= i Then
                CanPlayerCriticalHit = True
                Call AttributeSkillsExp(Index, SKILL_TYPE_CHANCE, SKILL_CHANCE_CRIT, i, False)
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
Dim i As Long, J As Long, n As Long
Dim ShieldSlot As Long
Dim AmuletSlot As Long
Dim RingSlot As Long

    CanPlayerBlockHit = False
    
    ShieldSlot = GetPlayerShieldSlot(Index)
    AmuletSlot = GetPlayerAmuletSlot(Index)
    RingSlot = GetPlayerRingSlot(Index)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
        
            n = Int(Rnd * 100) + 1
            For J = 1 To MAX_PLAYER_SKILLS
                If GetPlayerSkill(Index, J) > 0 Then
                    If (Skill(GetPlayerSkill(Index, J)).Type = SKILL_TYPE_CHANCE) And (Skill(GetPlayerSkill(Index, J)).Data1 = SKILL_CHANCE_BLOCK) Then
                        n = n - (Skill(GetPlayerSkill(Index, J)).Data2 * GetPlayerSkillLevel(Index, J))
                    End If
                End If
            Next J
            If AmuletSlot > 0 Then
                If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDBLOCK Then
                    n = n - Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
                End If
            End If
            If RingSlot > 0 Then
                If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDBLOCK Then
                    n = n - Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
                End If
            End If
            If n <= i Then
                CanPlayerBlockHit = True
                Call SetPlayerInvItemDur(Index, ShieldSlot, GetPlayerInvItemDur(Index, ShieldSlot) - 1)
                Call SendInventoryUpdate(Index, ShieldSlot)
                Call AttributeSkillsExp(Index, SKILL_TYPE_CHANCE, SKILL_CHANCE_BLOCK, i, False)
                Call AttributeSkillsExp(Index, SKILL_TYPE_ATTRIBUTE, SKILL_ATTRIBUTE_DEF, i, False)
            End If
        End If
    End If
End Function

Function IsAccurate(ByVal Index As Long) As Boolean
Dim i As Long, J As Long, n As Long
Dim AmuletSlot As Long
Dim RingSlot As Long

    IsAccurate = False
    AmuletSlot = GetPlayerAmuletSlot(Index)
    RingSlot = GetPlayerRingSlot(Index)
    
    If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 = WEAPON_SUBTYPE_BOW Then
        i = 50
        n = Int(Rnd * 100) + 1
        For J = 1 To MAX_PLAYER_SKILLS
            If GetPlayerSkill(Index, J) > 0 Then
                If (Skill(GetPlayerSkill(Index, J)).Type = SKILL_TYPE_CHANCE) And (Skill(GetPlayerSkill(Index, J)).Data1 = SKILL_CHANCE_ACCU) Then
                    n = n - (Skill(GetPlayerSkill(Index, J)).Data2 * GetPlayerSkillLevel(Index, J))
                End If
            End If
        Next J
        If AmuletSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data1 = CHARM_TYPE_ADDACCU Then
                n = n - Item(GetPlayerInvItemNum(Index, AmuletSlot)).Data2
            End If
        End If
        If RingSlot > 0 Then
            If Item(GetPlayerInvItemNum(Index, RingSlot)).Data1 = CHARM_TYPE_ADDACCU Then
                n = n - Item(GetPlayerInvItemNum(Index, RingSlot)).Data2
            End If
        End If
        If n <= i Then
            IsAccurate = True
            Call AttributeSkillsExp(Index, SKILL_TYPE_CHANCE, SKILL_CHANCE_ACCU, 1, False)
        End If
    End If
End Function

