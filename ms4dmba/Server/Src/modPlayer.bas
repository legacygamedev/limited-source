Attribute VB_Name = "modPlayer"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Sub JoinGame(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
        
    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", White)
    End If

    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
        
    ' Send an ok to client to start receiving in game data
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong Index
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
        
    TotalPlayersOnline = TotalPlayersOnline + 1
    Call UpdateHighIndex

    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
    Next
    
    Call SendStats(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            
    ' Send welcome messages
    Call SendWelcome(Index)

    ' Send the flag so they know they can start doing stuff
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SInGame
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub LeftGame(ByVal Index As Long)
Dim n As Long

    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
        ' Check for boot map
        'If Map(GetPlayerMap(Index)).BootMap > 0 Then
        '    Call SetPlayerX(Index, Map(GetPlayerMap(Index)).BootX)
        '    Call SetPlayerY(Index, Map(GetPlayerMap(Index)).BootY)
        '    Call SetPlayerMap(Index, Map(GetPlayerMap(Index)).BootMap)
        'End If
        
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If TempPlayer(Index).InParty = YES Then
            n = TempPlayer(Index).PartyPlayer
            
            Call PlayerMsg(n, GetPlayerName(Index) & " has left " & GAME_NAME & ", disbanning party.", Pink)
            TempPlayer(n).InParty = NO
            TempPlayer(n).PartyPlayer = 0
        End If
            
        Call SavePlayer(Index)
    
        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", White)
        End If
        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".")
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
        Call UpdateHighIndex
    End If
    
    Call ClearPlayer(Index)
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim n As Long
Dim i As Long
Dim STR As Long
Dim DEF As Long
Dim MapNum As Long
Dim NpcNum As Long
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    Name = Trim$(Npc(NpcNum).Name)
    
    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAttack
    Buffer.WriteLong Attacker
    
    SendDataToMapBut Attacker, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
     
    ' Check for weapon
    n = 0
    If GetPlayerEquipmentSlot(Attacker, Weapon) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerEquipmentSlot(Attacker, Weapon))
    End If
    
    If Damage >= MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) Then
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points, killing it.", BrightRed)
        Else
            Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points, killing it.", BrightRed)
        End If
                        
        ' Calculate exp to give attacker
        STR = Npc(NpcNum).Stat(Stats.Strength)
        DEF = Npc(NpcNum).Stat(Stats.Defense)
        Exp = STR * DEF * 2
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        ' Check if in party, if so divide the exp up by 2
        If TempPlayer(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You have gained " & Exp & " experience points.", BrightBlue)
        Else
            Exp = Exp / 2
            
            If Exp < 0 Then
                Exp = 1
            End If
            
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You have gained " & Exp & " party experience points.", BrightBlue)
            
            n = TempPlayer(Attacker).PartyPlayer
            If n > 0 Then
                Call SetPlayerExp(n, GetPlayerExp(n) + Exp)
                Call PlayerMsg(n, "You have gained " & Exp & " party experience points.", BrightBlue)
            End If
        End If
                                
        ' Drop the goods if they get it
        n = Int(Rnd * Npc(NpcNum).DropChance) + 1
        If n = 1 Then
            Call SpawnItem(Npc(NpcNum).DropItem, Npc(NpcNum).DropItemValue, MapNum, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(MapNpcNum).Num = 0
        MapNpc(MapNum).Npc(MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = 0
        
        Set Buffer = New clsBuffer
        
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong MapNpcNum
        
        SendDataToMap MapNum, Buffer.ToArray()
        
        Set Buffer = Nothing
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check for level up party member
        If TempPlayer(Attacker).InParty = YES Then
            Call CheckPlayerLevelUp(TempPlayer(Attacker).PartyPlayer)
        End If
    
        ' Check if target is npc that died and if so set target to 0
        If TempPlayer(Attacker).TargetType = TARGET_TYPE_NPC Then
            If TempPlayer(Attacker).Target = MapNpcNum Then
                TempPlayer(Attacker).Target = 0
                TempPlayer(Attacker).TargetType = TARGET_TYPE_NONE
            End If
        End If
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) - Damage
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        Else
            Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If
        
        ' Check if we should send a message
        If MapNpc(MapNum).Npc(MapNpcNum).Target = 0 Then
            If LenB(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                Call PlayerMsg(Attacker, "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
            End If
        End If
        
        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(MapNpcNum).Target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(MapNpcNum).Num Then
                    MapNpc(MapNum).Npc(i).Target = Attacker
                End If
            Next
        End If
    End If
    
    ' Reduce durability of weapon
    Call DamageEquipment(Attacker, Weapon)
    
    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Exp As Long
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
        
    ' Check for weapon
    n = 0
    If GetPlayerEquipmentSlot(Attacker, Weapon) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerEquipmentSlot(Attacker, Weapon))
    End If
    
    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAttack
    Buffer.WriteLong Attacker
    
    SendDataToMapBut Attacker, GetPlayerMap(Attacker), Buffer.ToArray()
    
    Set Buffer = Nothing

    ' reduce dur. on victims equipment
    Call DamageEquipment(Victim, Armor)
    Call DamageEquipment(Victim, Helmet)
    
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
           
        ' Calculate exp to give attacker
        Exp = (GetPlayerExp(Victim) \ 10)
        
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
                
        ' Check for a level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check if target is player who died and if so set target to 0
        If TempPlayer(Attacker).TargetType = TARGET_TYPE_PLAYER Then
            If TempPlayer(Attacker).Target = Victim Then
                TempPlayer(Attacker).Target = 0
                TempPlayer(Attacker).TargetType = TARGET_TYPE_NONE
            End If
        End If
        
        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If
        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If
        
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
        End If
    End If
    
    ' Reduce durability of weapon
    Call DamageEquipment(Attacker, Weapon)
    
    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

Function GetPlayerDamage(ByVal Index As Long) As Long
Dim WeaponSlot As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    GetPlayerDamage = (GetPlayerStat(Index, Stats.Strength) \ 2)
    
    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If
    
    If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
        WeaponSlot = GetPlayerEquipmentSlot(Index, Weapon)
        
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2
    End If
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
Dim ArmorSlot As Long
Dim HelmSlot As Long
    
    GetPlayerProtection = 0
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ArmorSlot = GetPlayerEquipmentSlot(Index, Armor)
    HelmSlot = GetPlayerEquipmentSlot(Index, Helmet)
    
    GetPlayerProtection = (GetPlayerStat(Index, Stats.Defense) \ 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data2
    End If
    
    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, HelmSlot)).Data2
    End If
End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
Dim i As Long
Dim n As Long
   
    If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
    
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
Dim i As Long
Dim n As Long
Dim ShieldSlot As Long

    ShieldSlot = GetPlayerEquipmentSlot(Index, Shield)
    
    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)
        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.Defense) \ 2) + (GetPlayerLevel(Index) \ 2)
        
            n = Int(Rnd * 100) + 1
            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Sub CastSpell(ByVal Index As Long, ByVal SpellSlot As Long)
Dim SpellNum As Long
Dim MPReq As Long
Dim i As Long
Dim n As Long
Dim Damage As Long
Dim Casted As Boolean
Dim CanCast As Boolean
Dim TargetType As Byte
Dim TargetName As String
Dim Buffer As clsBuffer

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

    MPReq = Spell(SpellNum).MPReq
    
    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPReq Then
        Call PlayerMsg(Index, "Not enough mana points!", BrightRed)
        Exit Sub
    End If
        
    ' Make sure they are the right level
    If i > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & i & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ' Check if timer is ok
    If GetTickCount < TempPlayer(Index).AttackTimer + 1000 Then
        Exit Sub
    End If
    
    ' *** Self Cast Spells ***
    ' Check if the spell is a give item and do that instead of a stat modification
    If Spell(SpellNum).Type = SPELL_TYPE_GIVEITEM Then
        n = FindOpenInvSlot(Index, Spell(SpellNum).Data1)
        
        If n > 0 Then
            Call GiveItem(Index, Spell(SpellNum).Data1, Spell(SpellNum).Data2)
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & ".", BrightBlue)
            
            ' Take away the mana points
            Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPReq)
            Call SendVital(Index, Vitals.MP)
            Casted = True
        Else
            Call PlayerMsg(Index, "Your inventory is full!", BrightRed)
        End If
        
        Exit Sub
    End If
        
    n = TempPlayer(Index).Target
    TargetType = TempPlayer(Index).TargetType
    
    Select Case TargetType
        Case TARGET_TYPE_PLAYER
    
            If IsPlaying(n) Then
                
                If GetPlayerVital(n, Vitals.HP) > 0 Then
                    If GetPlayerMap(Index) = GetPlayerMap(n) Then
                        'If GetPlayerLevel(Index) >= 10 Then
                            'If GetPlayerLevel(n) >= 10 Then
                                If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Then
                                    If GetPlayerAccess(Index) <= 0 Then
                                        If GetPlayerAccess(n) <= 0 Then
                                            If n <> Index Then
                                                CanCast = True
                                            End If
                                        End If
                                    End If
                                End If
                            'End If
                        'End If
                    End If
                End If
                
                TargetName = GetPlayerName(n)
                
                If Spell(SpellNum).Type = SPELL_TYPE_SUBHP Or _
                   Spell(SpellNum).Type = SPELL_TYPE_SUBMP Or _
                   Spell(SpellNum).Type = SPELL_TYPE_SUBSP Then
                   
                    If CanCast Then
                        Select Case Spell(SpellNum).Type
                            Case SPELL_TYPE_SUBHP
                                Damage = (GetPlayerStat(Index, Stats.Magic) \ 4) + Spell(SpellNum).Data1 - GetPlayerProtection(n)
                                If Damage > 0 Then
                                    Call AttackPlayer(Index, n, Damage)
                                Else
                                    Call PlayerMsg(Index, "The spell was to weak to hurt " & GetPlayerName(n) & "!", BrightRed)
                                End If
                        
                            Case SPELL_TYPE_SUBMP
                                Call SetPlayerVital(n, Vitals.MP, GetPlayerVital(n, Vitals.MP) - Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.MP)
                        
                            Case SPELL_TYPE_SUBSP
                                Call SetPlayerVital(n, Vitals.SP, GetPlayerVital(n, Vitals.SP) - Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.SP)
                        End Select
                        
                        Casted = True
                        
                    End If
                
                ElseIf Spell(SpellNum).Type = SPELL_TYPE_ADDHP Or _
                       Spell(SpellNum).Type = SPELL_TYPE_ADDMP Or _
                       Spell(SpellNum).Type = SPELL_TYPE_ADDSP Then
                
                    If GetPlayerMap(Index) = GetPlayerMap(n) Then
                        CanCast = True
                    End If
                    
                    If CanCast Then
                        Select Case Spell(SpellNum).Type
                            Case SPELL_TYPE_ADDHP
                                Call SetPlayerVital(n, Vitals.HP, GetPlayerVital(n, Vitals.HP) + Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.HP)
                                        
                            Case SPELL_TYPE_ADDMP
                                Call SetPlayerVital(n, Vitals.MP, GetPlayerVital(n, Vitals.MP) + Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.MP)
                        
                            Case SPELL_TYPE_ADDSP
                                Call SetPlayerVital(n, Vitals.SP, GetPlayerVital(n, Vitals.SP) + Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.SP)
                        End Select
                        
                        Casted = True
                    End If
                    
                End If
            End If
        
        Case TARGET_TYPE_NPC
    
            If Npc(MapNpc(GetPlayerMap(Index)).Npc(n).Num).Behavior <> NPC_BEHAVIOR_FRIENDLY Then
                If Npc(MapNpc(GetPlayerMap(Index)).Npc(n).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    CanCast = True
                End If
            End If
            
            TargetName = Npc(MapNpc(GetPlayerMap(Index)).Npc(n).Num).Name
                
            If CanCast Then
                Select Case Spell(SpellNum).Type
                    Case SPELL_TYPE_ADDHP
                        MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.HP) = MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.HP) + Spell(SpellNum).Data1
                    
                    Case SPELL_TYPE_SUBHP
                        
                        Damage = (GetPlayerStat(Index, Stats.Magic) \ 4) + Spell(SpellNum).Data1 - (Npc(MapNpc(GetPlayerMap(Index)).Npc(n).Num).Stat(Stats.Defense) \ 2)
                        If Damage > 0 Then
                            Call AttackNpc(Index, n, Damage)
                        Else
                            Call PlayerMsg(Index, "The spell was to weak to hurt " & Trim$(Npc(MapNpc(GetPlayerMap(Index)).Npc(n).Num).Name) & "!", BrightRed)
                        End If
                        
                    Case SPELL_TYPE_ADDMP
                        MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.MP) = MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.MP) + Spell(SpellNum).Data1
                    
                    Case SPELL_TYPE_SUBMP
                        MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.MP) = MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.MP) - Spell(SpellNum).Data1
                
                    Case SPELL_TYPE_ADDSP
                        MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.SP) = MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.SP) + Spell(SpellNum).Data1
                    
                    Case SPELL_TYPE_SUBSP
                        MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.SP) = MapNpc(GetPlayerMap(Index)).Npc(n).Vital(Vitals.SP) - Spell(SpellNum).Data1
                End Select
    
                Casted = True
            End If
            
    End Select

    If Casted Then
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & Trim$(TargetName) & ".", BrightBlue)
        
        Set Buffer = New clsBuffer
        
        Buffer.WriteLong SCastSpell
        Buffer.WriteLong TargetType
        Buffer.WriteLong n
        Buffer.WriteLong SpellNum
        
        SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
        
        Set Buffer = Nothing

        ' Take away the mana points
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPReq)
        Call SendVital(Index, Vitals.MP)
    
        TempPlayer(Index).AttackTimer = GetTickCount
        TempPlayer(Index).CastedSpell = YES
    Else
        Call PlayerMsg(Index, "Could not cast spell!", BrightRed)
    End If
    
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim ShopNum As Long
Dim OldMap As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Check if you are out of bounds
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
    
    ' Check if there was a shop on the map the player is leaving, and if so say goodbye
    ShopNum = Map(GetPlayerMap(Index)).Shop
    If ShopNum > 0 Then
        If LenB(Trim$(Shop(ShopNum).LeaveSay)) > 0 Then
            Call PlayerMsg(Index, Trim$(Shop(ShopNum).Name) & " says, '" & Trim$(Shop(ShopNum).LeaveSay) & "'", SayColor)
        End If
    End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    
    If OldMap <> MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If
    
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    
    ' Check if there is a shop on the map and say hello if so
    ShopNum = Map(GetPlayerMap(Index)).Shop
    If ShopNum > 0 Then
        If LenB(Trim$(Shop(ShopNum).JoinSay)) > 0 Then
            Call PlayerMsg(Index, Trim$(Shop(ShopNum).Name) & " says, '" & Trim$(Shop(ShopNum).JoinSay) & "'", SayColor)
        End If
    End If
            
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
        
        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(OldMap).Npc(i).Num > 0 Then
                MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP)
            End If
        Next
        
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    
    TempPlayer(Index).GettingMap = YES
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong MapNum
    Buffer.WriteLong Map(MapNum).Revision
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Buffer As clsBuffer
Dim MapNum As Long
Dim x As Long
Dim y As Long
Dim Moved As Byte
Dim NewMapX As Byte, NewMapY As Byte

    'TempPlayer(Index).GettingMap = Yes

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    'If TempPlayer(Index).CanPlayerMove = 1 Then
    '    Exit Sub
    'End If
    
    Call SetPlayerDir(Index, Dir)
    
    Moved = NO
    MapNum = GetPlayerMap(Index)
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                        
                        Set Buffer = New clsBuffer
                        
                        With Buffer
                            .WriteLong SPlayerMove
                            .WriteLong Index
                            .WriteLong GetPlayerX(Index)
                            .WriteLong GetPlayerY(Index)
                            .WriteLong GetPlayerDir(Index)
                            .WriteLong Movement
                            SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                        End With
                        
                        Set Buffer = Nothing
                        
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(Index)).Up).MaxY
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), NewMapY)
                    Moved = YES
                End If
            End If
                    
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < Map(MapNum).MaxY Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                        
                        Set Buffer = New clsBuffer
                        
                        With Buffer
                            .WriteLong SPlayerMove
                            .WriteLong Index
                            .WriteLong GetPlayerX(Index)
                            .WriteLong GetPlayerY(Index)
                            .WriteLong GetPlayerDir(Index)
                            .WriteLong Movement
                            SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                        End With
                        
                        Set Buffer = Nothing
                        
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
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                        
                        Set Buffer = New clsBuffer
                        
                        With Buffer
                            .WriteLong SPlayerMove
                            .WriteLong Index
                            .WriteLong GetPlayerX(Index)
                            .WriteLong GetPlayerY(Index)
                            .WriteLong GetPlayerDir(Index)
                            .WriteLong Movement
                            SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                        End With
                        
                        Set Buffer = Nothing
                        
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(Index)).Left).MaxX
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, NewMapX, GetPlayerY(Index))
                    Moved = YES
                End If
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < Map(MapNum).MaxX Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                        
                        Set Buffer = New clsBuffer
                        
                        With Buffer
                            .WriteLong SPlayerMove
                            .WriteLong Index
                            .WriteLong GetPlayerX(Index)
                            .WriteLong GetPlayerY(Index)
                            .WriteLong GetPlayerDir(Index)
                            .WriteLong Movement
                            SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                        End With
                        
                        Set Buffer = Nothing
                        
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
        
        'TempPlayer(Index).CanPlayerMove = 1
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
            
            Set Buffer = New clsBuffer
            
            Buffer.WriteLong SMapKey
            Buffer.WriteLong x
            Buffer.WriteLong y
            Buffer.WriteByte 1
                            
            SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
            
            Set Buffer = Nothing
            Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
        End If
    End If
    
    ' They tried to hack
    If Moved = NO Then
        'Call HackingAttempt(Index, "Position Modification")
        'TempPlayer(Index).CanPlayerMove = 1
        Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    End If
End Sub

Sub CheckEquippedItems(ByVal Index As Long)
Dim Slot As Long
Dim ItemNum As Long
Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        Slot = GetPlayerEquipmentSlot(Index, i)
        If Slot > 0 Then
            ItemNum = GetPlayerInvItemNum(Index, Slot)
            
            If ItemNum > 0 Then
                Select Case i
                    Case Equipment.Weapon
                        If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipmentSlot Index, 0, i
                    Case Equipment.Armor
                        If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipmentSlot Index, 0, i
                    Case Equipment.Helmet
                        If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipmentSlot Index, 0, i
                    Case Equipment.Shield
                        If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipmentSlot Index, 0, i
                End Select
            Else
               SetPlayerEquipmentSlot Index, 0, i
            End If
        End If
    Next
End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
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
        Next
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next
End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long

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
    Next
End Function

Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long
Dim n As Long
Dim TakeItem As Boolean
    
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
                        If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Weapon) Then
                                Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                
                    Case ITEM_TYPE_ARMOR
                        If GetPlayerEquipmentSlot(Index, Armor) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Armor) Then
                                Call SetPlayerEquipmentSlot(Index, 0, Armor)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_HELMET
                        If GetPlayerEquipmentSlot(Index, Helmet) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Helmet) Then
                                Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_SHIELD
                        If GetPlayerEquipmentSlot(Index, Shield) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Shield) Then
                                Call SetPlayerEquipmentSlot(Index, 0, Shield)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield)) Then
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
                            
            If TakeItem Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Call SetPlayerInvItemDur(Index, i, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
        End If
    Next
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
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
    End If
End Sub

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next
End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
Dim i As Long
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next
End Function

Sub PlayerMapGetItem(ByVal Index As Long)
Dim i As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String

    If Not IsPlaying(Index) Then Exit Sub
    
    MapNum = GetPlayerMap(Index)
    
    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) Then
            If (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, i).x = GetPlayerX(Index)) Then
                
                    If (MapItem(MapNum, i).y = GetPlayerY(Index)) Then
                    
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
                            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), 0, 0)
                            Call PlayerMsg(Index, Msg, Yellow)
                            Exit For
                        Else
                            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                        
                    End If
                    
                End If
            
            End If
            
        End If
    Next
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        
            i = FindOpenMapItemSlot(GetPlayerMap(Index))
            
            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Dur = 0
                
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                    Case ITEM_TYPE_ARMOR
                        If InvNum = GetPlayerEquipmentSlot(Index, Armor) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Armor)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                    Case ITEM_TYPE_WEAPON
                        If InvNum = GetPlayerEquipmentSlot(Index, Weapon) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                        
                    Case ITEM_TYPE_HELMET
                        If InvNum = GetPlayerEquipmentSlot(Index, Helmet) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                                        
                    Case ITEM_TYPE_SHIELD
                        If InvNum = GetPlayerEquipmentSlot(Index, Shield) Then
                            Call SetPlayerEquipmentSlot(Index, 0, Shield)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                End Select
                                    
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).x = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).y = GetPlayerY(Index)
                            
                If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                        MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                        Call SetPlayerInvItemDur(Index, InvNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), i).Value = Amount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
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
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Amount, MapItem(GetPlayerMap(Index), i).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Else
                Call PlayerMsg(Index, "To many items already on the ground.", BrightRed)
            End If
        End If
    End If
End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
Dim i As Long
Dim expRollover As Long

    ' Check if attacker got a level up
    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        expRollover = CLng(GetPlayerExp(Index) - GetPlayerNextLevel(Index))
        Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                   
        ' Get the amount of skill points to add
        i = (GetPlayerStat(Index, Stats.Speed) \ 10)
        If i < 1 Then i = 1
        If i > 3 Then i = 3
           
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + i)
        Call SetPlayerExp(Index, expRollover)
        Call GlobalMsg(GetPlayerName(Index) & " has gained a level!", Brown)
        Call PlayerMsg(Index, "You have gained a level!  You now have " & GetPlayerPOINTS(Index) & " stat points to distribute.", BrightBlue)
    End If
   
End Sub

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If
    
    Select Case Vital
        Case HP
            i = (GetPlayerStat(Index, Stats.Defense) \ 2)
        Case MP
            i = (GetPlayerStat(Index, Stats.Magic) \ 2)
        Case SP
            i = (GetPlayerStat(Index, Stats.Speed) \ 2)
    End Select
        
    If i < 2 Then i = 2

    GetPlayerVitalRegen = i
End Function

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String
If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Char(TempPlayer(Index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(TempPlayer(Index).CharNum).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(TempPlayer(Index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Char(TempPlayer(Index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Char(TempPlayer(Index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    If Level > MAX_LEVELS Then Exit Sub
    Player(Index).Char(TempPlayer(Index).CharNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = (GetPlayerLevel(Index) + 1) * (GetPlayerStat(Index, Stats.Strength) + GetPlayerStat(Index, Stats.Defense) + GetPlayerStat(Index, Stats.Magic) + GetPlayerStat(Index, Stats.Speed) + GetPlayerPOINTS(Index)) * 25
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(TempPlayer(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Char(TempPlayer(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).Char(TempPlayer(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = Value
    
    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If
    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = 0
    End If
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
Dim CharNum As Long
    
    If Index > MAX_PLAYERS Then Exit Function

    Select Case Vital
        Case HP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + (GetPlayerStat(Index, Stats.Strength) \ 2) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Strength)) * 2
        Case MP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + (GetPlayerStat(Index, Stats.Magic) \ 2) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Magic)) * 2
        Case SP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + (GetPlayerStat(Index, Stats.Speed) \ 2) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Speed)) * 2
    End Select
End Function

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Char(TempPlayer(Index).CharNum).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).Char(TempPlayer(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(Index).Char(TempPlayer(Index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(TempPlayer(Index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).Char(TempPlayer(Index).CharNum).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Char(TempPlayer(Index).CharNum).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Char(TempPlayer(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemNum = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemDur = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(Index).Char(TempPlayer(Index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerEquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipmentSlot = Player(Index).Char(TempPlayer(Index).CharNum).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipmentSlot(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Char(TempPlayer(Index).CharNum).Equipment(EquipmentSlot) = InvNum
End Sub

' ToDo
Sub OnDeath(ByVal Index As Long)
Dim i As Long

    ' Set HP to nothing
    Call SetPlayerVital(Index, Vitals.HP, 0)

    ' Drop all worn items
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipmentSlot(Index, i) > 0 Then
            PlayerMapDropItem Index, GetPlayerEquipmentSlot(Index, i), 0
        End If
    Next
    
    ' Warp player away
    Call PlayerWarp(Index, START_MAP, START_X, START_Y)
    
    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.HP, GetPlayerMaxVital(Index, Vitals.HP))
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerMaxVital(Index, Vitals.MP))
    Call SetPlayerVital(Index, Vitals.SP, GetPlayerMaxVital(Index, Vitals.SP))
    Call SendVital(Index, Vitals.HP)
    Call SendVital(Index, Vitals.MP)
    Call SendVital(Index, Vitals.SP)
    
    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
    End If
End Sub

Sub DamageEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment)
Dim Slot As Long
    
    Slot = GetPlayerEquipmentSlot(Index, EquipmentSlot)
    
    If Slot > 0 Then
        Call SetPlayerInvItemDur(Index, Slot, GetPlayerInvItemDur(Index, Slot) - 1)
            
        If GetPlayerInvItemDur(Index, Slot) <= 0 Then
            Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name) & " has broken.", Yellow)
            Call TakeItem(Index, GetPlayerInvItemNum(Index, Slot), 0)
        Else
            If GetPlayerInvItemDur(Index, Slot) <= 5 Then
                Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name) & " is about to break!", Yellow)
            End If
        End If
    End If
End Sub

