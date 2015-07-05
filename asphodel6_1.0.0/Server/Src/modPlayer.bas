Attribute VB_Name = "modPlayer"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

Public Sub JoinGame(ByVal Index As Long)
Dim i As Long
Dim LoopI As Long

    If Player(Index).Char(TempPlayer(Index).CharNum).Muted Then
        If Player(Index).Char(TempPlayer(Index).CharNum).MuteTime > 30 Then
            Player(Index).Char(TempPlayer(Index).CharNum).MuteTime = CInt((Player(Index).Char(TempPlayer(Index).CharNum).MuteTime - GetTickCountNew) / 60000) + 1
        End If
        If Player(Index).Char(TempPlayer(Index).CharNum).MuteTime < 1 Then
            Player(Index).Char(TempPlayer(Index).CharNum).Muted = False
            Player(Index).Char(TempPlayer(Index).CharNum).MuteTime = 0
        End If
        Player(Index).Char(TempPlayer(Index).CharNum).MuteTime = (Player(Index).Char(TempPlayer(Index).CharNum).MuteTime * 60000) + GetTickCountNew
    End If
    
    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) <= StaffType.Monitor Then
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", Color.White)
    End If
    
    ' Update the log
    UpdatePlayerTable Index
    
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(Index, SLoginOk & SEP_CHAR & Index & END_CHAR)
    
    SendMap Index, GetPlayerMap(Index)
    
    For LoopI = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipmentSlot(Index, LoopI) > 0 Then
            For i = 1 To Stats.Stat_Count - 1
                AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, LoopI))).BuffStats(i)
            Next
            For i = 1 To Vitals.Vital_Count - 1
                AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, LoopI))).BuffVitals(i)
            Next
        End If
    Next
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendSigns(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(Index, i)
    Next
    
    For i = 1 To MAX_ANIMS
        If LenB(Trim$(Animation(i).Name)) > 0 Then SendUpdateAnimTo Index, i
    Next
    
    Call SendStats(Index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    ' Send welcome messages
    Call SendWelcome(Index)
    
    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(Index, SInGame & END_CHAR)
    
End Sub

Public Sub LeftGame(ByVal Index As Long)
Dim n As Long

    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False
       
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = False
        End If
        
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If TempPlayer(Index).InParty = YES Then
            n = TempPlayer(Index).PartyPlayer
            
            Call PlayerMsg(n, GetPlayerName(Index) & " has left " & GAME_NAME & ", disbanning party.", Color.Pink)
            TempPlayer(n).InParty = NO
            TempPlayer(n).PartyPlayer = 0
        End If
        
        If Player(Index).Char(TempPlayer(Index).CharNum).Muted Then
            Player(Index).Char(TempPlayer(Index).CharNum).MuteTime = ((Player(Index).Char(TempPlayer(Index).CharNum).MuteTime - GetTickCountNew) \ 60000) + 1
        End If
        
        Call SavePlayer(Index)
    
        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= StaffType.Monitor Then
            Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has left " & GAME_NAME & "!", Color.White)
        End If
        Call TextAdd(frmServer.txtText, GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".")
        For n = 1 To Vitals.Vital_Count - 1
            VitalBonus(Index, n) = 0
        Next
        For n = 1 To Stats.Stat_Count - 1
            StatBonus(Index, n) = 0
        Next
        Call ClearPlayer(Index)
        Call SendLeftGame(Index)
    End If
    
End Sub

Public Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, Optional ByVal UseSpellTimer As Boolean = False)
Dim Name As String
Dim Exp As Long
Dim n As Long
Dim i As Long
Dim STR As Long
Dim DEF As Long
Dim MapNum As Long
Dim NpcNum As Long

    ' Check for subscript out of range
    If Not IsPlaying(Attacker) Or MapNpcNum <= 0 Or MapNpcNum > UBound(MapSpawn(GetPlayerMap(Attacker)).Npc) Or Damage < 0 Then Exit Sub
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).MapNpc(MapNpcNum).Num
    Name = Trim$(Npc(NpcNum).Name)
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, MapNum, SAttack & SEP_CHAR & Attacker & END_CHAR)
    
    ' Check for weapon
    n = 0
    If GetPlayerEquipmentSlot(Attacker, Weapon) > 0 Then
        n = GetPlayerInvItemNum(Attacker, GetPlayerEquipmentSlot(Attacker, Weapon))
    End If
    
    If Damage >= MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) Then
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points, killing it.", Color.BrightRed)
        Else
            Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points, killing it.", Color.BrightRed)
        End If
        
        If LenB(Trim$(Npc(MapNpc(MapNum).MapNpc(MapNpcNum).Num).Sound(NpcSound.Death_))) > 0 Then SendSound GetPlayerMap(Attacker), Trim$(Npc(MapNpc(MapNum).MapNpc(MapNpcNum).Num).Sound(NpcSound.Death_))
        
        ' Calculate exp to give attacker
        STR = Npc(NpcNum).Stat(Stats.Strength)
        DEF = Npc(NpcNum).Stat(Stats.Defense)
        Exp = Npc(NpcNum).Experience
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then Exp = 1
        
        ' Check if in party, if so divide the exp up by 2
        If TempPlayer(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You have gained " & Exp & " experience points.", Color.BrightBlue)
        Else
            Exp = Exp * 0.5
            
            If Exp < 0 Then
                Exp = 1
            End If
            
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You have gained " & Exp & " party experience points.", Color.BrightBlue)
            
            n = TempPlayer(Attacker).PartyPlayer
            If n > 0 Then
                Call SetPlayerExp(n, GetPlayerExp(n) + Exp)
                Call PlayerMsg(n, "You have gained " & Exp & " party experience points.", Color.BrightBlue)
            End If
        End If
        
        SendNPCVital MapNum, MapNpcNum
        
        ' Drop the goods if they get it
        n = Random(1, Npc(NpcNum).DropChance)
        If n = Random(1, Npc(NpcNum).DropChance) Then
            Call SpawnItem(Npc(NpcNum).DropItem, Npc(NpcNum).DropItemValue, MapNum, MapNpc(MapNum).MapNpc(MapNpcNum).X, MapNpc(MapNum).MapNpc(MapNpcNum).Y)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).MapNpc(MapNpcNum).Num = 0
        MapNpc(MapNum).MapNpc(MapNpcNum).SpawnWait = GetTickCountNew + (Npc(MapSpawn(MapNum).Npc(NpcNum).Num).SpawnSecs * 1000)
        MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) = 0
        Call SendDataToMap(MapNum, SNpcDead & SEP_CHAR & MapNpcNum & END_CHAR)
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check for level up party member
        If TempPlayer(Attacker).InParty = YES Then
            Call CheckPlayerLevelUp(TempPlayer(Attacker).PartyPlayer)
        End If
    
        ' Check if target is npc that died and if so set target to 0
        If TempPlayer(Attacker).TargetType = E_Target.NPC_ Then
            If TempPlayer(Attacker).Target = MapNpcNum Then
                TempPlayer(Attacker).Target = 0
                TempPlayer(Attacker).TargetType = E_Target.None
            End If
        End If
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) - Damage
        
        SendNPCVital MapNum, MapNpcNum
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", Color.White)
        Else
            Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", Color.White)
        End If
        
        ' Check if we should send a message
        If MapNpc(MapNum).MapNpc(MapNpcNum).Target = 0 Then
            If LenB(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                Call PlayerMsg(Attacker, "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
            End If
        End If
        
        ' Set the NPC target to the player
        MapNpc(MapNum).MapNpc(MapNpcNum).Target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).MapNpc(MapNpcNum).Num).Behavior = NPC_Behavior.Guard Then
            For i = 1 To UBound(MapSpawn(MapNum).Npc)
                If MapNpc(MapNum).MapNpc(i).Num = MapNpc(MapNum).MapNpc(MapNpcNum).Num Then
                    MapNpc(MapNum).MapNpc(i).Target = Attacker
                End If
            Next
        End If
    End If
    
    ' Reduce durability of weapon
    Call DamageEquipment(Attacker, Weapon)
    
    ' Reset attack timer
    If Not UseSpellTimer Then TempPlayer(Attacker).AttackTimer = GetTickCountNew
    
End Sub

Public Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal UseSpellTimer As Boolean = False)
Dim Exp As Long
Dim n As Long

    ' Check for subscript out of range
    If Not IsPlaying(Attacker) Or Not IsPlaying(Victim) Or Damage < 0 Then Exit Sub
    
    ' Check for weapon
    n = 0
    
    If GetPlayerEquipmentSlot(Attacker, Weapon) > 0 Then n = GetPlayerInvItemNum(Attacker, GetPlayerEquipmentSlot(Attacker, Weapon))
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), SAttack & SEP_CHAR & Attacker & END_CHAR)
    
    ' reduce dur. on victims equipment
    Call DamageEquipment(Victim, Armor)
    Call DamageEquipment(Victim, Helmet)
    
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", Color.White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", Color.BrightRed)
        Else
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", Color.White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", Color.BrightRed)
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), Color.BrightRed)
        
        ' Calculate exp to give attacker
        Exp = Int(GetPlayerExp(Victim) * 0.1)
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 0
        End If
        
        If Exp = 0 Then
            Call PlayerMsg(Victim, "You lost no experience points.", Color.BrightRed)
            Call PlayerMsg(Attacker, "You received no experience points from that weak insignificant player.", Color.BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            Call PlayerMsg(Victim, "You lost " & Exp & " experience points.", Color.BrightRed)
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            Call PlayerMsg(Attacker, "You got " & Exp & " experience points for killing " & GetPlayerName(Victim) & ".", Color.BrightBlue)
        End If
        
        ' Check for a level up
        Call CheckPlayerLevelUp(Attacker)
        
        ' Check if target is player who died and if so set target to 0
        If TempPlayer(Attacker).TargetType = E_Target.Player_ Then
            If TempPlayer(Attacker).Target = Victim Then
                TempPlayer(Attacker).Target = 0
                TempPlayer(Attacker).TargetType = E_Target.None
            End If
        End If
        
        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!", Color.BrightRed)
            End If
        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!", Color.BrightRed)
        End If
        
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' Check for a weapon and say damage
        If n = 0 Then
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", Color.White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", Color.BrightRed)
        Else
            Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", Color.White)
            Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", Color.BrightRed)
        End If
    End If
    
    ' Reduce durability of weapon
    Call DamageEquipment(Attacker, Weapon)
    
    ' Reset attack timer
    If Not UseSpellTimer Then TempPlayer(Attacker).AttackTimer = GetTickCountNew
    
End Sub

Public Sub DirectDamagePlayer(ByVal Index As Long, ByVal Damage As Long, Optional ByVal Message As String = "default")
Dim Exp As Long

    If Message = "default" Then Message = "You have taken " & Damage & " points of damage!"
    
    PlayerMsg Index, Message, Color.BrightRed
    
    If Damage >= GetPlayerVital(Index, Vitals.HP) Then
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Index) & " has been killed!", Color.BrightRed)
        
        ' Calculate exp to lose
        Exp = Int(GetPlayerExp(Index) * 0.1)
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then Exp = 0
        
        If Exp = 0 Then
            Call PlayerMsg(Index, "You lost no experience points.", Color.BrightRed)
        Else
            Call SetPlayerExp(Index, GetPlayerExp(Index) - Exp)
            Call PlayerMsg(Index, "You lost " & Exp & " experience points.", Color.BrightRed)
        End If
        
        If GetPlayerPK(Index) = YES Then
            Call GlobalMsg(GetPlayerName(Index) & " has paid the price for being a Player Killer!", Color.BrightRed)
        End If
        
        Call OnDeath(Index)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) - Damage)
        Call SendVital(Index, Vitals.HP)
    End If
    
End Sub

Public Function GetPlayerStat_withBonus(ByVal Index As Long, ByVal Stat As Stats)
Dim LoopI As Long

    GetPlayerStat_withBonus = GetPlayerStat(Index, Stat) + StatBonus(Index, Stat)
    
    'For LoopI = 1 To Equipment.Equipment_Count - 1
    '    If GetPlayerEquipmentSlot(Index, LoopI) > 0 Then
    '        If Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, LoopI))).BuffStats(Stat) > 0 Then
    '            GetPlayerStat_withBonus = GetPlayerStat_withBonus + Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, LoopI))).BuffStats(Stat)
    '        End If
    '    End If
    'Next
    
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long

    GetPlayerDamage = 0
    
    ' Check for subscript out of range
    If Not IsPlaying(Index) Or Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerDamage = Int(GetPlayerStat_withBonus(Index, Stats.Strength) * 0.5)
    
    If GetPlayerDamage <= 0 Then GetPlayerDamage = 1
    
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
Dim ArmorSlot As Long
Dim HelmSlot As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Or Index <= 0 Or Index > MAX_PLAYERS Then Exit Function
    
    ArmorSlot = GetPlayerEquipmentSlot(Index, Armor)
    HelmSlot = GetPlayerEquipmentSlot(Index, Helmet)
    
    GetPlayerProtection = Int(GetPlayerStat_withBonus(Index, Stats.Defense) * 0.2)
    
End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
Dim i As Long
Dim n As Long

    If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
        n = Random(1, 2)
        If n = 1 Then
            i = Int(GetPlayerStat_withBonus(Index, Stats.Strength) * 0.5) + Int(GetPlayerLevel(Index) * 0.5)
            
            n = Random(1, 100)
            CanPlayerCriticalHit = (n <= i)
        End If
    End If
    
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
Dim i As Long
Dim n As Long
Dim ShieldSlot As Long
    
    ShieldSlot = GetPlayerEquipmentSlot(Index, Shield)
    
    If ShieldSlot > 0 Then
        n = Random(1, 2)
        If n = 1 Then
            i = Int(GetPlayerStat_withBonus(Index, Stats.Defense) * 0.5) + Int(GetPlayerLevel(Index) * 0.5)
        
            n = Random(1, 100)
            CanPlayerBlockHit = (n <= i)
        End If
    End If
End Function

Public Function CanCastSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Boolean

    ' Check if timer is ok
    If GetTickCountNew < TempPlayer(Index).CastTimer(SpellSlot) + Spell(GetPlayerSpell(Index, SpellSlot)).Timer Then Exit Function
    
    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Function
    
    ' Make sure player has the spell
    If Not HasSpell(Index, GetPlayerSpell(Index, SpellSlot)) Then
        Call PlayerMsg(Index, "You do not have this spell!", Color.BrightRed)
        Exit Function
    End If
    
    ' Check if they have enough MP
    If GetPlayerVital_withBonus(Index, Vitals.MP) < Spell(GetPlayerSpell(Index, SpellSlot)).MPReq Then
        Call PlayerMsg(Index, "Not enough mana points!", Color.BrightRed)
        Exit Function
    End If
    
    CanCastSpell = True
    
End Function

Public Function CastSpell(ByVal Index As Long, ByVal SpellSlot As Long, Optional ByVal IgnoreMessage As Boolean = False) As Boolean
Dim SpellNum As Long
Dim MPReq As Long
Dim i As Long
Dim n As Long
Dim Damage As Long
Dim Casted As Boolean
Dim CanCast As Boolean
Dim ErrorMessage As String
Dim TargetType As Byte
Dim TargetName As String

    Damage = -1
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    MPReq = Spell(SpellNum).MPReq
    ErrorMessage = "Could not cast spell!"
    
    ' Check if the spell is a give item and do that instead of a stat modification
    If Spell(SpellNum).Type = Spell_Type.GiveItem_ Then
        If Spell(SpellNum).AOE = 0 Then
            n = FindOpenInvSlot(Index, Spell(SpellNum).Data1)
            If n > 0 Then
                Call GiveItem(Index, Spell(SpellNum).Data1, Spell(SpellNum).Data2)
                Casted = True
            Else
                Call PlayerMsg(Index, "Your inventory is full!", Color.BrightRed)
            End If
            Exit Function
        End If
    End If
    
    n = TempPlayer(Index).Target
    TargetType = TempPlayer(Index).TargetType
    
    Select Case TargetType
        Case E_Target.Player_
            
            If IsPlaying(n) Then
                If GetPlayerVital(n, Vitals.HP) > 0 Then
                    If GetPlayerMap(Index) = GetPlayerMap(n) Then
                        If IsWithinPVPLimit(Index, n) Then
                            If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Then
                                If Not AdminSafety(Index, n) Then
                                    If n <> Index Then
                                        CanCast = True
                                    Else
                                        ErrorMessage = "You cannot attack yourself!"
                                    End If
                                End If
                            Else
                                ErrorMessage = "This is a safe map!"
                            End If
                        End If
                    Else
                        ErrorMessage = "You aren't on the same map as the other player!"
                    End If
                End If
                
                If Spell(SpellNum).Range > 0 Then
                    If Not IsInRange(GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(n), GetPlayerY(n), Spell(SpellNum).Range) Then
                        CanCast = False
                        ErrorMessage = "You are not in range of the target!"
                    End If
                End If
                
                TargetName = GetPlayerName(n)
                
                If Spell(SpellNum).Type = Spell_Type.SubHP_ Or _
                   Spell(SpellNum).Type = Spell_Type.SubMP_ Or _
                   Spell(SpellNum).Type = Spell_Type.SubSP_ Then
                    
                    If CanCast Then
                        Select Case Spell(SpellNum).Type
                            Case Spell_Type.SubHP_
                                Damage = (GetPlayerStat_withBonus(Index, Stats.Magic) \ 4) + Spell(SpellNum).Data1 - GetPlayerProtection(n)
                                If Damage < 0 Then Damage = 0
                                
                            Case Spell_Type.SubMP_
                                Call SetPlayerVital(n, Vitals.MP, GetPlayerVital(n, Vitals.MP) - Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.MP)
                                
                            Case Spell_Type.SubSP_
                                Call SetPlayerVital(n, Vitals.SP, GetPlayerVital(n, Vitals.SP) - Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.SP)
                        End Select
                        
                        Casted = True
                        
                    End If
                    
                ElseIf Spell(SpellNum).Type = Spell_Type.AddHP_ Or _
                       Spell(SpellNum).Type = Spell_Type.AddMP_ Or _
                       Spell(SpellNum).Type = Spell_Type.AddSP_ Then
                    
                    If GetPlayerMap(Index) = GetPlayerMap(n) Then CanCast = True
                    
                    If Index = n Then CanCast = True
                    
                    If Spell(SpellNum).Range > 0 Then
                        If Not IsInRange(GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(n), GetPlayerY(n), Spell(SpellNum).Range) Then
                            CanCast = False
                            ErrorMessage = "You are not in range of the target!"
                        End If
                    End If
                    
                    If CanCast Then
                        Select Case Spell(SpellNum).Type
                            Case Spell_Type.AddHP_
                                Call SetPlayerVital(n, Vitals.HP, GetPlayerVital(n, Vitals.HP) + Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.HP)
                                
                            Case Spell_Type.AddMP_
                                Call SetPlayerVital(n, Vitals.MP, GetPlayerVital(n, Vitals.MP) + Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.MP)
                                
                            Case Spell_Type.AddSP_
                                Call SetPlayerVital(n, Vitals.SP, GetPlayerVital(n, Vitals.SP) + Spell(SpellNum).Data1)
                                Call SendVital(n, Vitals.SP)
                        End Select
                        
                        Casted = True
                    End If
                    
                End If
            End If
            
        Case E_Target.NPC_
        
            If Npc(MapNpc(GetPlayerMap(Index)).MapNpc(n).Num).Behavior <> NPC_Behavior.Friendly Then
                If Npc(MapNpc(GetPlayerMap(Index)).MapNpc(n).Num).Behavior <> NPC_Behavior.ShopKeeper Then
                    CanCast = True
                Else
                    ErrorMessage = "You cannot attack a shop keeper!"
                End If
            Else
                ErrorMessage = "You cannot attack the friendly " & Trim$(Npc(MapNpc(GetPlayerMap(Index)).MapNpc(n).Num).Name) & "!"
            End If
            
            If Spell(SpellNum).Range > 0 Then
                If Not IsInRange(GetPlayerX(Index), GetPlayerY(Index), MapNpc(GetPlayerMap(Index)).MapNpc(n).X, MapNpc(GetPlayerMap(Index)).MapNpc(n).Y, Spell(SpellNum).Range) Then
                    CanCast = False
                    ErrorMessage = "You are not in range of the target!"
                End If
            End If
            
            TargetName = Trim$(Npc(MapNpc(GetPlayerMap(Index)).MapNpc(n).Num).Name)
            
            If CanCast Then
                Select Case Spell(SpellNum).Type
                    'Case Spell_Type.AddHP_
                        'MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.HP) = MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.HP) + Spell(SpellNum).Data1
                        
                    Case Spell_Type.SubHP_
                    
                        Damage = (GetPlayerStat_withBonus(Index, Stats.Magic) \ 4) + Spell(SpellNum).Data1 - (Npc(MapNpc(GetPlayerMap(Index)).MapNpc(n).Num).Stat(Stats.Defense) \ 2)
                        If Damage < 0 Then Damage = 0
                        
                        Casted = True
                        
                    'Case Spell_Type.AddMP_
                        'MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.MP) = MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.MP) + Spell(SpellNum).Data1
                        
                    'Case Spell_Type.SubMP_
                        'MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.MP) = MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.MP) - Spell(SpellNum).Data1
                        
                    'Case Spell_Type.AddSP_
                        'MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.SP) = MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.SP) + Spell(SpellNum).Data1
                        
                    'Case Spell_Type.SubSP_
                        'MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.SP) = MapNpc(GetPlayerMap(Index)).MapNpc(n).Vital(Vitals.SP) - Spell(SpellNum).Data1
                        
                End Select
            End If
            
        Case Else
            
            ErrorMessage = "You have no target!"
            
    End Select
    
    If Casted Then
        If LenB(Trim$(Spell(SpellNum).CastSound)) > 0 Then SendSound GetPlayerMap(Index), Trim$(Spell(SpellNum).CastSound)
        'Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & TargetName & ".", Color.BrightBlue)
        If Spell(SpellNum).Anim > 0 Then SendDataToMap GetPlayerMap(Index), SAnimation & SEP_CHAR & Spell(SpellNum).Anim & SEP_CHAR & n & SEP_CHAR & TargetType & END_CHAR
        
        If Damage <> -1 Then
            Select Case TargetType
            
                Case E_Target.Player_
                    If Damage > 0 Then
                        Call AttackPlayer(Index, n, Damage, True)
                        TempPlayer(Index).CastTimer(SpellSlot) = GetTickCountNew
                    Else
                        If Not IgnoreMessage Then Call PlayerMsg(Index, "The spell was too weak to hurt " & TargetName & "!", Color.BrightRed)
                    End If
                    
                Case E_Target.NPC_
                    If Damage > 0 Then
                        Call AttackNpc(Index, n, Damage, True)
                        TempPlayer(Index).CastTimer(SpellSlot) = GetTickCountNew
                        ' if they didn't kill the NPC, then do the reflection
                        If MapNpc(GetPlayerMap(Index)).MapNpc(n).Num > 0 Then
                            If Npc(MapNpc(GetPlayerMap(Index)).MapNpc(n).Num).Reflection(NPC_Reflection.Magic_) > 0 Then
                                Damage = Damage * (Npc(MapNpc(GetPlayerMap(Index)).MapNpc(n).Num).Reflection(NPC_Reflection.Magic_) * 0.01)
                                NpcAttackPlayer n, Index, Damage, True
                            End If
                        End If
                    Else
                        If Not IgnoreMessage Then Call PlayerMsg(Index, "The spell was too weak to hurt " & TargetName & "!", Color.BrightRed)
                    End If
                    
            End Select
        End If
    Else
        ' send cast success with an extra parse to tell the client it failed in case we missed it
        SendDataTo Index, SCastSuccess & SEP_CHAR & SpellSlot & SEP_CHAR & 0 & END_CHAR
        If Not IgnoreMessage Then Call PlayerMsg(Index, ErrorMessage, Color.BrightRed)
    End If
    
    CastSpell = Casted
    
End Function

Public Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim OldMap As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Or MapNum < 1 Or MapNum > MAX_MAPS Then Exit Sub
    
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = E_Target.None
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    
    If OldMap <> MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If
    
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)
    
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then PlayersOnMap(OldMap) = False
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = True
    
    TempPlayer(Index).GettingMap = YES
    Call SendDataTo(Index, SCheckForMap & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)
    
End Sub

Public Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String
Dim MapNum As Long
Dim X As Long
Dim Y As Long
Dim Moved As Boolean

    ' Check for subscript out of range
    If Not IsPlaying(Index) Or Dir < E_Direction.Up_ Or Dir > E_Direction.Right_ Or Movement < 1 Or Movement > 2 Then Exit Sub
    
    Call SetPlayerDir(Index, Dir)
    
    Select Case Dir
        Case E_Direction.Up_
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> Tile_Type.Blocked_ Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> Tile_Type.Key_ Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = Tile_Type.Key_ And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                        
                        Packet = SPlayerMove & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = True
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Up > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), MAX_MAPY)
                    Moved = True
                End If
            End If
                    
        Case E_Direction.Down_
            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < MAX_MAPY Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> Tile_Type.Blocked_ Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> Tile_Type.Key_ Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = Tile_Type.Key_ And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                        Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                        
                        Packet = SPlayerMove & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = True
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Down > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    Moved = True
                End If
            End If
        
        Case E_Direction.Left_
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> Tile_Type.Blocked_ Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> Tile_Type.Key_ Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = Tile_Type.Key_ And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                        
                        Packet = SPlayerMove & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = True
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Left > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, MAX_MAPX, GetPlayerY(Index))
                    Moved = True
                End If
            End If
        
        Case E_Direction.Right_
            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < MAX_MAPX Then
                ' Check to make sure that the tile is walkable
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> Tile_Type.Blocked_ Then
                    ' Check to see if the tile is a key and if it is check if its opened
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> Tile_Type.Key_ Or (Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = Tile_Type.Key_ And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                        Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                        
                        Packet = SPlayerMove & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & Movement & END_CHAR
                        Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                        Moved = True
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(Index)).Right > 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Moved = True
                End If
            End If
    End Select
    
    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = Tile_Type.Warp_ Then
        MapNum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3
                        
        Call PlayerWarp(Index, MapNum, X, Y)
        Moved = True
    End If
    
    ' Check to see if the tile is a shop
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = Tile_Type.Shop_ Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 > 0 Then
            SendTrade Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        Else
            PlayerMsg Index, "There is no shop here!", Color.BrightRed
        End If
        Moved = True
    End If
    
    ' Check to see if the tile is a heal
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = Tile_Type.Heal_ Then
        For X = 1 To Vitals.Vital_Count - 1
            SetPlayerVital Index, X, GetPlayerMaxVital(Index, X)
            SendVital Index, X
        Next
        PlayerMsg Index, "You have been fully healed!", Color.Green
        Moved = True
    End If
    
    ' Check for guild making tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = Tile_Type.Guild_ Then
        Dim GuildIndex As Long
        Dim LoopI As Long
        
        If CanMakeGuild(Index, vbNullString) Then SendDataTo Index, SGuildCreation & END_CHAR
        Moved = True
    End If
    
    ' Check for key trigger open
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = Tile_Type.KeyOpen_ Then
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        
        If Map(GetPlayerMap(Index)).Tile(X, Y).Type = Tile_Type.Key_ Then
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCountNew + 5000
                
                Call SendDataToMap(GetPlayerMap(Index), SMapKey & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", Color.White)
            End If
        End If
        Moved = True
    End If
    
    ' They tried to hack (or it could just be lag)
    'If Not Moved Then Call HackingAttempt(Index, "Position Modification")
    If Not Moved Then SendPlayerData Index
    
End Sub

Public Sub CheckEquippedItems(ByVal Index As Long)
Dim Slot As Long
Dim ItemNum As Long
Dim i As Long
Dim LoopI As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        Slot = GetPlayerEquipmentSlot(Index, i)
        If Slot > 0 Then
            ItemNum = GetPlayerInvItemNum(Index, Slot)
            
            If ItemNum > 0 Then
                Select Case i
                    Case Equipment.Weapon
                        If Item(ItemNum).Type <> ItemType.Weapon_ Then
                            For LoopI = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffStats(LoopI), False
                            Next
                            For LoopI = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffVitals(LoopI), False
                            Next
                            SetPlayerEquipmentSlot Index, 0, i
                        End If
                    Case Equipment.Armor
                        If Item(ItemNum).Type <> ItemType.Armor_ Then
                            For LoopI = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffStats(LoopI), False
                            Next
                            For LoopI = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffVitals(LoopI), False
                            Next
                            SetPlayerEquipmentSlot Index, 0, i
                        End If
                    Case Equipment.Helmet
                        If Item(ItemNum).Type <> ItemType.Helmet_ Then
                            For LoopI = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffStats(LoopI), False
                            Next
                            For LoopI = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffVitals(LoopI), False
                            Next
                            SetPlayerEquipmentSlot Index, 0, i
                        End If
                    Case Equipment.Shield
                        If Item(ItemNum).Type <> ItemType.Shield_ Then
                            For LoopI = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffStats(LoopI), False
                            Next
                            For LoopI = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffVitals(LoopI), False
                            Next
                            SetPlayerEquipmentSlot Index, 0, i
                        End If
                End Select
            Else
                For LoopI = 1 To Stats.Stat_Count - 1
                    AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, i))).BuffStats(LoopI), False
                Next
                For LoopI = 1 To Vitals.Vital_Count - 1
                    AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, i))).BuffVitals(LoopI), False
                Next
                SetPlayerEquipmentSlot Index, 0, i
            End If
        End If
    Next
    
End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
    
    If Item(ItemNum).Type = ItemType.Currency_ Then
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
    If Not IsPlaying(Index) Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ItemType.Currency_ Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next
    
End Function

Public Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long
Dim LoopI As Long
Dim n As Long
Dim TakeItem As Boolean

    ' Check for subscript out of range
    If Not IsPlaying(Index) Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ItemType.Currency_ Then
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
                    Case ItemType.Weapon_
                        If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Weapon) Then
                                For LoopI = 1 To Stats.Stat_Count - 1
                                    AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffStats(LoopI), False
                                Next
                                For LoopI = 1 To Vitals.Vital_Count - 1
                                    AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffVitals(LoopI), False
                                Next
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
                
                    Case ItemType.Armor_
                        If GetPlayerEquipmentSlot(Index, Armor) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Armor) Then
                                For LoopI = 1 To Stats.Stat_Count - 1
                                    AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffStats(LoopI), False
                                Next
                                For LoopI = 1 To Vitals.Vital_Count - 1
                                    AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffVitals(LoopI), False
                                Next
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
                    
                    Case ItemType.Helmet_
                        If GetPlayerEquipmentSlot(Index, Helmet) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Helmet) Then
                                For LoopI = 1 To Stats.Stat_Count - 1
                                    AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffStats(LoopI), False
                                Next
                                For LoopI = 1 To Vitals.Vital_Count - 1
                                    AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffVitals(LoopI), False
                                Next
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
                    
                    Case ItemType.Shield_
                        If GetPlayerEquipmentSlot(Index, Shield) > 0 Then
                            If i = GetPlayerEquipmentSlot(Index, Shield) Then
                                For LoopI = 1 To Stats.Stat_Count - 1
                                    AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffStats(LoopI), False
                                Next
                                For LoopI = 1 To Vitals.Vital_Count - 1
                                    AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffVitals(LoopI), False
                                Next
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
                If (n <> ItemType.Weapon_) Then
                    If (n <> ItemType.Armor_) Then
                        If (n <> ItemType.Helmet_) Then
                            If (n <> ItemType.Shield_) Then
                                TakeItem = True
                            End If
                        End If
                    End If
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

Public Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    i = FindOpenInvSlot(Index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ItemType.Armor_) Or (Item(ItemNum).Type = ItemType.Weapon_) Or (Item(ItemNum).Type = ItemType.Helmet_) Or (Item(ItemNum).Type = ItemType.Shield_) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Durability)
        End If
        
        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, "Your inventory is full.", Color.BrightRed)
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

Public Sub PlayerMapGetItem(ByVal Index As Long)
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
                If (MapItem(MapNum, i).X = GetPlayerX(Index)) Then
                
                    If (MapItem(MapNum, i).Y = GetPlayerY(Index)) Then
                    
                        ' Find open slot
                        n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
                        
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
                            If Item(GetPlayerInvItemNum(Index, n)).Type = ItemType.Currency_ Then
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
                            MapItem(MapNum, i).Anim = 0
                            MapItem(MapNum, i).X = 0
                            MapItem(MapNum, i).Y = 0
                            
                            Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                            Call PlayerMsg(Index, Msg, Color.Yellow)
                            Exit For
                        Else
                            Call PlayerMsg(Index, "Your inventory is full.", Color.BrightRed)
                            Exit For
                        End If
                        
                    End If
                    
                End If
            
            End If
            
        End If
    Next
    
End Sub

Public Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Ammount As Long)
Dim i As Long
Dim LoopI As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Or InvNum <= 0 Or InvNum > MAX_INV Then Exit Sub
    
    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        
            i = FindOpenMapItemSlot(GetPlayerMap(Index))
            
            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Dur = 0
                
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                    Case ItemType.Armor_
                        If InvNum = GetPlayerEquipmentSlot(Index, Armor) Then
                            For LoopI = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffStats(LoopI), False
                            Next
                            For LoopI = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffVitals(LoopI), False
                            Next
                            Call SetPlayerEquipmentSlot(Index, 0, Armor)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                    Case ItemType.Weapon_
                        If InvNum = GetPlayerEquipmentSlot(Index, Weapon) Then
                            For LoopI = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffStats(LoopI), False
                            Next
                            For LoopI = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffVitals(LoopI), False
                            Next
                            Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                        
                    Case ItemType.Helmet_
                        If InvNum = GetPlayerEquipmentSlot(Index, Helmet) Then
                            For LoopI = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffStats(LoopI), False
                            Next
                            For LoopI = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffVitals(LoopI), False
                            Next
                            Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                                        
                    Case ItemType.Shield_
                        If InvNum = GetPlayerEquipmentSlot(Index, Shield) Then
                            For LoopI = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffStats(LoopI), False
                            Next
                            For LoopI = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, LoopI, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffVitals(LoopI), False
                            Next
                            Call SetPlayerEquipmentSlot(Index, 0, Shield)
                            Call SendWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                End Select
                
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, InvNum)
                MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).Y = GetPlayerY(Index)
                MapItem(GetPlayerMap(Index), i).Anim = Item(GetPlayerInvItemNum(Index, InvNum)).Anim
                
                If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ItemType.Currency_ Then
                    ' Check if its more then they have and if so drop it all
                    If Ammount >= GetPlayerInvItemValue(Index, InvNum) Then
                        MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Color.Yellow)
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                        Call SetPlayerInvItemDur(Index, InvNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), i).Value = Ammount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Ammount & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Color.Yellow)
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Ammount)
                    End If
                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), i).Value = 0
                    If Item(GetPlayerInvItemNum(Index, InvNum)).Type >= ItemType.Weapon_ And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ItemType.Shield_ Then
                        If GetPlayerInvItemDur(Index, InvNum) > 0 Then
                            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " (" & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Durability & ").", Color.Yellow)
                        Else
                            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " (End.).", Color.Yellow)
                        End If
                    Else
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Color.Yellow)
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
                Call PlayerMsg(Index, "Too many items already on the ground.", Color.BrightRed)
            End If
        End If
    End If
End Sub

Public Sub CheckPlayerLevelUp(ByVal Index As Long)
Dim i As Long
Dim expRollover As Long
Dim LevelCount As Byte

    If Not GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then Exit Sub
    If GetPlayerLevel(Index) >= MAX_LEVELS Then Exit Sub
    
    ' Check if attacker got a level up
    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        expRollover = CLng(GetPlayerExp(Index) - GetPlayerNextLevel(Index))
        Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
        
        ' Get the ammount of skill points to add
        i = Int(GetPlayerStat_withBonus(Index, Stats.SPEED) * 0.1)
        If i < 1 Then i = 1
        If i > 3 Then i = 3
        
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + Class(GetPlayerClass(Index)).PointsPerLevel)
        Call SetPlayerExp(Index, expRollover)
        LevelCount = LevelCount + 1
    Loop
    
    Call GlobalMsg(GetPlayerName(Index) & " has gained " & LevelCount & " level(s)!", Color.Brown)
    Call PlayerMsg(Index, "You have gained " & LevelCount & " level(s)!  You now have " & GetPlayerPOINTS(Index) & " stat points to distribute.", Color.BrightBlue)
    
End Sub

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
Dim i As Long

    ' Prevent subscript out of range
    If Not IsPlaying(Index) Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If
    
    Select Case Vital
        Case HP
            i = Int(GetPlayerStat_withBonus(Index, Stats.Defense) * 0.5)
        Case MP
            i = Int(GetPlayerStat_withBonus(Index, Stats.Magic) * 0.5)
        Case SP
            i = Int(GetPlayerStat_withBonus(Index, Stats.SPEED) * 0.5)
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

Public Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Public Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(TempPlayer(Index).CharNum).Name)
End Function

Public Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(TempPlayer(Index).CharNum).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(TempPlayer(Index).CharNum).Class
End Function

Public Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(TempPlayer(Index).CharNum).Sprite
End Function

Public Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(TempPlayer(Index).CharNum).Level
End Function

Public Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    If Level > MAX_LEVELS Then Exit Sub
    Player(Index).Char(TempPlayer(Index).CharNum).Level = Level
    SendDataTo Index, SPlayerLevel & SEP_CHAR & Level & END_CHAR
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = (GetPlayerLevel(Index) + 1) * (GetPlayerStat(Index, Stats.Strength) + GetPlayerStat(Index, Stats.Defense) + GetPlayerStat(Index, Stats.Magic) + GetPlayerStat(Index, Stats.SPEED) + GetPlayerPOINTS(Index)) * 25
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(TempPlayer(Index).CharNum).Exp
End Function

Public Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)

    If Index < 1 Then Exit Sub
    
    Player(Index).Char(TempPlayer(Index).CharNum).Exp = Exp
    
    If IsPlaying(Index) Then SendDataTo Index, SExpUpdate & SEP_CHAR & Exp & SEP_CHAR & GetPlayerNextLevel(Index) & END_CHAR
    
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(TempPlayer(Index).CharNum).Access
End Function

Public Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(TempPlayer(Index).CharNum).PK
End Function

Public Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    GetPlayerVital = Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital)
End Function

Public Function GetPlayerVital_withBonus(ByVal Index As Long, ByVal Vital As Vitals)
Dim LoopI As Long

    For LoopI = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipmentSlot(Index, LoopI) > 0 Then
            If Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, LoopI))).BuffVitals(Vital) > 0 Then
                GetPlayerVital_withBonus = Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, LoopI))).BuffVitals(Vital)
            End If
        End If
    Next
    
End Function

Public Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)

    Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = Value
    
    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If
    
    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = 0
    End If
    
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals, Optional Without As Boolean = False) As Long
Dim CharNum As Long

    Select Case Vital
        Case HP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + Int(GetPlayerStat_withBonus(Index, Stats.Strength) * 0.5) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Strength)) * 2
        Case MP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + Int(GetPlayerStat_withBonus(Index, Stats.Magic) * 0.5) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Magic)) * 2
        Case SP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + Int(GetPlayerStat_withBonus(Index, Stats.SPEED) * 0.5) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.SPEED)) * 2
    End Select
    
    If Without Then Exit Function
    
    GetPlayerMaxVital = GetPlayerMaxVital + GetPlayerVital_withBonus(Index, Vital)
    
End Function

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    GetPlayerStat = Player(Index).Char(TempPlayer(Index).CharNum).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(TempPlayer(Index).CharNum).POINTS
End Function

Public Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).POINTS = POINTS
    SendDataTo Index, SPlayerPoints & SEP_CHAR & POINTS & END_CHAR
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(TempPlayer(Index).CharNum).Map
End Function

Public Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 Then
        If MapNum <= MAX_MAPS Then
            Player(Index).Char(TempPlayer(Index).CharNum).Map = MapNum
        End If
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(TempPlayer(Index).CharNum).X
End Function

Public Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(TempPlayer(Index).CharNum).Y
End Function

Public Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(TempPlayer(Index).CharNum).Dir
End Function

Public Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Num
End Function

Public Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Value
End Function

Public Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Dur
End Function

Public Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(TempPlayer(Index).CharNum).Spell(SpellSlot)
End Function

Public Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerEquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
    GetPlayerEquipmentSlot = Player(Index).Char(TempPlayer(Index).CharNum).Equipment(EquipmentSlot)
End Function

Public Sub SetPlayerEquipmentSlot(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Char(TempPlayer(Index).CharNum).Equipment(EquipmentSlot) = InvNum
End Sub

Public Sub OnDeath(ByVal Index As Long)
Dim i As Long

    ' Set HP to nothing
    Call SetPlayerVital(Index, Vitals.HP, 0)
    
    ' Drop all worn items
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipmentSlot(Index, i) > 0 Then
            '25% chance of drop
            If Random(1, 4) = 2 Then PlayerMapDropItem Index, GetPlayerEquipmentSlot(Index, i), 0
        End If
    Next
    
    ' Warp player away
    Call PlayerWarp(Index, Class(GetPlayerClass(Index)).StartLoc.MapNum, Class(GetPlayerClass(Index)).StartLoc.X, Class(GetPlayerClass(Index)).StartLoc.Y)
    
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

Public Sub DamageEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment)
Dim Slot As Long

    Slot = GetPlayerEquipmentSlot(Index, EquipmentSlot)
    
    If Slot > 0 Then
        If GetPlayerInvItemDur(Index, Slot) = -1 Then Exit Sub
        
        Call SetPlayerInvItemDur(Index, Slot, GetPlayerInvItemDur(Index, Slot) - 1)
        
        If GetPlayerInvItemDur(Index, Slot) < 1 Then
            Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name) & " has broken.", Color.Yellow)
            Call TakeItem(Index, GetPlayerInvItemNum(Index, Slot), 0)
        Else
            If GetPlayerInvItemDur(Index, Slot) <= 5 Then
                Call PlayerMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, Slot)).Name) & " is about to break!", Color.Yellow)
            End If
        End If
    End If
    
End Sub
