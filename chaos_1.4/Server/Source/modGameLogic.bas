Attribute VB_Name = "modGameLogic"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

Sub AddToGrid(ByVal NewMap, _
   ByVal NewX, _
   ByVal NewY)
    Grid(NewMap).Loc(NewX, NewY).Blocked = True
End Sub

Sub AttackNpc(ByVal Attacker As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim N As Long, i As Long, X As Long, o As Long
Dim MapNum As Long, npcnum As Long
Dim AP As Long
Dim StamRemove As Long
Dim LB As Long, SB As Long, BW As Long, PA As Long, AA As Long, TW As Long, xb As Long, MW As Long, CS As Long, BA As Long

If GetPlayerFP(Attacker) < 11 Then
    Call PlayerMsg(Attacker, "Your Hunger Level Is Low, You Need To Eat In Order To Have Strength !", BrightRed)
    Exit Sub
    End If

If GetPlayerWeaponSlot(Attacker) > 0 Then
    If GetPlayerSP(Attacker) > 0 Then
    StamRemove = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).StamRemove
    Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - StamRemove)
    Call SendSP(Attacker)
End If
End If

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        N = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    Else
        N = 0
    End If

    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Attacker), "ATTACKNPC" & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
     
     If GetPlayerWeaponSlot(Attacker) > 0 Then
             LB = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).LBA
             SB = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).SBA
             BW = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).BWA
             PA = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).PAA
             AA = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AA
             TW = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).TWA
             xb = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).XBA
             BA = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).BA
             If LB > 0 Then
             Call GoLargeBlades(Attacker)
             End If
             If SB > 0 Then
             Call GoSmallBlades(Attacker)
             End If
             If BW > 0 Then
             Call GoBluntWeapons(Attacker)
             End If
             If PA > 0 Then
             Call GoPoles(Attacker)
             End If
             If AA > 0 Then
             Call GoAxes(Attacker)
             End If
             If TW > 0 Then
             Call GoThrown(Attacker)
             End If
             If xb > 0 Then
             Call GoXbows(Attacker)
             End If
             If BA > 0 Then
             Call GoBows(Attacker)
             End If
        End If
     
    MapNum = GetPlayerMap(Attacker)
    npcnum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim$(Npc(npcnum).Name)
    MapNpc(MapNum, MapNpcNum).LastAttack = GetTickCount

    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then

        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "You killed a " & Name, BrightRed, 0)
        'If GetPlayerAlignment(Attacker) < 9989 Then
         '   Call SetPlayerAlignment(Attacker, GetPlayerAlignment(Attacker) + 10)
          '  Call BattleMsg(Attacker, "You Gain 10 Alignment Points !", BrightGreen, 0)
        'End If
        Call SendPlayerData(Attacker)
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

        If GetPlayerHelmetSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerHelmetSlot(Attacker))).AddEXP
        End If
        
        If GetPlayerLegsSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerLegsSlot(Attacker))).AddEXP
        End If
        
        If GetPlayerBootsSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerBootsSlot(Attacker))).AddEXP
        End If
        
        If GetPlayerGlovesSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerGlovesSlot(Attacker))).AddEXP
        End If
        
        If GetPlayerRing1Slot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerRing1Slot(Attacker))).AddEXP
        End If
        
        If GetPlayerRing2Slot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerRing2Slot(Attacker))).AddEXP
        End If
        
        If GetPlayerAmuletSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerAmuletSlot(Attacker))).AddEXP
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
            Exp = Npc(npcnum).Exp + (Npc(npcnum).Exp * Val(Add))
        Else
            Exp = Npc(npcnum).Exp
        End If

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        If Add > 0 Then
            AP = Npc(npcnum).AP + (Npc(npcnum).AP * Val(Add))
        Else
            AP = Npc(npcnum).AP
        End If

        ' Make sure we dont get less then 0
        If AP < 0 Then
            AP = 1
        End If
        
        If GetPlayerAlignment(Attacker) >= 9999 Then
        Call SetPlayerAlignment(Attacker, 9999)
        Call SendPlayerData(Attacker)
        End If

        ' Check if in party, if so divide up the exp
        If Player(Attacker).InParty = NO Then
            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "You cant gain anymore experience!", BrightBlue, 0)
            Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "You have gained " & Exp & " experience.", BrightBlue, 0)
                If GetPlayerAlignment(Attacker) <= 9500 Then
                Call SetPlayerAlignment(Attacker, GetPlayerAlignment(Attacker) + AP)
                Call BattleMsg(Attacker, "You have gained " & AP & " Alignment Points !", BrightCyan, 0)
              End If
            End If
        Else
            o = 0
            For i = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Attacker).PartyID).Member(i) <> Attacker Then
                    If Party(Player(Attacker).PartyID).Member(i) <> 0 Then
                        If GetPlayerMap(Attacker) = GetPlayerMap(Party(Player(Attacker).PartyID).Member(i)) Then
                            o = o + 1
                        End If
                    End If
                End If
            Next

            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "You can't gain anymore experience!", BrightBlue, 0)
            Else

                If o <> 0 Then
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Int(Exp * 0.75))
                    Call BattleMsg(Attacker, "You have gained " & Int(Exp * 0.75) & " experience and shared " & Int(Exp * 0.25) & " with your party.", BrightBlue, 0)
                Else
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                    Call BattleMsg(Attacker, "You have gained " & Exp & " experience but couldn't share any with your party.", BrightBlue, 0)
                End If
            End If

            If o <> 0 Then
                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Attacker).PartyID).Member(i) <> Attacker And Party(Player(Attacker).PartyID).Member(i) <> 0 Then
                        If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                            Call SetPlayerExp(Party(Player(Attacker).PartyID).Member(i), Experience(MAX_LEVEL))
                            Call BattleMsg(Party(Player(Attacker).PartyID).Member(i), "You cant gain anymore experience!", BrightBlue, 0)
                        Else
                            Call SetPlayerExp(Party(Player(Attacker).PartyID).Member(i), GetPlayerExp(Party(Player(Attacker).PartyID).Member(i)) + Int(Exp * (0.25 / o)))
                            Call BattleMsg(Party(Player(Attacker).PartyID).Member(i), "You have gained " & Int(Exp * (0.25 / o)) & " experience from your party.", BrightBlue, 0)
                        End If
                    End If
                Next
            End If
        End If
        For i = 1 To MAX_NPC_DROPS

            ' Drop the goods if they get it
            N = Int(Rnd * Npc(npcnum).ItemNPC(i).Chance) + 1

            If N = 1 Then
                Call SpawnItem(Npc(npcnum).ItemNPC(i).ItemNum, Npc(npcnum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y)
            End If
        Next

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

                If Party(Player(Attacker).PartyID).Member(X) <> 0 Then
                    Call CheckPlayerLevelUp(Party(Player(Attacker).PartyID).Member(X))
                End If
            Next
        End If
        Call TakeFromGrid(MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y)

        ' Check if target is npc that died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_NPC And Player(Attacker).Target = MapNpcNum Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
    Else

        ' NPC not dead, just do the damage
        MapNpc(MapNum, MapNpcNum).HP = MapNpc(MapNum, MapNpcNum).HP - Damage

        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "You hit a " & Name & " for " & Damage & " damage.", White, 0)

        If N = 0 Then

            'Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        Else

            'Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If

        ' Check if we should send a message
        If MapNpc(MapNum, MapNpcNum).Target = 0 And MapNpc(MapNum, MapNpcNum).Target <> Attacker Then
            If Trim$(Npc(npcnum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, "A " & Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay) & "", SayColor)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum, MapNpcNum).Target = Attacker
        MapNpc(MapNum, MapNpcNum).TargetType = TARGET_TYPE_PLAYER

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS

                If MapNpc(MapNum, i).num = MapNpc(MapNum, MapNpcNum).num Then
                    MapNpc(MapNum, i).Target = Attacker
                    MapNpc(MapNum, i).TargetType = TARGET_TYPE_PLAYER
                End If
            Next
        End If
    End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub AttackPlayer(ByVal Attacker As Long, _
   ByVal Victim As Long, _
   ByVal Damage As Long)
Dim Exp As Long
Dim N As Long
Dim OldMap, oldx, oldy As Long
Dim RedoNum As Long
Dim StamRemove As Long
Dim MaxiNum As Long
Dim LSANum As Long
Dim spellnum As Long

If GetPlayerFP(Attacker) < 11 Then
    Call PlayerMsg(Attacker, "Your Hunger Level Is Low, You Need To Eat In Order To Have Strength !", BrightRed)
    Exit Sub
    End If

If GetPlayerWeaponSlot(Attacker) > 0 Then
    If GetPlayerSP(Attacker) > 0 Then
    StamRemove = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).StamRemove
    Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - StamRemove)
    Call SendSP(Attacker)
End If
End If

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        N = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    Else
        N = 0
    End If
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Attacker), "ATTACKPLAYER" & SEP_CHAR & Attacker & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)

    If Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA Then
        If Damage >= GetPlayerHP(Victim) Then

            ' Set HP to nothing
            Call SetPlayerHP(Victim, 0)
            
            If GetPlayerAlignment(Attacker) > 1501 Then
            Call SetPlayerAlignment(Attacker, GetPlayerAlignment(Attacker) - 1500)
            Call BattleMsg(Attacker, "You have Lost 1,500 Alignment Points !", BrightGreen, 0)
            End If

            ' Check for a weapon and say damage
            Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
            Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)

            ' Player is dead
            Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
            Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "Dead" & SEP_CHAR & END_CHAR)

            If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            ' XCORPSEX
                Call CreateCorpse(Victim)
                ' XCORPSEX
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
                    
                    If GetPlayerBootsSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerBootsSlot(Victim), 0)
                    End If
                    
                    If GetPlayerGlovesSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerGlovesSlot(Victim), 0)
                    End If
                    
                    If GetPlayerRing1Slot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerRing1Slot(Victim), 0)
                    End If
                    
                    If GetPlayerRing2Slot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerRing2Slot(Victim), 0)
                    End If
                    
                    If GetPlayerAmuletSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerAmuletSlot(Victim), 0)
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
            OldMap = GetPlayerMap(Victim)
            oldx = GetPlayerX(Victim)
            oldy = GetPlayerY(Victim)

            ' Warp player away
            If SCRIPTING = 1 Then
                Call OnDeath(Victim)
            Else
                Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
            End If
            Call UpdateGrid(OldMap, oldx, oldy, GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim))

            ' Restore vitals
            Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
            Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
            Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
            Call SetPlayerFP(Victim, GetPlayerMaxFP(Victim))
            Call SendHP(Victim)
            Call SendMP(Victim)
            Call SendSP(Victim)
            Call SendFP(Victim)

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
            Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
            Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)

            If N = 0 Then

                'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
            Else

                'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
            End If
        End If
    ElseIf Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then

        If Damage >= GetPlayerHP(Victim) Then

            ' Set HP to nothing
            Call SetPlayerHP(Victim, 0)

            ' Check for a weapon and say damage
            Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
            Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)

            If N = 0 Then

                'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
            Else

                'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
            End If

            ' Player is dead
            Call GlobalMsg(GetPlayerName(Victim) & " has been killed in the arena by " & GetPlayerName(Attacker), BrightRed)
            Call UpdateGrid(GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim), Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data1, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data2, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data3)

            ' Warp player away
            Call PlayerWarp(Victim, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data1, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data2, Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Data3)

            ' Restore vitals
            Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
            Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
            Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
            Call SetPlayerFP(Victim, GetPlayerMaxFP(Victim))
            Call SendHP(Victim)
            Call SendMP(Victim)
            Call SendSP(Victim)
            Call SendFP(Victim)
            
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
            Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
            Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)

            If N = 0 Then

                'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
            Else

                'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
                'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", BrightRed)
            End If
        End If
    End If

    ' Drop the SP
    If GetPlayerSP(Attacker) > 0 Then
    Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - 2)
    Call SendSP(Attacker)
    End If

    ' Reset timer for attacking
    Player(Attacker).AttackTimer = GetTickCount
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "Pain" & SEP_CHAR & END_CHAR)
End Sub

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, npcnum As Long
Dim AttackSpeed As Long
Dim X As Long
Dim y As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 0
    End If
    CanAttackNpc = False
    
    ' Check sp
    If GetPlayerSP(Attacker) > 0 Then
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
    npcnum = MapNpc(MapNum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If npcnum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then

            ' Check if at same coordinates
            X = DirToX(GetPlayerX(Attacker), GetPlayerDir(Attacker))
            y = DirToY(GetPlayerY(Attacker), GetPlayerDir(Attacker))

            If (MapNpc(MapNum, MapNpcNum).y = y) And (MapNpc(MapNum, MapNpcNum).X = X) Then
                If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED And Npc(npcnum).Behavior <> NPC_BEHAVIOR_QUEST And Npc(npcnum).Behavior <> NPC_BEHAVIOR_BANKER Then
                    CanAttackNpc = True
                Else

                    If Npc(npcnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                       MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(npcnum).SpawnSecs
                    ElseIf Npc(npcnum).Behavior = NPC_BEHAVIOR_QUEST Then
                       Call DoQuest(Npc(npcnum).Quest, Attacker, npcnum)
                    ElseIf Npc(npcnum).Behavior = NPC_BEHAVIOR_BANKER Then
                       Call SendDataTo(Attacker, "vaultverify" & SEP_CHAR & END_CHAR)
                    Else
                       Call PlayerMsg(Attacker, Trim(Npc(npcnum).Name) & " :" & Trim(Npc(npcnum).AttackSay), Green)
                    End If

                    If Npc(npcnum).Speech <> 0 Then
                        Call SendDataTo(Attacker, "STARTSPEECH" & SEP_CHAR & Npc(npcnum).Speech & SEP_CHAR & 0 & SEP_CHAR & npcnum & SEP_CHAR & END_CHAR)
                    End If

                End If
            End If
        End If
    End If
End If
End Function

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, npcnum As Long
Dim AttackSpeed As Long
Dim Dir As Long

If GetPlayerFP(Attacker) < 11 Then
    Call PlayerMsg(Attacker, "Your Hunger Level Is Low, You Need To Eat In Order To Have Strength !", BrightRed)
    Exit Function
    End If
    
    If Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ROOF Then
        Call PlayerMsg(Attacker, "You Cannot Use Ranged Weapons Within Buildings !", BrightRed)
        Exit Function
    End If
    
    If GetPlayerArrowsAmount(Attacker) < 1 Then
       Call PlayerMsg(Attacker, "You Are Out of Ammo ! Reload !", BrightRed)
       CanAttackNpcWithArrow = False
       Exit Function
    End If

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 0
    End If
    CanAttackNpcWithArrow = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker), MapNpcNum).num <= 0 Then
        Exit Function
    End If
    MapNum = GetPlayerMap(Attacker)
    npcnum = MapNpc(MapNum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
   ' Make sure they are on the same map
If IsPlaying(Attacker) Then
If npcnum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpcWithArrow = True
            Call SetPlayerAlignment(Attacker, GetPlayerAlignment(Attacker) + 10)
            Call BattleMsg(Attacker, "You Gain 10 Alignment Points !", BrightGreen, 0)
Else

                If Trim$(Npc(npcnum).AttackSay) <> "" Then
                    Call PlayerMsg(Attacker, Trim$(Npc(npcnum).Name) & " : " & Trim$(Npc(npcnum).AttackSay), Green)
                End If
                
                If Npc(npcnum).Speech <> 0 Then
                    For Dir = 0 To 3
                        If DirToX(GetPlayerX(Attacker), Dir) = MapNpc(MapNum, MapNpcNum).X And DirToY(GetPlayerY(Attacker), Dir) = MapNpc(MapNum, MapNpcNum).y Then
                            Call SendDataTo(Attacker, "STARTSPEECH" & SEP_CHAR & Npc(npcnum).Speech & SEP_CHAR & 0 & SEP_CHAR & npcnum & SEP_CHAR & END_CHAR)
                        End If
                    Next Dir
                End If
            End If
        End If
    End If
End Function

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
Dim AttackSpeed As Long
Dim X As Long
Dim y As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 0
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
        X = DirToX(GetPlayerX(Attacker), GetPlayerDir(Attacker))
        y = DirToY(GetPlayerY(Attacker), GetPlayerDir(Attacker))

        If (GetPlayerY(Victim) = y) And (GetPlayerX(Victim) = X) Then
            If Map(GetPlayerMap(Victim)).Tile(X, y).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then

                ' Check to make sure that they dont have access
                If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                    Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
                Else

                    ' Check to make sure the victim isn't an admin
                    If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                    Else

                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then

                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < 10 Then
                                Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                            Else

                                If GetPlayerLevel(Victim) < 10 Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                                Else

                                    If Trim$(GetPlayerGuild(Attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                        If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                            Call SetPlayerAlignment(Attacker, GetPlayerAlignment(Attacker) - 30)
                                            Call BattleMsg(Attacker, "You Lost 30 Alignment Points !", BrightRed, 0)
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
            ElseIf Map(GetPlayerMap(Victim)).Tile(X, y).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                CanAttackPlayer = True
            End If
        End If
    End If
End Function

Function CanAttackPlayerWithArrow(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    
    CanAttackPlayerWithArrow = False
    
    If GetPlayerFP(Attacker) < 11 Then
    Call PlayerMsg(Attacker, "Your Hunger Level Is Low, You Need To Eat In Order To Have Strength !", BrightRed)
    Exit Function
    End If
    
    If Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ROOF Then
        Call PlayerMsg(Attacker, "You Cannot Use Ranged Weapons Within Buildings !", BrightRed)
        Exit Function
    End If
    
    If GetPlayerArrowsAmount(Attacker) < 1 Then
       Call PlayerMsg(Attacker, "You Are Out of Ammo ! Reload !", BrightRed)
       CanAttackPlayerWithArrow = False
       Exit Function
    End If

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
    If GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then
        If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then

            ' Check to make sure that they dont have access
            If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
                Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
            Else

                ' Check to make sure the victim isn't an admin
                If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                    Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                Else

                    ' Check if map is attackable
                    If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then

                        ' Make sure they are high enough level
                        If GetPlayerLevel(Attacker) < 10 Then
                            Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
                        Else

                            If GetPlayerLevel(Victim) < 10 Then
                                Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
                            Else

                                If Trim$(GetPlayerGuild(Attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                                    If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                        CanAttackPlayerWithArrow = True
                                        Call SetPlayerAlignment(Attacker, GetPlayerAlignment(Attacker) - 30)
                                            Call BattleMsg(Attacker, "You Lost 30 Alignment Points !", BrightRed, 0)
                                    Else
                                        Call PlayerMsg(Attacker, "You cant attack a guild member!", BrightRed)
                                    End If
                                Else
                                    CanAttackPlayerWithArrow = True
                                End If
                            End If
                        End If
                    Else
                        Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
                    End If
                End If
            End If
        ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
            CanAttackPlayerWithArrow = True
        End If
    End If
End Function

Function CanNpcAttackPet(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long, npcnum As Long
Dim X As Long
Dim y As Long

    CanNpcAttackPet = False

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Index) = False Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index), MapNpcNum).num <= 0 Then
        Exit Function
    End If
    MapNum = Player(Index).Pet.Map
    npcnum = MapNpc(MapNum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If npcnum > 0 Then
            X = DirToX(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Dir)
            y = DirToY(MapNpc(MapNum, MapNpcNum).y, MapNpc(MapNum, MapNpcNum).Dir)

            ' Check if at same coordinates
            If (Player(Index).Pet.y = y) And (Player(Index).Pet.X = X) Then
                CanNpcAttackPet = True
            End If
        End If
    End If
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long, npcnum As Long
Dim X As Long
Dim y As Long

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
    npcnum = MapNpc(MapNum, MapNpcNum).num

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
        If npcnum > 0 Then
            X = DirToX(MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Dir)
            y = DirToY(MapNpc(MapNum, MapNpcNum).y, MapNpc(MapNum, MapNpcNum).Dir)

            ' Check if at same coordinates
            If (GetPlayerY(Index) = y) And (GetPlayerX(Index) = X) Then
                CanNpcAttackPlayer = True
            End If
        End If
        End If
End Function

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
Dim X As Long, y As Long

    CanNpcMove = False

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Function
    X = DirToX(MapNpc(MapNum, MapNpcNum).X, Dir)
    y = DirToY(MapNpc(MapNum, MapNpcNum).y, Dir)

    If Not IsValid(X, y) Then Exit Function
    If Grid(MapNum).Loc(X, y).Blocked = True Then Exit Function
    If Map(MapNum).Tile(X, y).Type <> TILE_TYPE_WALKABLE And Map(MapNum).Tile(X, y).Type <> TILE_TYPE_ITEM Then Exit Function
    CanNpcMove = True
End Function

Function CanPetAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, npcnum As Long
Dim X As Long
Dim y As Long
Dim Dir As Long

    CanPetAttackNpc = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(Player(Attacker).Pet.Map, MapNpcNum).num <= 0 Then
        Exit Function
    End If
    MapNum = Player(Attacker).Pet.Map
    npcnum = MapNpc(MapNum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If npcnum > 0 And GetTickCount > Player(Attacker).Pet.AttackTimer + 1000 Then
            If Npc(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                For Dir = 0 To 3

                    ' Check if at same coordinates
                    X = DirToX(Player(Attacker).Pet.X, Dir)
                    y = DirToY(Player(Attacker).Pet.y, Dir)

                    If (MapNpc(MapNum, MapNpcNum).y = y) And (MapNpc(MapNum, MapNpcNum).X = X) Then
                        CanPetAttackNpc = True
                    End If
                Next
            End If
        End If
    End If
End Function

Function CanPetMove(ByVal PetNum As Long, ByVal Dir) As Boolean
Dim X As Long, y As Long
Dim i As Long, Packet As String

    CanPetMove = False

    If PetNum <= 0 Or PetNum > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Function
    X = DirToX(Player(PetNum).Pet.X, Dir)
    y = DirToY(Player(PetNum).Pet.y, Dir)

    If Not IsValid(X, y) Then
        If Dir = DIR_UP Then
            If Map(Player(PetNum).Pet.Map).Up > 0 And Map(Player(PetNum).Pet.Map).Up = Player(PetNum).Pet.MapToGo Then
                CanPetMove = True
            End If
        End If

        If Dir = DIR_DOWN Then
            If Map(Player(PetNum).Pet.Map).Down > 0 And Map(Player(PetNum).Pet.Map).Down = Player(PetNum).Pet.MapToGo Then
                CanPetMove = True
            End If
        End If

        If Dir = DIR_LEFT Then
            If Map(Player(PetNum).Pet.Map).Left > 0 And Map(Player(PetNum).Pet.Map).Left = Player(PetNum).Pet.MapToGo Then
                CanPetMove = True
            End If
        End If

        If Dir = DIR_RIGHT Then
            If Map(Player(PetNum).Pet.Map).Right > 0 And Map(Player(PetNum).Pet.Map).Right = Player(PetNum).Pet.MapToGo Then

                'i = Player(PetNum).Pet.Map
                'Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Right
                'Packet = "PETDATA" & SEP_CHAR
                'Packet = Packet & PetNum & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Alive & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Map & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.x & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.y & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Dir & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Sprite & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.HP & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Level * 5 & SEP_CHAR
                'Packet = Packet & END_CHAR
                'Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
                'Call SendDataToMap(i, Packet)
                CanPetMove = True
            End If
        End If
        Exit Function
    End If

    If Grid(Player(PetNum).Pet.Map).Loc(X, y).Blocked = True Then Exit Function
    CanPetMove = True
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
Dim i As Long, N As Long, ShieldSlot As Long, LegsSlot As Long, BootsSlot As Long, GlovesSlot As Long, Ring1Slot As Long, Ring2Slot As Long, AmuletSlot As Long

    CanPlayerBlockHit = False
    ShieldSlot = GetPlayerShieldSlot(Index)

    If ShieldSlot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
    
    CanPlayerBlockHit = False
    LegsSlot = GetPlayerLegsSlot(Index)

    If LegsSlot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
    
    CanPlayerBlockHit = False
    BootsSlot = GetPlayerBootsSlot(Index)

    If LegsSlot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
    
    CanPlayerBlockHit = False
    GlovesSlot = GetPlayerGlovesSlot(Index)

    If GlovesSlot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
    
    CanPlayerBlockHit = False
    Ring1Slot = GetPlayerRing1Slot(Index)

    If Ring1Slot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
    
    CanPlayerBlockHit = False
    Ring2Slot = GetPlayerRing2Slot(Index)

    If Ring2Slot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
    
    CanPlayerBlockHit = False
    AmuletSlot = GetPlayerAmuletSlot(Index)

    If AmuletSlot > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
Dim i As Long, N As Long

    CanPlayerCriticalHit = False

    If GetPlayerWeaponSlot(Index) > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerstr(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            N = Int(Rnd * 100) + 1

            If N <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Sub CastSpell(ByVal Index As Long, _
   ByVal SpellSlot As Long)
Dim spellnum As Long, i As Long, N As Long, Damage As Long
Dim Casted As Boolean
Dim X As Long, y As Long
Dim Packet As String

    Casted = False
    
    Call SendPlayerXY(Index)

    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    spellnum = GetPlayerSpell(Index, SpellSlot)

    ' Make sure player has the spell
    If Not HasSpell(Index, spellnum) Then
        Call BattleMsg(Index, "You do not have this spell!", BrightRed, 0)
        Exit Sub
    End If
    i = GetSpellReqLevel(spellnum)

    ' Check if they have enough MP
    If GetPlayerMP(Index) < Spell(spellnum).MPCost Then
        Call BattleMsg(Index, "Not enough mana!", BrightRed, 0)
        Exit Sub
    End If

    ' Make sure they are the right level
    If i > GetPlayerLevel(Index) Then
        Call BattleMsg(Index, "You must be level " & i & " to cast this spell.", BrightRed, 0)
        Exit Sub
    End If

    ' Check if timer is ok
    If GetTickCount < Player(Index).AttackTimer + 1000 Then
        Exit Sub
    End If
    
    ' Check if the spell is scripted and do that instead of a stat modification
    If Spell(spellnum).Type = SPELL_TYPE_SCRIPTED Then

        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedSpell " & Index & "," & Spell(spellnum).Data1

       Exit Sub
    End If

    ' Check if the spell is a give item and do that instead of a stat modification
    'If Spell(SpellNum).Type = SPELL_TYPE_GIVEITEM Then
    '
    '    N = FindOpenInvSlot(Index, Spell(SpellNum).Data1)
    '    If N > 0 Then
    '
    '        Call GiveItem(Index, Spell(SpellNum).Data1, Spell(SpellNum).Data2)
    '        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & ".", BrightBlue)
    '        ' Take away the mana points
    '        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
    '        Call SendMP(Index)
    '        Casted = True
    '
    '    Else
    '
    '        Call PlayerMsg(Index, "Your inventory is full!", BrightRed)
    '
    '    End If
    '    Exit Sub
    'End If
    ' Check if the spell is a summon and do that instead of a stat modification
    If Spell(spellnum).Type = SPELL_TYPE_PET Then
        Player(Index).Pet.Alive = YES
        Player(Index).Pet.Sprite = Spell(spellnum).Data1
        Player(Index).Pet.Dir = DIR_UP
        Player(Index).Pet.Map = GetPlayerMap(Index)
        Player(Index).Pet.MapToGo = 0
        Player(Index).Pet.X = GetPlayerX(Index) + Int(Rnd * 3 - 1)

        If Player(Index).Pet.X < 0 Or Player(Index).Pet.X > MAX_MAPX Then Player(Index).Pet.X = GetPlayerX(Index)
        Player(Index).Pet.XToGo = -1
        Player(Index).Pet.y = GetPlayerY(Index) + Int(Rnd * 3 - 1)

        If Player(Index).Pet.y < 0 Or Player(Index).Pet.y > MAX_MAPY Then Player(Index).Pet.y = GetPlayerY(Index)
        Player(Index).Pet.YToGo = -1
        Player(Index).Pet.Level = Spell(spellnum).Range
        Player(Index).Pet.HP = Player(Index).Pet.Level * 5
        Call AddToGrid(Player(Index).Pet.Map, Player(Index).Pet.X, Player(Index).Pet.y)
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.X & SEP_CHAR
        Packet = Packet & Player(Index).Pet.y & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR

        ' Excuse the messy code, I'm rushing
        Call PlayerMsg(Index, "You summon a beast", White)
        Call SendDataToMap(GetPlayerMap(Index), Packet)
        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(spellnum).MPCost)
        Call SendMP(Index)
        Casted = True
        Exit Sub
    End If

    If Spell(spellnum).AE = 1 Then
        For y = GetPlayerY(Index) - Spell(spellnum).Range To GetPlayerY(Index) + Spell(spellnum).Range
            For X = GetPlayerX(Index) - Spell(spellnum).Range To GetPlayerX(Index) + Spell(spellnum).Range
                N = -1

                If IsValid(X, y) Then
                    For i = 1 To MAX_PLAYERS

                        If IsPlaying(i) = True Then
                            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                                If GetPlayerX(i) = X And GetPlayerY(i) = y Then
                                    If i = Index Then
                                        If Spell(spellnum).Type = SPELL_TYPE_ADDHP Or Spell(spellnum).Type = SPELL_TYPE_ADDMP Or Spell(spellnum).Type = SPELL_TYPE_ADDSP Then
                                            Player(Index).Target = i
                                            Player(Index).TargetType = TARGET_TYPE_PLAYER
                                            N = Player(Index).Target
                                        End If
                                    Else
                                        Player(Index).Target = i
                                        Player(Index).TargetType = TARGET_TYPE_PLAYER
                                        N = Player(Index).Target
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS

                        If MapNpc(GetPlayerMap(Index), i).num > 0 Then
                            If MapNpc(GetPlayerMap(Index), i).X = X And MapNpc(GetPlayerMap(Index), i).y = y Then
                                Player(Index).Target = i
                                Player(Index).TargetType = TARGET_TYPE_NPC
                                N = Player(Index).Target
                            End If
                        End If
                    Next

                    If N < 0 Then
                        Player(Index).Target = MakeLoc(X, y)
                        Player(Index).TargetType = TARGET_TYPE_LOCATION
                        N = MakeLoc(X, y)
                    End If
                    Casted = False

                    If N > 0 Then
                        If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
                            If IsPlaying(N) Then
                                If N <> Index Then
                                    Player(Index).TargetType = TARGET_TYPE_PLAYER

                                    If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then

                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Select Case Spell(spellnum).Type

                                            Case SPELL_TYPE_SUBHP
                                                Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(spellnum).Data1) - GetPlayerProtection(N) + (Rnd * 5) - 2

                                                If Damage > 0 Then
                                                    Call AttackPlayer(Index, N, Damage)
                                                    Call SetPlayerAlignment(N, GetPlayerAlignment(N) - 30)
                                                    Call BattleMsg(N, "You have Lost 30 Alignment Points !", BrightGreen, 0)
                                                    Call SendPlayerData(N)
                                                Else
                                                    Call BattleMsg(Index, "The spell was to weak to hurt " & GetPlayerName(N) & "!", BrightRed, 0)
                                                End If

                                            Case SPELL_TYPE_SUBMP
                                                Call SetPlayerMP(N, GetPlayerMP(N) - Spell(spellnum).Data1)
                                                Call SendMP(N)

                                            Case SPELL_TYPE_SUBSP
                                                Call SetPlayerSP(N, GetPlayerSP(N) - Spell(spellnum).Data1)
                                                Call SendSP(N)
                                        End Select
                                        Casted = True
                                    Else

                                        If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(spellnum).Type >= SPELL_TYPE_ADDHP And Spell(spellnum).Type <= SPELL_TYPE_ADDSP Then

                                            Select Case Spell(spellnum).Type

                                                Case SPELL_TYPE_ADDHP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerHP(N, GetPlayerHP(N) + Spell(spellnum).Data1)
                                                    Call SendHP(N)
                                                     If GetPlayerAlignment(N) < 9994 Then
                                                    Call SetPlayerAlignment(N, GetPlayerAlignment(N) + 5)
                                                    Call BattleMsg(N, "You Gaint 5 Alignment Points !", BrightGreen, 0)
                                                    Call SendPlayerData(N)
                                                    End If

                                                Case SPELL_TYPE_ADDMP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerMP(N) + Spell(spellnum).Data1)
                                                    Call SendMP(N)
                                                    If GetPlayerAlignment(N) < 9994 Then
                                                    Call SetPlayerAlignment(N, GetPlayerAlignment(N) + 5)
                                                    Call BattleMsg(N, "You Gaint 5 Alignment Points !", BrightGreen, 0)
                                                    Call SendPlayerData(N)
                                                    End If

                                                Case SPELL_TYPE_ADDSP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerSP(N) + Spell(spellnum).Data1)
                                                    Call SendMP(N)
                                                    If GetPlayerAlignment(N) < 9994 Then
                                                     Call SetPlayerAlignment(N, GetPlayerAlignment(N) + 5)
                                                    Call BattleMsg(N, "You Gaint 5 Alignment Points !", BrightGreen, 0)
                                                    Call SendPlayerData(N)
                                                    End If
                                            End Select
                                            Casted = True
                                        Else
                                            Call PlayerMsg(Index, "Could not cast spell!", BrightRed)
                                        End If
                                    End If
                                Else
                                    Player(Index).TargetType = TARGET_TYPE_PLAYER

                                    If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then
                                    Else

                                        If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(spellnum).Type >= SPELL_TYPE_ADDHP And Spell(spellnum).Type <= SPELL_TYPE_ADDSP Then

                                            Select Case Spell(spellnum).Type

                                                Case SPELL_TYPE_ADDHP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerHP(N, GetPlayerHP(N) + Spell(spellnum).Data1)
                                                    Call SendHP(N)
                                                    If GetPlayerAlignment(N) < 9994 Then
                                                     Call SetPlayerAlignment(N, GetPlayerAlignment(N) + 5)
                                                    Call BattleMsg(N, "You Gaint 5 Alignment Points !", BrightGreen, 0)
                                                    Call SendPlayerData(N)
                                                    End If

                                                Case SPELL_TYPE_ADDMP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerMP(N) + Spell(spellnum).Data1)
                                                    Call SendMP(N)
                                                    If GetPlayerAlignment(N) < 9994 Then
                                                     Call SetPlayerAlignment(N, GetPlayerAlignment(N) + 5)
                                                    Call BattleMsg(N, "You Gaint 5 Alignment Points !", BrightGreen, 0)
                                                    Call SendPlayerData(N)
                                                    End If

                                                Case SPELL_TYPE_ADDSP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerSP(N) + Spell(spellnum).Data1)
                                                    Call SendMP(N)
                                                    If GetPlayerAlignment(N) < 9994 Then
                                                     Call SetPlayerAlignment(N, GetPlayerAlignment(N) + 5)
                                                    Call BattleMsg(N, "You Gaint 5 Alignment Points !", BrightGreen, 0)
                                                    Call SendPlayerData(N)
                                                    End If
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

                            If Player(Index).TargetType = TARGET_TYPE_NPC Then
                                If Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_QUEST Then
                                    If Spell(spellnum).Type >= SPELL_TYPE_SUBHP And Spell(spellnum).Type <= SPELL_TYPE_SUBSP Then

                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                                        Select Case Spell(spellnum).Type

                                            Case SPELL_TYPE_SUBHP
                                                Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(spellnum).Data1) - Int(Npc(MapNpc(GetPlayerMap(Index), N).num).DEF / 2) + (Rnd * 5) - 2

                                                If Damage > 0 Then
                                                    If Spell(spellnum).Element <> 0 And Npc(MapNpc(GetPlayerMap(Index), N).num).Element <> 0 Then
                                                If Element(Spell(spellnum).Element).Strong = Npc(MapNpc(GetPlayerMap(Index), N).num).Element Or Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Weak = Spell(spellnum).Element Then
                                                   Call BattleMsg(Index, "     A Deadly Mix of Elements Harm The " & Trim(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & "!", BrightGreen, 0)
                                                   Damage = Int(Damage * 1.25)
                                                If Element(Spell(spellnum).Element).Strong = Npc(MapNpc(GetPlayerMap(Index), N).num).Element And Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Weak = Spell(spellnum).Element Then Damage = Int(Damage * 1.2)
                                                End If
                                
                                                If Element(Spell(spellnum).Element).Weak = Npc(MapNpc(GetPlayerMap(Index), N).num).Element Or Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Strong = Spell(spellnum).Element Then
                                                   Call BattleMsg(Index, " The " & Trim(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & " aborbs much of the elemental damage!", BrightRed, 0)
                                                   Damage = Int(Damage * 0.75)
                                                If Element(Spell(spellnum).Element).Weak = Npc(MapNpc(GetPlayerMap(Index), N).num).Element And Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Strong = Spell(spellnum).Element Then Damage = Int(Damage * (2 / 3))
                                                End If
                                                End If
                                                    Call AttackNpc(Index, N, Damage)
                                                    If GetPlayerAlignment(N) < 9994 Then
                                                     Call SetPlayerAlignment(N, GetPlayerAlignment(N) + 5)
                                                    Call BattleMsg(N, "You Gain 5 Alignment Points !", BrightGreen, 0)
                                                    Call SendPlayerData(N)
                                                    End If
                                                Else
                                                    Call BattleMsg(Index, "The spell was to weak to hurt " & Trim$(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & "!", BrightRed, 0)
                                                End If

                                            Case SPELL_TYPE_SUBMP
                                                MapNpc(GetPlayerMap(Index), N).MP = MapNpc(GetPlayerMap(Index), N).MP - Spell(spellnum).Data1

                                            Case SPELL_TYPE_SUBSP
                                                MapNpc(GetPlayerMap(Index), N).SP = MapNpc(GetPlayerMap(Index), N).SP - Spell(spellnum).Data1
                                        End Select
                                        Casted = True
                                    Else

                                        Select Case Spell(spellnum).Type

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
                                    Call BattleMsg(Index, "Could not cast spell!", BrightRed, 0)
                                End If
                            Else
                                Player(Index).TargetType = TARGET_TYPE_LOCATION
                                Casted = True
                            End If
                        End If
                    End If

                    If Casted = True Then
                        Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & spellnum & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & END_CHAR)

                        'Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
                    End If
                End If
            Next
        Next
        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(spellnum).MPCost)
        Call SendMP(Index)
    Else
        N = Player(Index).Target

        If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
            If IsPlaying(N) Then
                If GetPlayerName(N) <> GetPlayerName(Index) Then
                    If CInt(Sqr((GetPlayerX(Index) - GetPlayerX(N)) ^ 2 + ((GetPlayerY(Index) - GetPlayerY(N)) ^ 2))) > Spell(spellnum).Range Then
                        Call BattleMsg(Index, "You are too far away to hit the target.", BrightRed, 0)
                        Exit Sub
                    End If
                End If
                Player(Index).TargetType = TARGET_TYPE_PLAYER

                If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then

                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                    Select Case Spell(spellnum).Type

                        Case SPELL_TYPE_SUBHP
                            Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(spellnum).Data1) - GetPlayerProtection(N) + (Rnd * 5) - 2

                            If Damage > 0 Then
                                Call AttackPlayer(Index, N, Damage)
                                Call SetPlayerAlignment(N, GetPlayerAlignment(N) - 30)
                                Call BattleMsg(N, "You have Lost 30 Alignment Points !", BrightGreen, 0)
                                Call SendPlayerData(N)
                            Else
                                Call BattleMsg(Index, "The spell was to weak to hurt " & GetPlayerName(N) & "!", BrightRed, 0)
                            End If

                        Case SPELL_TYPE_SUBMP
                            Call SetPlayerMP(N, GetPlayerMP(N) - Spell(spellnum).Data1)
                            Call SendMP(N)
                            Call SetPlayerAlignment(N, GetPlayerAlignment(N) - 30)
                                Call BattleMsg(N, "You have Lost 30 Alignment Points !", BrightGreen, 0)
                                Call SendPlayerData(N)

                        Case SPELL_TYPE_SUBSP
                            Call SetPlayerSP(N, GetPlayerSP(N) - Spell(spellnum).Data1)
                            Call SendSP(N)
                            Call SetPlayerAlignment(N, GetPlayerAlignment(N) - 30)
                                Call BattleMsg(N, "You have Lost 30 Alignment Points !", BrightGreen, 0)
                                Call SendPlayerData(N)
                    End Select

                    ' Take away the mana points
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(spellnum).MPCost)
                    Call SendMP(Index)
                    Casted = True
                Else

                    If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(spellnum).Type >= SPELL_TYPE_ADDHP And Spell(spellnum).Type <= SPELL_TYPE_ADDSP Then

                        Select Case Spell(spellnum).Type

                            Case SPELL_TYPE_ADDHP

                                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerHP(N, GetPlayerHP(N) + Spell(spellnum).Data1)
                                Call SendHP(N)

                            Case SPELL_TYPE_ADDMP

                                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(N, GetPlayerMP(N) + Spell(spellnum).Data1)
                                Call SendMP(N)

                            Case SPELL_TYPE_ADDSP

                                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(N, GetPlayerSP(N) + Spell(spellnum).Data1)
                                Call SendMP(N)
                        End Select

                        ' Take away the mana points
                        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(spellnum).MPCost)
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

            If CInt(Sqr((GetPlayerX(Index) - MapNpc(GetPlayerMap(Index), N).X) ^ 2 + ((GetPlayerY(Index) - MapNpc(GetPlayerMap(Index), N).y) ^ 2))) > Spell(spellnum).Range Then
                Call BattleMsg(Index, "You are too far away to hit the target.", BrightRed, 0)
                Exit Sub
            End If
            Player(Index).TargetType = TARGET_TYPE_NPC

            If Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then

                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                Select Case Spell(spellnum).Type

                    Case SPELL_TYPE_ADDHP
                        MapNpc(GetPlayerMap(Index), N).HP = MapNpc(GetPlayerMap(Index), N).HP + Spell(spellnum).Data1

                    Case SPELL_TYPE_SUBHP
                        Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(spellnum).Data1) - Int(Npc(MapNpc(GetPlayerMap(Index), N).num).DEF / 2 + (Rnd * 5) - 2)

                        If Damage > 0 Then
                        If Spell(spellnum).Element <> 0 And Npc(MapNpc(GetPlayerMap(Index), N).num).Element <> 0 Then
                                If Element(Spell(spellnum).Element).Strong = Npc(MapNpc(GetPlayerMap(Index), N).num).Element Or Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Weak = Spell(spellnum).Element Then
                                    Call BattleMsg(Index, "     A Deadly Mix of Elements Harm The " & Trim(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & "!", BrightGreen, 0)
                                    Damage = Int(Damage * 1.25)
                                If Element(Spell(spellnum).Element).Strong = Npc(MapNpc(GetPlayerMap(Index), N).num).Element And Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Weak = Spell(spellnum).Element Then Damage = Int(Damage * 1.2)
                                End If
                                
                                If Element(Spell(spellnum).Element).Weak = Npc(MapNpc(GetPlayerMap(Index), N).num).Element Or Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Strong = Spell(spellnum).Element Then
                                    Call BattleMsg(Index, " The " & Trim(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & " aborbs much of the elemental damage!", BrightRed, 0)
                                    Damage = Int(Damage * 0.75)
                                If Element(Spell(spellnum).Element).Weak = Npc(MapNpc(GetPlayerMap(Index), N).num).Element And Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Strong = Spell(spellnum).Element Then Damage = Int(Damage * (2 / 3))
                                End If
                                End If
                            Call AttackNpc(Index, N, Damage)
                            If GetPlayerAlignment(N) < 9994 Then
                            Call SetPlayerAlignment(N, GetPlayerAlignment(N) + 5)
                                Call BattleMsg(N, "You Gain 5 Alignment Points !", BrightGreen, 0)
                                Call SendPlayerData(N)
                                End If
                        Else
                            Call BattleMsg(Index, "The spell was to weak to hurt " & Trim$(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & "!", BrightRed, 0)
                        End If

                    Case SPELL_TYPE_ADDMP
                        MapNpc(GetPlayerMap(Index), N).MP = MapNpc(GetPlayerMap(Index), N).MP + Spell(spellnum).Data1

                    Case SPELL_TYPE_SUBMP
                        MapNpc(GetPlayerMap(Index), N).MP = MapNpc(GetPlayerMap(Index), N).MP - Spell(spellnum).Data1

                    Case SPELL_TYPE_ADDSP
                        MapNpc(GetPlayerMap(Index), N).SP = MapNpc(GetPlayerMap(Index), N).SP + Spell(spellnum).Data1

                    Case SPELL_TYPE_SUBSP
                        MapNpc(GetPlayerMap(Index), N).SP = MapNpc(GetPlayerMap(Index), N).SP - Spell(spellnum).Data1
                End Select

                ' Take away the mana points
                Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(spellnum).MPCost)
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
        Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & spellnum & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & Player(Index).CastedSpell & SEP_CHAR & END_CHAR)

        If Spell(spellnum).sound > 0 Then Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Magic" & Spell(spellnum).sound & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub CheckEquippedItems(ByVal Index As Long)
Dim Slot As Long, ItemNum As Long
Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    Slot = GetPlayerWeaponSlot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)
        
        ' Make sure they are the right class
                            
                                

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
    Slot = GetPlayerLegsSlot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_LEGS Then
                Call SetPlayerLegsSlot(Index, 0)
            End If
        Else
            Call SetPlayerLegsSlot(Index, 0)
        End If
    End If
    Slot = GetPlayerBootsSlot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_BOOTS Then
                Call SetPlayerBootsSlot(Index, 0)
            End If
        Else
            Call SetPlayerBootsSlot(Index, 0)
        End If
    End If
    Slot = GetPlayerGlovesSlot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_GLOVES Then
                Call SetPlayerGlovesSlot(Index, 0)
            End If
        Else
            Call SetPlayerGlovesSlot(Index, 0)
        End If
    End If
    Slot = GetPlayerRing1Slot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_RING1 Then
                Call SetPlayerRing1Slot(Index, 0)
            End If
        Else
            Call SetPlayerRing1Slot(Index, 0)
        End If
    End If
    Slot = GetPlayerRing2Slot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_RING2 Then
                Call SetPlayerRing2Slot(Index, 0)
            End If
        Else
            Call SetPlayerRing2Slot(Index, 0)
        End If
    End If
    Slot = GetPlayerAmuletSlot(Index)

    If Slot > 0 Then
        ItemNum = GetPlayerInvItemNum(Index, Slot)

        If ItemNum > 0 Then
            If Item(ItemNum).Type <> ITEM_TYPE_AMULET Then
                Call SetPlayerAmuletSlot(Index, 0)
            End If
        Else
            Call SetPlayerAmuletSlot(Index, 0)
        End If
    End If
End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
Dim i As Long
Dim d As Long
Dim c As Long
Dim xT As Long

xT = 2
    c = 0

    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        If GetPlayerLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerExp(Index) < GetPlayerNextLevel(Index)
                    DoEvents

                    If GetPlayerLevel(Index) < MAX_LEVEL Then
                        If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
                            d = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
                            Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerExp(Index, d)
                            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + xT)
                            Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                            Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
                            Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " Levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a Level!", 14)
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
    Call SendFP(Index)
    Call SendStats(Index)
End Sub

' Another thing I want to be widely used. Instead of the giant select statements,
' just throw in a few of these and everything works fine
Public Function DirToX(ByVal X As Long, _
   ByVal Dir As Byte) As Long
    DirToX = X

    If Dir = DIR_UP Or Dir = DIR_DOWN Then Exit Function

    ' LEFT = 2, RIGHT = 3
    ' 2 * 2 = 4, 4 - 5 = -1
    ' 3 * 2 = 6, 6 - 5 = 1
    DirToX = X + ((Dir * 2) - 5)
End Function

Public Function DirToY(ByVal y As Long, _
   ByVal Dir As Byte) As Long
    DirToY = y

    If Dir = DIR_LEFT Or Dir = DIR_RIGHT Then Exit Function

    ' UP = 0, DOWN = 1
    ' 0 * 2 = 0, 0 - 1 = -1
    ' 1 * 2 = 2, 2 - 1 = 1
    DirToY = y + ((Dir * 2) - 1)
End Function

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long

    FindOpenInvSlot = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then

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

Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
   
    FindOpenBankSlot = 0
   
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
   
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i
    End If
   
    For i = 1 To MAX_BANK
        ' Try to find an open free slot
        If GetPlayerBankItemNum(Index, i) = 0 Then
            FindOpenBankSlot = i
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
    Next
End Function

Function FindOpenPlayerSlot() As Long
Dim i As Long

    FindOpenPlayerSlot = 0
    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next
End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
Dim i As Long

    FindOpenSpellSlot = 0
    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next
End Function

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next
    FindPlayer = 0
End Function

Function GetNpcHPRegen(ByVal npcnum As Long)
Dim i As Long

    'Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcHPRegen = 0
        Exit Function
    End If
    i = Int(Npc(npcnum).DEF / 3)

    If i < 1 Then i = 1
    GetNpcHPRegen = i
End Function

Function GetNpcMaxHP(ByVal npcnum As Long)

    ' Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcMaxHP = 0
        Exit Function
    End If
    GetNpcMaxHP = Npc(npcnum).MaxHp
End Function

Function GetNpcMaxMP(ByVal npcnum As Long)

    ' Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcMaxMP = 0
        Exit Function
    End If
    GetNpcMaxMP = Npc(npcnum).Magi * 2
End Function

Function GetNpcMaxSP(ByVal npcnum As Long)

    ' Prevent subscript out of range
    If npcnum <= 0 Or npcnum > MAX_NPCS Then
        GetNpcMaxSP = 0
        Exit Function
    End If
    GetNpcMaxSP = Npc(npcnum).Speed * 2
End Function

Function GetPlayerDamage(ByVal Index As Long) As Long
Dim WeaponSlot As Long

    GetPlayerDamage = (Rnd * 5) - 2

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    
    GetPlayerDamage = Int(GetPlayerstr(Index) / 2)

    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If

    If GetPlayerWeaponSlot(Index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(Index)
        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2

        If GetPlayerInvItemDur(Index, WeaponSlot) > 0 Then
            Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) - 1)

            If GetPlayerInvItemDur(Index, WeaponSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, WeaponSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, WeaponSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data1), Yellow, 0)
                End If
            End If
        Else
            If GetPlayerInvItemDur(Index, WeaponSlot) < 0 Then
                Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) + 1)
    
                If GetPlayerInvItemDur(Index, WeaponSlot) = 0 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " has broken.", Yellow, 0)
                    Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
                Else
                    If GetPlayerInvItemDur(Index, WeaponSlot) >= -10 Then
                        Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, WeaponSlot) * -1 & "/" & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data1) * -1, Yellow, 0)
                    End If
                End If
            End If
        End If
    End If

    If GetPlayerDamage < 0 Then
        GetPlayerDamage = 0
    End If
End Function

Function GetPlayerHPRegen(ByVal Index As Long)
Dim i As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen") = 1 Then

        ' Prevent subscript out of range
        If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
            GetPlayerHPRegen = 0
            Exit Function
        End If
        i = Int(GetPlayerDEF(Index) / 2)

        If i < 2 Then i = 2
        GetPlayerHPRegen = i
    End If
End Function

Function GetPlayerMPRegen(ByVal Index As Long)
Dim i As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen") = 1 Then

        ' Prevent subscript out of range
        If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
            GetPlayerMPRegen = 0
            Exit Function
        End If
        i = Int(GetPlayerMAGI(Index) / 2)

        If i < 2 Then i = 2
        GetPlayerMPRegen = i
    End If
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long

    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    ArmorSlot = GetPlayerArmorSlot(Index)
    HelmSlot = GetPlayerHelmetSlot(Index)
    ShieldSlot = GetPlayerShieldSlot(Index)
    GetPlayerProtection = Int(GetPlayerDEF(Index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data2

        If GetPlayerInvItemDur(Index, ArmorSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, ArmorSlot, GetPlayerInvItemDur(Index, ArmorSlot) - 1)

            If GetPlayerInvItemDur(Index, ArmorSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ArmorSlot), 0)
            Else

                If GetPlayerInvItemDur(Index, ArmorSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, ArmorSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If

    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, HelmSlot)).Data2

        If GetPlayerInvItemDur(Index, HelmSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, HelmSlot, GetPlayerInvItemDur(Index, HelmSlot) - 1)

            If GetPlayerInvItemDur(Index, HelmSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, HelmSlot), 0)
            Else

                If GetPlayerInvItemDur(Index, HelmSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, HelmSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If

    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data2

        If GetPlayerInvItemDur(Index, ShieldSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, ShieldSlot, GetPlayerInvItemDur(Index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(Index, ShieldSlot) = 0 Then
                Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " has broken.", Yellow, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ShieldSlot), 0)
            Else

                If GetPlayerInvItemDur(Index, ShieldSlot) <= 10 Then
                    Call BattleMsg(Index, "Your " & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(Index, ShieldSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data1), Yellow, 0)
                End If
            End If
        End If
    End If
End Function

Function GetPlayerSPRegen(ByVal Index As Long)
Dim i As Long

    If GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen") = 1 Then

        ' Prevent subscript out of range
        If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
            GetPlayerSPRegen = 0
            Exit Function
        End If
        i = Int(GetPlayerSPEED(Index) / 2)

        If i < 2 Then i = 2
        GetPlayerSPRegen = i
    End If
End Function

Function GetSpellReqLevel(ByVal spellnum As Long)
    GetSpellReqLevel = Spell(spellnum).LevelReq ' - Int(GetClassMAGI(GetPlayerClass(index)) / 4)
End Function

Function GetItemReqLevel(ByVal ItemNum As Long)
    GetItemReqLevel = Item(ItemNum).LevelReq ' - Int(GetClassMAGI(GetPlayerClass(index)) / 4)
End Function

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
Dim i As Long, N As Long

    N = 0
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            N = N + 1
        End If
    Next
    GetTotalMapPlayers = N
End Function

Sub GiveItem(ByVal Index As Long, _
   ByVal ItemNum As Long, _
   ByVal ItemVal As Long)
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

        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_BOOTS) Or (Item(ItemNum).Type = ITEM_TYPE_GLOVES) Or (Item(ItemNum).Type = ITEM_TYPE_RING1) Or (Item(ItemNum).Type = ITEM_TYPE_RING2) Or (Item(ItemNum).Type = ITEM_TYPE_AMULET) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
        End If
        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
    End If
End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long, N As Long
Dim TakeBankItem As Boolean

    TakeBankItem = False
   
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
   
    For i = 1 To MAX_BANK
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                ' Is what we are trying to take away more then what they have? If so just set it to zero
                If ItemVal >= GetPlayerBankItemValue(Index, i) Then
                    TakeBankItem = True
                Else
                    Call SetPlayerBankItemValue(Index, i, GetPlayerBankItemValue(Index, i) - ItemVal)
                    Call SendBankUpdate(Index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerBankItemNum(Index, i)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If i = GetPlayerWeaponSlot(Index) Then
                                Call SetPlayerWeaponSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                Call SendInvSlots(Index)
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
                            If i = GetPlayerArmorSlot(Index) Then
                                Call SetPlayerArmorSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                Call SendInvSlots(Index)
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
                            If i = GetPlayerHelmetSlot(Index) Then
                                Call SetPlayerHelmetSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                Call SendInvSlots(Index)
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
                            If i = GetPlayerShieldSlot(Index) Then
                                Call SetPlayerShieldSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                Call SendInvSlots(Index)
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
                            If i = GetPlayerLegsSlot(Index) Then
                                Call SetPlayerLegsSlot(Index, 0)
                                Call SendInvSlots(Index)
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
                        
                            Case ITEM_TYPE_BOOTS
                        If GetPlayerBootsSlot(Index) > 0 Then
                            If i = GetPlayerBootsSlot(Index) Then
                                Call SetPlayerBootsSlot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerBootsSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                        
                        Case ITEM_TYPE_GLOVES
                        If GetPlayerGlovesSlot(Index) > 0 Then
                            If i = GetPlayerGlovesSlot(Index) Then
                                Call SetPlayerGlovesSlot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerGlovesSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                        
                        Case ITEM_TYPE_RING1
                        If GetPlayerRing1Slot(Index) > 0 Then
                            If i = GetPlayerRing1Slot(Index) Then
                                Call SetPlayerRing1Slot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerRing1Slot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                        
                        Case ITEM_TYPE_RING2
                        If GetPlayerRing2Slot(Index) > 0 Then
                            If i = GetPlayerRing2Slot(Index) Then
                                Call SetPlayerRing2Slot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerRing2Slot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                        
                        Case ITEM_TYPE_AMULET
                        If GetPlayerAmuletSlot(Index) > 0 Then
                            If i = GetPlayerAmuletSlot(Index) Then
                                Call SetPlayerAmuletSlot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerAmuletSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                End Select

               
                N = Item(GetPlayerBankItemNum(Index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (N <> ITEM_TYPE_WEAPON) And (N <> ITEM_TYPE_ARMOR) And (N <> ITEM_TYPE_HELMET) And (N <> ITEM_TYPE_SHIELD) And (N <> ITEM_TYPE_LEGS) And (N <> ITEM_TYPE_BOOTS) And (N <> ITEM_TYPE_GLOVES) And (N <> ITEM_TYPE_RING1) And (N <> ITEM_TYPE_RING2) And (N <> ITEM_TYPE_AMULET) Then
                    TakeBankItem = True
                End If
            End If
                           
            If TakeBankItem = True Then
                Call SetPlayerBankItemNum(Index, i, 0)
                Call SetPlayerBankItemValue(Index, i, 0)
                Call SetPlayerBankItemDur(Index, i, 0)
               
                ' Send the Bank update
                Call SendBankUpdate(Index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal BankSlot As Long)
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
   
    i = BankSlot
   
    ' Check to see if Bankentory is full
    If i <> 0 Then
        Call SetPlayerBankItemNum(Index, i, ItemNum)
        Call SetPlayerBankItemValue(Index, i, GetPlayerBankItemValue(Index, i) + ItemVal)
       
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_BOOTS) Or (Item(ItemNum).Type = ITEM_TYPE_GLOVES) Or (Item(ItemNum).Type = ITEM_TYPE_RING1) Or (Item(ItemNum).Type = ITEM_TYPE_RING2) Or (Item(ItemNum).Type = ITEM_TYPE_AMULET) Then
            Call SetPlayerBankItemDur(Index, i, Item(ItemNum).Data1)
        End If
    Else
        Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "Bank full!" & SEP_CHAR & END_CHAR)
    End If
End Sub

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long

    HasItem = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next
End Function

Function HasSpell(ByVal Index As Long, ByVal spellnum As Long) As Boolean
Dim i As Long

    HasSpell = False
    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = spellnum Then
            HasSpell = True
            Exit Function
        End If
    Next
End Function

Public Function IsValid(ByVal X As Long, _
   ByVal y As Long) As Boolean
    IsValid = True

    If X < 0 Or X > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then IsValid = False
End Function

Sub JoinGame(ByVal Index As Long)
Dim MOTD As String

    ' Set the flag so we know the person is in the game
    Player(Index).InGame = True

    ' Send an ok to client to start receiving in game data
    Call SendDataTo(Index, "LOGINOK" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
    Call SendDataTo(Index, "sound" & SEP_CHAR & "LoggingIntoServer" & SEP_CHAR & END_CHAR)

    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendEmoticons(Index)
    Call SendElements(Index)
    Call SendSpeech(Index)
    Call SendQuest(Index)
    Call SendArrows(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendInventory(Index)
    Call SendBank(Index)
    Call SendInvSlots(Index)
    Call SendWornEquipment(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendFP(Index)
    Call SendStats(Index)
    Call SendDataTo(Index, "Sethands" & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).Hands & SEP_CHAR & END_CHAR)
    Call SendWeatherTo(Index)
    Call SendTimeTo(Index)
    Call SendGameClockTo(Index)
    Call SendNewsTo(Index)
    Call SendOnlineList
    Call SendFriendListTo(Index)
    Call SendFriendListToNeeded(GetPlayerName(Index))
    Call SendAllCorpseTo(Index)
    

    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), False)
    Call SendPlayerData(Index)

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "JoinGame " & Index
    Else

        If Not ExistVar("motd.ini", "MOTD", "Msg") Then Call MsgBox("OMG OMG!")
        MOTD = GetVar("motd.ini", "MOTD", "Msg")

        ' Send a global message that he/she joined
        If GetPlayerAccess(Index) <= ADMIN_MONITER Then
            Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", 7)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has joined " & GAME_NAME & "!", 15)
        End If
        Call SendDataToAllBut(Index, "sound" & SEP_CHAR & "ANewPlayerHasJoined" & SEP_CHAR & END_CHAR)

        ' Send them welcome
        Call PlayerMsg(Index, "Welcome to " & GAME_NAME & "!", 15)

        ' Send motd
        If Trim$(MOTD) <> "" Then
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
Dim N As Long
Dim i As Long

    If Player(Index).InGame = True Then
        Player(Index).InGame = False
        Call SendDataTo(Index, "sound" & SEP_CHAR & "LoggingOutOfServer" & SEP_CHAR & END_CHAR)
        Call SendDataToAllBut(Index, "sound" & SEP_CHAR & "APlayerHasLeft" & SEP_CHAR & END_CHAR)

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) = 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If

        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If Player(Index).InParty = YES Then
            N = 0
            For i = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Index).PartyID).Member(i) = Index Then N = i
            Next
            For i = N To MAX_PARTY_MEMBERS - 1
                Party(Player(Index).PartyID).Member(i) = Party(Player(Index).PartyID).Member(i + 1)
            Next
            Party(Player(Index).PartyID).Member(MAX_PARTY_MEMBERS) = 0
            N = 0
            For i = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Index).PartyID).Member(i) <> 0 And Party(Player(Index).PartyID).Member(i) <> Index Then
                    N = N + 1
                    Call PlayerMsg(Party(Player(Index).PartyID).Member(i), GetPlayerName(Index) & " has left the party.", Pink)
                End If
            Next

            If N < 2 Then
                If Party(Player(Index).PartyID).Member(1) <> 0 Then
                    Call PlayerMsg(Party(Player(Index).PartyID).Member(1), "Party disbanded.", Pink)
                    Player(Party(Player(Index).PartyID).Member(1)).InParty = NO
                    Player(Party(Player(Index).PartyID).Member(1)).PartyID = 0
                    Party(Player(Index).PartyID).Member(1) = 0
                End If
            End If
            Player(Index).PartyID = 0
            Player(Index).InParty = NO
        End If

        If SCRIPTING = 1 Then
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
        For N = 1 To MAX_PLAYERS
            Call ShowPLR(N)
        Next
    End If
    Call SendFriendListToNeeded(GetPlayerName(Index))
    Call ClearPlayer(Index)
    Call SendOnlineList
End Sub

' I want to start using the loc system. Instead of two variables...
' (x and y), you can store both as a "loc" and extract them back
Public Function MakeLoc(ByVal X As Long, _
   ByVal y As Long) As Long
    MakeLoc = (y * MAX_MAPX) + X
End Function

Public Function MakeX(ByVal Loc As Long) As Long
    MakeX = Loc - (MakeY(Loc) * MAX_MAPX)
End Function

Public Function MakeY(ByVal Loc As Long) As Long
    MakeY = Int(Loc / MAX_MAPX)
End Function

Sub NpcAttackPet(ByVal MapNpcNum As Long, _
   ByVal Victim As Long, _
   ByVal Damage As Long)
Dim Name As String
Dim MapNum As Long
Dim Packet As String

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(Player(Victim).Pet.Map, MapNpcNum).num <= 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the npc attacking
    Call SendDataToMap(Player(Victim).Pet.Map, "NPCATTACKPET" & SEP_CHAR & MapNpcNum & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    MapNum = Player(Victim).Pet.Map
    Name = Trim$(Npc(MapNpc(MapNum, MapNpcNum).num).Name)

    If Damage >= Player(Victim).Pet.HP Then
        Call BattleMsg(Victim, "Your pet died!", Red, 1)
        Player(Victim).Pet.Alive = NO
        Call TakeFromGrid(Player(Victim).Pet.Map, Player(Victim).Pet.X, Player(Victim).Pet.y)
        MapNpc(MapNum, MapNpcNum).Target = 0
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Victim & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.X & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.y & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataTo(Victim, Packet)
        Call SendDataToMapBut(Victim, Player(Victim).Pet.Map, Packet)
    Else

        ' Pet not dead, just do the damage
        Player(Victim).Pet.HP = Player(Victim).Pet.HP - Damage
        Packet = "PETHP" & SEP_CHAR & Player(Victim).Pet.Level * 5 & SEP_CHAR & Player(Victim).Pet.HP & SEP_CHAR & END_CHAR
        Call SendDataTo(Victim, Packet)
    End If

    'Call SendDataTo(Victim, "BLITNPCDMGPET" & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
End Sub

Sub PoisonActive(ByVal Index As Long)
Dim Damage As Long
Dim NpCPoisonDamage As Long

NpCPoisonDamage = 50

          Call SetPlayerHP(Index, GetPlayerHP(Index) - NpCPoisonDamage)
          Call PlayerMsg(Index, "You have Lost " & NpCPoisonDamage & " HP Due To Poison !", BrightRed)
          Call SendStats(Index)
          If Damage >= GetPlayerHP(Index) Then

            ' Set HP to nothing
            Call SetPlayerHP(Index, 0)
            If Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_NO_PENALTY Then
                Call CreateCorpse(Index)
            End If
            If SCRIPTING = 1 Then
            Call OnDeath(Index)
            Call SendDataToMap(GetPlayerMap(Index), "poisonover" & SEP_CHAR & END_CHAR)
            Call SetPlayerHP(Index, (GetPlayerMaxHP(Index)))
            End If
            
            ' Player is dead
            Call GlobalMsg(GetPlayerName(Index) & " has been killed by " & " Poison.", BrightRed)
            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Dead" & SEP_CHAR & END_CHAR)
            
        End If
          Call SendStats(Index)
End Sub

Sub DiseaseActive(ByVal Index As Long)
Dim Damage As Long
Dim NpCDiseaseDamage As Long

NpCDiseaseDamage = 30

          Call SetPlayerHP(Index, GetPlayerHP(Index) - NpCDiseaseDamage)
          Call PlayerMsg(Index, "You have Lost " & NpCDiseaseDamage & " HP Due To Disease !", BrightRed)
          Call SendStats(Index)
          If Damage >= GetPlayerHP(Index) Then

            ' Set HP to nothing
            Call SetPlayerHP(Index, 0)
            If Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_NO_PENALTY Then
                Call CreateCorpse(Index)
            End If
            If SCRIPTING = 1 Then
            Call OnDeath(Index)
            Call SendDataToMap(GetPlayerMap(Index), "diseaseover" & SEP_CHAR & END_CHAR)
            Call SetPlayerHP(Index, (GetPlayerMaxHP(Index)))
            End If
            
            ' Player is dead
            Call GlobalMsg(GetPlayerName(Index) & " has been killed by " & " Disease.", BrightRed)
            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Dead" & SEP_CHAR & END_CHAR)
            
        End If
          Call SendStats(Index)
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, _
   ByVal Victim As Long, _
   ByVal Damage As Long)
Dim Name As String
Dim Exp As Long
Dim MapNum As Long
Dim OldMap, oldx, oldy As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim), MapNpcNum).num <= 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    
    If Npc(MapNpcNum).Poison = 1 Then
    Call SendDataToMap(GetPlayerMap(Victim), "poisonbegin" & SEP_CHAR & END_CHAR)
    End If
    
    If Npc(MapNpcNum).Disease = 1 Then
    Call SendDataToMap(GetPlayerMap(Victim), "diseasebegin" & SEP_CHAR & END_CHAR)
    End If
    
    
    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(MapNpc(MapNum, MapNpcNum).num).Name)

    If Damage >= GetPlayerHP(Victim) Then

        ' Say damage
        Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)

        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by a " & Name, BrightRed)
        Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "Dead" & SEP_CHAR & END_CHAR)

        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            ' XCORPSEX
                Call CreateCorpse(Victim)
                ' XCORPSEX
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
            End If
            
            ' Calculate exp to take from the player
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
        OldMap = GetPlayerMap(Victim)
        oldx = GetPlayerX(Victim)
        oldy = GetPlayerY(Victim)

        ' Warp player away
        If SCRIPTING = 1 Then
            Call OnDeath(Victim)
        Else
            Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        End If
        Call UpdateGrid(OldMap, oldx, oldy, GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim))

        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SetPlayerFP(Victim, GetPlayerMaxFP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)
        Call SendFP(Victim)

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
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "Pain" & SEP_CHAR & END_CHAR)
End Sub

Sub NpcDir(ByVal MapNum As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Dir As Long)
Dim Packet As String

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    Packet = "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub NpcMove(ByVal MapNum As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Dir As Long, _
   ByVal Movement As Long)
Dim Packet As String
Dim X As Long
Dim y As Long

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    X = DirToX(MapNpc(MapNum, MapNpcNum).X, Dir)
    y = DirToY(MapNpc(MapNum, MapNpcNum).y, Dir)
    Call UpdateGrid(MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y, MapNum, X, y)
    MapNpc(MapNum, MapNpcNum).y = y
    MapNpc(MapNum, MapNpcNum).X = X
    Packet = "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub PetAttackNpc(ByVal Attacker As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Damage As Long)
Dim Name As String
Dim N As Long, i As Long
Dim MapNum As Long, npcnum As Long
Dim Dir As Long, X As Long, y As Long
Dim Packet As String

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the pet attacking
    Call SendDataToMap(Player(Attacker).Pet.Map, "PETATTACKNPC" & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    MapNum = Player(Attacker).Pet.Map
    npcnum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim$(Npc(npcnum).Name)
    MapNpc(MapNum, MapNpcNum).LastAttack = GetTickCount
    For Dir = 0 To 3

        If MapNpc(MapNum, npcnum).X = DirToX(Player(Attacker).Pet.X, Dir) And MapNpc(MapNum, npcnum).y = DirToY(Player(Attacker).Pet.y, Dir) Then
            Packet = "CHANGEPETDIR" & SEP_CHAR & Dir & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR
            Call SendDataToMap(Player(Attacker).Pet.Map, Packet)
        End If
    Next

    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        For i = 1 To MAX_NPC_DROPS

            ' Drop the goods if they get it
            N = Int(Rnd * Npc(npcnum).ItemNPC(i).Chance) + 1

            If N = 1 Then
                Call SpawnItem(Npc(npcnum).ItemNPC(i).ItemNum, Npc(npcnum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y)
            End If
        Next
        Call BattleMsg(Attacker, "Your pet killed a " & Name & ".", Red, 1)

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
        Call TakeFromGrid(MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y)

        ' Check if target is npc that died and if so set target to 0
        If Player(Attacker).Pet.TargetType = TARGET_TYPE_NPC And Player(Attacker).Pet.Target = MapNpcNum Then
            Player(Attacker).Pet.Target = 0
            Player(Attacker).Pet.TargetType = 0
            Player(Attacker).Pet.MapToGo = 0
        End If
    Else

        ' NPC not dead, just do the damage
        MapNpc(MapNum, MapNpcNum).HP = MapNpc(MapNum, MapNpcNum).HP - Damage

        ' Set the NPC target to the pet
        MapNpc(MapNum, MapNpcNum).TargetType = TARGET_TYPE_PET
        MapNpc(MapNum, MapNpcNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS

                If MapNpc(MapNum, i).num = MapNpc(MapNum, MapNpcNum).num Then
                    MapNpc(MapNum, i).TargetType = TARGET_TYPE_PET
                    MapNpc(MapNum, i).Target = Attacker
                End If
            Next
        End If
    End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)
    ' Reset attack timer
    Player(Attacker).Pet.AttackTimer = GetTickCount
End Sub

Sub PetMove(ByVal PetNum As Long, _
   ByVal Dir As Long, _
   ByVal Movement As Long)
Dim Packet As String
Dim X As Long
Dim y As Long
Dim i As Long

    If GetPlayerMap(PetNum) <= 0 Or GetPlayerMap(PetNum) > MAX_MAPS Or PetNum <= 0 Or PetNum > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub
    Player(PetNum).Pet.Dir = Dir
    X = DirToX(Player(PetNum).Pet.X, Dir)
    y = DirToY(Player(PetNum).Pet.y, Dir)

    If IsValid(X, y) Then
        If Grid(Player(PetNum).Pet.Map).Loc(X, y).Blocked = True Then
            Packet = "CHANGEPETDIR" & SEP_CHAR & Dir & SEP_CHAR & PetNum & SEP_CHAR & END_CHAR
            Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
            Exit Sub
        End If
        Call UpdateGrid(Player(PetNum).Pet.Map, Player(PetNum).Pet.X, Player(PetNum).Pet.y, Player(PetNum).Pet.Map, X, y)
        Player(PetNum).Pet.y = y
        Player(PetNum).Pet.X = X
        Packet = "PETMOVE" & SEP_CHAR & PetNum & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
        Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
    Else
        i = Player(PetNum).Pet.Map

        If Dir = DIR_UP Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Up
            Player(PetNum).Pet.y = MAX_MAPY
        End If

        If Dir = DIR_DOWN Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Down
            Player(PetNum).Pet.y = 0
        End If

        If Dir = DIR_LEFT Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Left
            Player(PetNum).Pet.X = MAX_MAPX
        End If

        If Dir = DIR_RIGHT Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Right
            Player(PetNum).Pet.X = 0
        End If
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & PetNum & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Map & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.X & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.y & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.HP & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
        Call SendDataToMap(i, Packet)
    End If
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, _
   ByVal InvNum As Long, _
   ByVal Amount As Long)
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
                        Call SendInvSlots(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)

                Case ITEM_TYPE_WEAPON

                    If InvNum = GetPlayerWeaponSlot(Index) Then
                        Call SetPlayerWeaponSlot(Index, 0)
                        Call SendWornEquipment(Index)
                        Call SendInvSlots(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)

                Case ITEM_TYPE_HELMET

                    If InvNum = GetPlayerHelmetSlot(Index) Then
                        Call SetPlayerHelmetSlot(Index, 0)
                        Call SendWornEquipment(Index)
                        Call SendInvSlots(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)

                Case ITEM_TYPE_SHIELD

                    If InvNum = GetPlayerShieldSlot(Index) Then
                        Call SetPlayerShieldSlot(Index, 0)
                        Call SendWornEquipment(Index)
                        Call SendInvSlots(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            
                Case ITEM_TYPE_LEGS

                    If InvNum = GetPlayerLegsSlot(Index) Then
                        Call SetPlayerLegsSlot(Index, 0)
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
                    
                Case ITEM_TYPE_BOOTS

                    If InvNum = GetPlayerBootsSlot(Index) Then
                        Call SetPlayerBootsSlot(Index, 0)
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            
                Case ITEM_TYPE_GLOVES

                    If InvNum = GetPlayerGlovesSlot(Index) Then
                        Call SetPlayerGlovesSlot(Index, 0)
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            
                Case ITEM_TYPE_RING1

                    If InvNum = GetPlayerRing1Slot(Index) Then
                        Call SetPlayerRing1Slot(Index, 0)
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            
                 Case ITEM_TYPE_RING2

                    If InvNum = GetPlayerRing2Slot(Index) Then
                        Call SetPlayerRing2Slot(Index, 0)
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            
                 Case ITEM_TYPE_AMULET

                    If InvNum = GetPlayerAmuletSlot(Index) Then
                        Call SetPlayerAmuletSlot(Index, 0)
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                    End If
                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            End Select
            MapItem(GetPlayerMap(Index), i).num = GetPlayerInvItemNum(Index, InvNum)
            MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
            MapItem(GetPlayerMap(Index), i).y = GetPlayerY(Index)

            If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then

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

                If Item(GetPlayerInvItemNum(Index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_SHIELD And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_LEGS And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_BOOTS And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_GLOVES And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_RING1 And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_RING2 And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_AMULET Then
                    If Item(GetPlayerInvItemNum(Index, InvNum)).Data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - Ind.", Yellow)
                    Else

                        If Item(GetPlayerInvItemNum(Index, InvNum)).Data1 > 0 Then
                            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - " & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Data1 & ".", Yellow)
                        Else
                            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - " & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Data1 * -1 & ".", Yellow)
                        End If
                    End If
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
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).num, Amount, MapItem(GetPlayerMap(Index), i).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
        Else
            Call PlayerMsg(Index, "To many items already on the ground.", BrightRed)
        End If
    End If
End Sub

Sub PlayerMapGetItem(ByVal Index As Long)
Dim i As Long
Dim N As Long
Dim MapNum As Long
Dim Msg As String

    If IsPlaying(Index) = False Then
        Exit Sub
    End If
    MapNum = GetPlayerMap(Index)
    For i = 1 To MAX_MAP_ITEMS

        ' See if theres even an item here
        If (MapItem(MapNum, i).num > 0) And (MapItem(MapNum, i).num <= MAX_ITEMS) Then

            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).X = GetPlayerX(Index)) And (MapItem(MapNum, i).y = GetPlayerY(Index)) Then

                ' Find open slot
                N = FindOpenInvSlot(Index, MapItem(MapNum, i).num)

                ' Open slot available?
                If N <> 0 Then

                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(Index, N, MapItem(MapNum, i).num)

                    If Item(GetPlayerInvItemNum(Index, N)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, N)).Stackable = 1 Then
                        Call SetPlayerInvItemValue(Index, N, GetPlayerInvItemValue(Index, N) + MapItem(MapNum, i).Value)
                        Msg = "You picked up " & MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, N)).Name) & "."
                    Else
                        Call SetPlayerInvItemValue(Index, N, 0)
                        Msg = "You picked up a " & Trim$(Item(GetPlayerInvItemNum(Index, N)).Name) & "."
                    End If
                    Call SetPlayerInvItemDur(Index, N, MapItem(MapNum, i).Dur)

                    ' Erase item from the map
                    MapItem(MapNum, i).num = 0
                    MapItem(MapNum, i).Value = 0
                    MapItem(MapNum, i).Dur = 0
                    MapItem(MapNum, i).X = 0
                    MapItem(MapNum, i).y = 0
                    Call SendInventoryUpdate(Index, N)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                    Call PlayerMsg(Index, Msg, Yellow)
                    Exit Sub
                Else
                    Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

Sub PlayerMove(ByVal Index As Long, _
   ByVal Dir As Long, _
   ByVal Movement As Long)
Dim Packet As String
Dim MapNum As Long
Dim X As Long
Dim y As Long
Dim oldx As Long
Dim oldy As Long
Dim OldMap As Long
Dim Moved As Byte

If Moved = YES Then
'reduce SP by 1 when moving
' Drop the SP
If GetPlayerSP(Index) > 0 Then
Call SetPlayerSP(Index, GetPlayerSP(Index) - 1)
Call SendSP(Index)
End If
End If

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
    X = DirToX(GetPlayerX(Index), Dir)
    y = DirToY(GetPlayerY(Index), Dir)
    Call TakeFromGrid(GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

    ' Move the player's pet out of the way if we need to
    If Player(Index).Pet.Alive = YES Then
        If Player(Index).Pet.Map = GetPlayerMap(Index) And Player(Index).Pet.X = X And Player(Index).Pet.y = y Then
            If Grid(GetPlayerMap(Index)).Loc(DirToX(X, Dir), DirToY(y, Dir)).Blocked = False Then
                Call UpdateGrid(Player(Index).Pet.Map, Player(Index).Pet.X, Player(Index).Pet.y, Player(Index).Pet.Map, DirToX(X, Dir), DirToY(y, Dir))
                Player(Index).Pet.y = DirToY(y, Dir)
                Player(Index).Pet.X = DirToX(X, Dir)
                Packet = "PETMOVE" & SEP_CHAR & Index & SEP_CHAR & DirToX(X, Dir) & SEP_CHAR & DirToY(y, Dir) & SEP_CHAR & Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                Call SendDataToMap(Player(Index).Pet.Map, Packet)
            End If
        End If
    End If

    ' Check to make sure not outside of boundries
    If IsValid(X, y) Then
        ' Check to make sure that the tile is walkable
        If Grid(GetPlayerMap(Index)).Loc(X, y).Blocked = False Then
            ' Check to see if the tile is a key and if it is check if its opened
            If (Map(GetPlayerMap(Index)).Tile(X, y).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(X, y).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(X, y).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(X, y).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = YES) Then
                Call SetPlayerX(Index, X)
                Call SetPlayerY(Index, y)
                Packet = "PLAYERMOVE" & SEP_CHAR & Index & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                Call SendDataToMapBut(Index, GetPlayerMap(Index), Packet)
                Moved = YES
            End If
        End If
    Else
        ' Check to see if we can move them to the another map
        If Map(GetPlayerMap(Index)).Up > 0 And Dir = DIR_UP Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), MAX_MAPY)
            Moved = YES
        End If

        If Map(GetPlayerMap(Index)).Down > 0 And Dir = DIR_DOWN Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
            Moved = YES
        End If

        If Map(GetPlayerMap(Index)).Left > 0 And Dir = DIR_LEFT Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, MAX_MAPX, GetPlayerY(Index))
            Moved = YES
        End If

        If Map(GetPlayerMap(Index)).Right > 0 And Dir = DIR_RIGHT Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
            Moved = YES
        End If
    End If
    
    If Moved = NO Then Call SendPlayerXY(Index)

    If GetPlayerX(Index) < 0 Or GetPlayerY(Index) < 0 Or GetPlayerX(Index) > MAX_MAPX Or GetPlayerY(Index) > MAX_MAPY Or GetPlayerMap(Index) <= 0 Then
        Call HackingAttempt(Index, "")
        Exit Sub
    End If

    'healing tiles code
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SendHP(Index)
        Call PlayerMsg(Index, "You feel a sudden rush through your body as you regain strength!", BrightGreen)
    End If

    'Check for kill tile, and if so kill them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KILL Then
        Call SetPlayerHP(Index, 0)
        Call PlayerMsg(Index, "You embrace the cold finger of death; and feel your life extinguished", BrightRed)

        ' Warp player away
        If SCRIPTING = 1 Then
            Call OnDeath(Index)
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

    If IsValid(X, y) Then
        If Map(GetPlayerMap(Index)).Tile(X, y).Type = TILE_TYPE_DOOR Then
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Key" & SEP_CHAR & END_CHAR)
            End If
        End If
    End If

    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3
        Call PlayerWarp(Index, MapNum, X, y)
        Moved = YES
    End If
    Call AddToGrid(GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

    ' Check for key trigger open
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KEYOPEN Then
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2

        If Map(GetPlayerMap(Index)).Tile(X, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = NO Then
            TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)

            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = "" Then
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", White)
            Else
                Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
            End If
            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Key" & SEP_CHAR & END_CHAR)
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
                    Call PlayerMsg(Index, "This sprite will cost you " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 & " " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", Yellow)
                Else
                    Call PlayerMsg(Index, "This sprite will cost you a " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", Yellow)
                End If
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    
    ' Check if player stepped on house buying tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HOUSE_BUY Then
        If GetPlayerName(Index) = Map(GetPlayerMap(Index)).Owner Then
            'Call PlayerMsg(Index, "You already own this house!", BrightRed)
            Call SendDataTo(Index, "housesell" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If Map(GetPlayerMap(Index)).Owner = "" Then
        If GetPlayerName(Index) = Map(GetPlayerMap(Index)).Owner Then
            Call QuestMsg(Index, "You already own this house!", BrightRed)
            Exit Sub
        Else
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 = 0 Then
                Call SendDataTo(Index, "housebuy" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Else
                If Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Type = ITEM_TYPE_CURRENCY Then
                    Call QuestMsg(Index, "This house will cost you " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 & " " & Trim(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Name) & "!", Yellow)
                Else
                    Call QuestMsg(Index, "This house will cost you a " & Trim(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Name) & "!", Yellow)
                End If
                Call SendDataTo(Index, "housebuy" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            End If
        End If
            Else
    Call QuestMsg(Index, "This house is not for sale!", BrightRed)
    Exit Sub
    End If
    End If
    
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SPAWNGATE Then
        Call SetPlayerSpawnGateMap(Index, GetPlayerMap(Index))
        Call SetPlayerSpawnGateX(Index, GetPlayerX(Index))
        Call SetPlayerSpawnGateY(Index, GetPlayerY(Index))
        Call QuestMsg(Index, "Your Spawn Gate has Been Marked !", BrightGreen)
 End If

    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 > 0 Then
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
            Call SetPlayerstr(Index, (Player(Index).Char(Player(Index).CharNum).STR - Class(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF - Class(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi - Class(GetPlayerClass(Index)).Magi))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed - Class(GetPlayerClass(Index)).Speed))
            Call SetPlayerClass(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
            Call SetPlayerstr(Index, (Player(Index).Char(Player(Index).CharNum).STR + Class(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF + Class(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi + Class(GetPlayerClass(Index)).Magi))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed + Class(GetPlayerClass(Index)).Speed))
            Call PlayerMsg(Index, "Your new class is a " & Trim$(Class(GetPlayerClass(Index)).Name) & "!", BrightGreen)
            Call SendStats(Index)
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & END_CHAR)
        End If
    End If

    ' Check if player stepped on notice tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_NOTICE Then
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> "" Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), Black)
        End If

        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2) <> "" Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2), Grey)
        End If
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 & SEP_CHAR & END_CHAR)
    End If

    ' Check if player stepped on sound tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1 & SEP_CHAR & END_CHAR)
    End If

    ' Check if player stepped on Bank tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_BANK Then
        Call SendDataTo(Index, "openbank" & SEP_CHAR & END_CHAR)
    End If

    If SCRIPTING = 1 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        End If
    End If
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal y As Long, Optional sound As Boolean = True)
Dim OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    'If Trim$(Shop(ShopNum).LeaveSay) <> "" Then
    'Call PlayerMsg(Index, Trim$(Shop(ShopNum).Name) & " : " & Trim$(Shop(ShopNum).LeaveSay) & "", SayColor)
    'End If
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    Call SendLeaveMap(Index, OldMap)
    Call UpdateGrid(OldMap, GetPlayerX(Index), GetPlayerY(Index), MapNum, X, y)
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, y)

    If Player(Index).Pet.Alive = YES Then
        Player(Index).Pet.MapToGo = MapNum
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    Player(Index).GettingMap = YES
    If sound Then Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Warp" & SEP_CHAR & END_CHAR)
    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendInvSlots(Index)
End Sub

Public Sub RemovePLR()
    frmServer.lvUsers.ListItems.Clear
End Sub

Sub SetUpGrid()
Dim i As Long
Dim X As Long
Dim y As Long

    Call ClearGrid
    For i = 1 To MAX_MAPS
        For X = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY

                If Map(i).Tile(X, y).Type = TILE_TYPE_BLOCKED Then Grid(i).Loc(X, y).Blocked = True
            Next
        Next
        For X = 1 To MAX_MAP_NPCS
            If MapNpc(i, X).num > 0 Then
                Grid(i).Loc(MapNpc(i, X).X, MapNpc(i, X).y).Blocked = True
            End If
        Next
    Next
End Sub

Public Sub ShowPLR(ByVal Index As Long)
Dim ls As ListItem

    On Error Resume Next

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(Index) = True Then
        frmServer.lvUsers.ListItems.Remove Index
    End If
    Set ls = frmServer.lvUsers.ListItems.Add(Index, , Index)

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

Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next
End Sub

Sub SpawnAllMapsItems()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next
End Sub

Sub SpawnItem(ByVal ItemNum As Long, _
   ByVal ItemVal As Long, _
   ByVal MapNum As Long, _
   ByVal X As Long, _
   ByVal y As Long)
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, X, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, _
   ByVal ItemNum As Long, _
   ByVal ItemVal As Long, _
   ByVal ItemDur As Long, _
   ByVal MapNum As Long, _
   ByVal X As Long, _
   ByVal y As Long)
Dim Packet As String
Dim i As Long

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    i = MapItemSlot

    If i <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
        MapItem(MapNum, i).num = ItemNum
        MapItem(MapNum, i).Value = ItemVal

        If ItemNum <> 0 Then
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type <= ITEM_TYPE_LEGS) Or (Item(ItemNum).Type <= ITEM_TYPE_BOOTS) Or (Item(ItemNum).Type <= ITEM_TYPE_GLOVES) Or (Item(ItemNum).Type <= ITEM_TYPE_RING1) Or (Item(ItemNum).Type <= ITEM_TYPE_RING2) Or (Item(ItemNum).Type <= ITEM_TYPE_AMULET) Then
                MapItem(MapNum, i).Dur = ItemDur
            Else
                MapItem(MapNum, i).Dur = 0
            End If
        Else
            MapItem(MapNum, i).Dur = 0
        End If
        MapItem(MapNum, i).X = X
        MapItem(MapNum, i).y = y
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
Dim X As Long
Dim y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(X, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If (Item(Map(MapNum).Tile(X, y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(MapNum).Tile(X, y).Data1).Stackable = 1) And Map(MapNum).Tile(X, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(X, y).Data1, 1, MapNum, X, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(X, y).Data1, Map(MapNum).Tile(X, y).Data2, MapNum, X, y)
                End If
            End If
        Next
    Next
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Packet As String
Dim npcnum As Long
Dim i As Long, X As Long, y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    Spawned = False
    npcnum = Map(MapNum).Npc(MapNpcNum)

    If npcnum > 0 Then
        If GameTime = TIME_NIGHT Then
            If Npc(npcnum).SpawnTime = 1 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        Else

            If Npc(npcnum).SpawnTime = 2 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
        MapNpc(MapNum, MapNpcNum).num = npcnum
        MapNpc(MapNum, MapNpcNum).Target = 0
        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxHP(npcnum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(npcnum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(npcnum)
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

        If Map(MapNum).NpcSpawn(MapNpcNum).Used <> 1 Then

            ' Well try  times to randomly place the sprite
            For i = 1 To 100
                X = Int(Rnd * MAX_MAPX)
                y = Int(Rnd * MAX_MAPY)

                ' Check if the tile is walkable
                If Map(MapNum).Tile(X, y).Type = TILE_TYPE_WALKABLE Then
                    MapNpc(MapNum, MapNpcNum).X = X
                    MapNpc(MapNum, MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
            Next

            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then
                For y = 0 To MAX_MAPY
                    For X = 0 To MAX_MAPX

                        If Map(MapNum).Tile(X, y).Type = TILE_TYPE_WALKABLE Then
                            MapNpc(MapNum, MapNpcNum).X = X
                            MapNpc(MapNum, MapNpcNum).y = y
                            Spawned = True
                            Exit For
                        End If
                    Next
                Next
            End If
        Else
            MapNpc(MapNum, MapNpcNum).X = Map(MapNum).NpcSpawn(MapNpcNum).X
            MapNpc(MapNum, MapNpcNum).y = Map(MapNum).NpcSpawn(MapNpcNum).y
            Spawned = True
        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Call AddToGrid(MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y)
        End If
    End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)
End Sub

Sub TakeFromGrid(ByVal OldMap, _
   ByVal oldx, _
   ByVal oldy)
    Grid(OldMap).Loc(oldx, oldy).Blocked = False

    If Map(OldMap).Tile(oldx, oldy).Type = TILE_TYPE_BLOCKED Then Grid(OldMap).Loc(oldx, oldy).Blocked = True
End Sub

Sub TakeItem(ByVal Index As Long, _
   ByVal ItemNum As Long, _
   ByVal ItemVal As Long)
Dim i As Long, N As Long
Dim TakeItem As Boolean

    TakeItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then

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
                                Call SendInvSlots(Index)
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
                                Call SendInvSlots(Index)
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
                                Call SendInvSlots(Index)
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
                                Call SendInvSlots(Index)
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
                            If i = GetPlayerLegsSlot(Index) Then
                                Call SetPlayerLegsSlot(Index, 0)
                                Call SendInvSlots(Index)
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
                        
                        Case ITEM_TYPE_BOOTS

                        If GetPlayerLegsSlot(Index) > 0 Then
                            If i = GetPlayerBootsSlot(Index) Then
                                Call SetPlayerBootsSlot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else

                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerBootsSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                        Case ITEM_TYPE_GLOVES

                        If GetPlayerGlovesSlot(Index) > 0 Then
                            If i = GetPlayerGlovesSlot(Index) Then
                                Call SetPlayerGlovesSlot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else

                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerGlovesSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                        Case ITEM_TYPE_RING1

                        If GetPlayerRing1Slot(Index) > 0 Then
                            If i = GetPlayerRing1Slot(Index) Then
                                Call SetPlayerRing1Slot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else

                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerRing1Slot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                        Case ITEM_TYPE_RING2

                        If GetPlayerRing2Slot(Index) > 0 Then
                            If i = GetPlayerRing2Slot(Index) Then
                                Call SetPlayerRing2Slot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else

                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerRing2Slot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                        
                        Case ITEM_TYPE_AMULET

                        If GetPlayerAmuletSlot(Index) > 0 Then
                            If i = GetPlayerAmuletSlot(Index) Then
                                Call SetPlayerAmuletSlot(Index, 0)
                                Call SendInvSlots(Index)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else

                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select
                N = Item(GetPlayerInvItemNum(Index, i)).Type

                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (N <> ITEM_TYPE_WEAPON) And (N <> ITEM_TYPE_ARMOR) And (N <> ITEM_TYPE_HELMET) And (N <> ITEM_TYPE_SHIELD) And (N <> ITEM_TYPE_LEGS) And (N <> ITEM_TYPE_BOOTS) And (N <> ITEM_TYPE_GLOVES) And (N <> ITEM_TYPE_RING1) And (N <> ITEM_TYPE_RING2) And (N <> ITEM_TYPE_AMULET) Then
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
    Next
End Sub

Function TotalOnlinePlayers() As Long
Dim i As Long

    TotalOnlinePlayers = 0
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next
End Function

Sub UpdateGrid(ByVal OldMap, _
   ByVal oldx, _
   ByVal oldy, _
   ByVal NewMap, _
   ByVal NewX, _
   ByVal NewY)
    Grid(OldMap).Loc(oldx, oldy).Blocked = False
    Grid(NewMap).Loc(NewX, NewY).Blocked = True

    If Map(OldMap).Tile(oldx, oldy).Type = TILE_TYPE_BLOCKED Then Grid(OldMap).Loc(oldx, oldy).Blocked = True
End Sub

Sub ResetMapGrid(ByVal MapNum As Long)
Dim X As Long
Dim y As Long
Dim i As Long

        For X = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                Grid(MapNum).Loc(X, y).Blocked = False
                If Map(MapNum).Tile(X, y).Type = TILE_TYPE_BLOCKED Then Grid(MapNum).Loc(X, y).Blocked = True
            Next
        Next
       
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(MapNum, i).num > 0 Then
                Grid(MapNum).Loc(MapNpc(MapNum, i).X, MapNpc(MapNum, i).y).Blocked = True
            End If
        Next
       
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    Grid(MapNum).Loc(GetPlayerX(i), GetPlayerY(i)).Blocked = True
                End If
            End If
        Next
End Sub

Sub ScriptSetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
Dim Packet As String
With Map(mapper).Tile(X, y)
    .Type = Attrib
    .Data1 = Data1
    .Data2 = Data2
    .Data3 = Data3
    .String1 = String1
    .String2 = String2
    .String3 = String3
End With

End Sub

Function FindOpenCorpseLoot(ByVal Index As Integer) As Byte
Dim i As Byte

FindOpenCorpseLoot = 0

For i = 1 To 4
If Player(Index).CorpseLoot(i).num = 0 Then
FindOpenCorpseLoot = i
Exit Function
End If
Next i
End Function
Sub ClearCorpse(ByVal Index As Integer)
Dim i As Byte

Player(Index).CorpseMap = 0
Player(Index).CorpseX = 0
Player(Index).CorpseY = 0


For i = 1 To 4
Player(Index).CorpseLoot(i).num = 0
Player(Index).CorpseLoot(i).Dur = 0
Player(Index).CorpseLoot(i).Value = 0
Next i
End Sub
Sub CreateCorpse(ByVal Index As Integer)
Dim N As Byte, b As Byte, i As Byte

If Player(Index).CorpseMap > 0 Then
For i = 1 To 4
If Player(Index).CorpseLoot(i).num > 0 Then
Call SpawnItem(Player(Index).CorpseLoot(i).num, 0, Player(Index).CorpseMap, Player(Index).CorpseX, Player(Index).CorpseY)
End If
Next i
End If


Player(Index).CorpseMap = GetPlayerMap(Index)
Player(Index).CorpseX = GetPlayerX(Index)
Player(Index).CorpseY = GetPlayerY(Index)


For i = 1 To 4
Player(Index).CorpseLoot(i).num = 0
Player(Index).CorpseLoot(i).Dur = 0
Player(Index).CorpseLoot(i).Value = 0
Next i

If GetPlayerWeaponSlot(Index) > 0 Then
N = GetPlayerWeaponSlot(Index)
b = FindOpenCorpseLoot(Index)
If b > 0 Then
Player(Index).CorpseLoot(b).num = GetPlayerInvItemNum(Index, N)
Player(Index).CorpseLoot(b).Dur = GetPlayerInvItemDur(Index, N)
Player(Index).CorpseLoot(b).Value = 0
Call TakeItem(Index, GetPlayerInvItemNum(Index, N), 1)
End If
End If


If GetPlayerArmorSlot(Index) > 0 Then
N = GetPlayerArmorSlot(Index)
b = FindOpenCorpseLoot(Index)
If b > 0 Then
Player(Index).CorpseLoot(b).num = GetPlayerInvItemNum(Index, N)
Player(Index).CorpseLoot(b).Dur = GetPlayerInvItemDur(Index, N)
Player(Index).CorpseLoot(b).Value = 0
Call TakeItem(Index, GetPlayerInvItemNum(Index, N), 1)
End If
End If


If GetPlayerHelmetSlot(Index) > 0 Then
N = GetPlayerHelmetSlot(Index)
b = FindOpenCorpseLoot(Index)
If b > 0 Then
Player(Index).CorpseLoot(b).num = GetPlayerInvItemNum(Index, N)
Player(Index).CorpseLoot(b).Dur = GetPlayerInvItemDur(Index, N)
Player(Index).CorpseLoot(b).Value = 0
Call TakeItem(Index, GetPlayerInvItemNum(Index, N), 1)
End If
End If

If GetPlayerShieldSlot(Index) > 0 Then
N = GetPlayerShieldSlot(Index)
b = FindOpenCorpseLoot(Index)
If b > 0 Then
Player(Index).CorpseLoot(b).num = GetPlayerInvItemNum(Index, N)
Player(Index).CorpseLoot(b).Dur = GetPlayerInvItemDur(Index, N)
Player(Index).CorpseLoot(b).Value = 0
Call TakeItem(Index, GetPlayerInvItemNum(Index, N), 1)
End If
End If

Player(Index).CorpseTimer = GetTickCount
Call PlayerMsg(Index, "You have Died !", BrightRed)
Call SendCorpseToAll(Index)
End Sub

Sub SendCorpseToAll(ByVal Index As Integer)
Dim i As Integer
Dim Packet As String

Packet = "playercorpse" & SEP_CHAR & Index & SEP_CHAR & Player(Index).CorpseMap & SEP_CHAR & Player(Index).CorpseX & SEP_CHAR & Player(Index).CorpseY & SEP_CHAR & END_CHAR

Call SendDataToAll(Packet)
End Sub
Sub SendCorpseTo(ByVal Index As Integer, ByVal Target As Integer)
Dim i As Integer
Dim Packet As String

Packet = "playercorpse" & SEP_CHAR & Index & SEP_CHAR & Player(Index).CorpseMap & SEP_CHAR & Player(Index).CorpseX & SEP_CHAR & Player(Index).CorpseY & SEP_CHAR & END_CHAR

Call SendDataTo(Target, Packet)
End Sub
Sub SendAllCorpseTo(ByVal Index As Integer)
Dim i As Integer

For i = 1 To MAX_PLAYERS
If IsPlaying(i) Then
Call SendCorpseTo(i, Index)
End If
Next i
End Sub

Function CanReachCorpse(ByVal Index As Integer, ByVal Corpse As Integer) As Boolean
    Dim X As Long
    Dim y As Long

    CanReachCorpse = False

    
    If IsPlaying(Index) = False Or IsPlaying(Corpse) = False Then
        Exit Function
    End If


    ' Make sure they are on the same map
    If (GetPlayerMap(Index) = GetPlayerMap(Corpse)) Then
        X = DirToX(GetPlayerX(Index), GetPlayerDir(Index))
        y = DirToY(GetPlayerY(Index), GetPlayerDir(Index))

        If (Player(Corpse).CorpseY = y) And (Player(Corpse).CorpseX = X) Then
        CanReachCorpse = True
        End If
    End If

End Function
Sub SendUseCorpseTo(ByVal Index As Integer, ByVal Corpse As Integer)
Dim Packet As String
Dim i As Byte

Packet = "usecorpse" & SEP_CHAR & Corpse & SEP_CHAR

For i = 1 To 4
Packet = Packet & Player(Corpse).CorpseLoot(i).num & SEP_CHAR
Next i
Packet = Packet & END_CHAR

Call SendDataTo(Index, Packet)

End Sub
Sub PickUpCorpseLoot(ByVal Index As Integer, ByVal Corpse As Integer, ByVal Loot As Byte)
Dim i As Byte, a As Long


If GetPlayerMap(Index) <> Player(Corpse).CorpseMap Then Exit Sub
If Player(Corpse).CorpseLoot(Loot).num = 0 Then Exit Sub

a = Player(Corpse).CorpseLoot(Loot).num

i = FindOpenInvSlot(Index, a)
If i = 0 Then Exit Sub

Call GiveItem(Index, a, 1)
Call PlayerMsg(Index, "You looted a " & Trim$(Item(Player(Corpse).CorpseLoot(Loot).num).Name) & " !", Yellow)
Player(Corpse).CorpseLoot(Loot).num = 0
Player(Corpse).CorpseLoot(Loot).Dur = 0
Player(Corpse).CorpseLoot(Loot).Value = 0
Call SendUseCorpseTo(Index, Corpse)
End Sub

Public Sub LoadWordfilter()
    Dim i
    ReDim Wordfilter(Val(GetVar(App.Path & "\wordfilter.ini", "WORDFILTER", "maxwords")))
    If FileExist("wordfilter.ini") Then
        WordList = Val(GetVar(App.Path & "\wordfilter.ini", "WORDFILTER", "maxwords"))
        If WordList >= 1 Then
            For i = 1 To WordList
                Wordfilter(i) = LCase(GetVar(App.Path & "\wordfilter.ini", "WORDFILTER", "word" & i))
            Next i
        End If
    Else
        Call MsgBox("Wordfilter.INI could not be found. Please make sure it exists.")
        WordList = 0
    End If
End Sub

Public Function SwearCheck(TextToSay As String) As Boolean
    Dim i
    Dim SayText As String
    SayText = LCase(TextToSay)
    SwearCheck = False
    If WordList <= 0 Then Exit Function
    For i = 1 To WordList
        If InStr(1, SayText, Wordfilter(i), vbBinaryCompare) > 0 Then
            SwearCheck = True
        End If
    Next i
End Function

Sub ElementDamage(ByVal Index As Long)
Dim Damage As Long
Dim N As Long
Dim ItemNum As Long

N = Player(Index).Target

If Damage > 0 Then
  If Item(ItemNum).Element <> 0 And Npc(MapNpc(GetPlayerMap(Index), N).num).Element <> 0 Then
   If Element(Item(ItemNum).Element).Strong = Npc(MapNpc(GetPlayerMap(Index), N).num).Element Or Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Weak = Item(ItemNum).Element Then
     Call BattleMsg(Index, "     A Deadly Mix of Elements Harm The " & Trim(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & "!", BrightGreen, 0)
       Damage = Int(Damage * 1.25)
         If Element(Item(ItemNum).Element).Strong = Npc(MapNpc(GetPlayerMap(Index), N).num).Element And Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Weak = Item(ItemNum).Element Then Damage = Int(Damage * 1.2)
      End If
      End If
      End If
                                
  If Element(Item(ItemNum).Element).Weak = Npc(MapNpc(GetPlayerMap(Index), N).num).Element Or Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Strong = Item(ItemNum).Element Then
    Call BattleMsg(Index, " The " & Trim(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & " aborbs much of the elemental damage!", BrightRed, 0)
       Damage = Int(Damage * 0.75)
   If Element(Item(ItemNum).Element).Weak = Npc(MapNpc(GetPlayerMap(Index), N).num).Element And Element(Npc(MapNpc(GetPlayerMap(Index), N).num).Element).Strong = Item(ItemNum).Element Then Damage = Int(Damage * (2 / 3))
     End If


End Sub

Sub callrequstedEditQuest(ByVal Index As Long)
' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "QUESTEDITOR" & SEP_CHAR & END_CHAR)
End Sub

Sub SetupQuest(ByVal Name As String)
Dim i As Long
Dim CharNum As Long
Dim Index As Long

For i = 1 To MAX_QUESTS
Call WriteINI(Name, "QUEST" & i, "0", App.Path + "\qflag.ini")
Next i
End Sub

Sub PlayerBuyHouse(ByVal Index As Long)
Dim CharNum As Long
Dim Msg As String
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Amount As Long
Dim Damage As Long
Dim PointType As Long
Dim Movement As Long
Dim i As Long, N As Long, X As Long, y As Long, f As Long
Dim MapNum As Long
Dim s As String
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long
Dim z As Long
Dim Packet As String
Dim o As Long
Dim TempNum As Long, TempVal As Long

If GetPlayerName(Index) = Map(GetPlayerMap(Index)).Owner Then

            Map(GetPlayerMap(Index)).Owner = ""
            Map(GetPlayerMap(Index)).Name = "House For Sale"
            Call TakeItem(Index, 2, 1)
            Call GiveItem(Index, 1, 25000)
            Map(GetPlayerMap(Index)).Revision = Map(GetPlayerMap(Index)).Revision + 1
            Call QuestMsg(Index, "You Have Sold Your House for 25,000 Gold !", BrightRed)
            Call QuestMsg(Index, "Your Keys have been Removed From Your Inventory !", BrightRed)
            Call SaveMap(GetPlayerMap(Index))
            Call SendDataToAll("maphouseupdate" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & (Map(GetPlayerMap(Index)).Owner) & SEP_CHAR & (Map(GetPlayerMap(Index)).Name) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type <> TILE_TYPE_HOUSE_BUY Then
            Call QuestMsg(Index, "You need to be on a house tile to buy it!", BrightRed)
            Exit Sub
        End If

If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 = 0 Then
            Map(GetPlayerMap(Index)).Owner = GetPlayerName(Index)
            Map(GetPlayerMap(Index)).Name = GetPlayerName(Index) & "'s House"
                    Map(GetPlayerMap(Index)).Revision = Map(GetPlayerMap(Index)).Revision + 1
            Call SaveMap(GetPlayerMap(Index))
            Call SendDataToAll("maphouseupdate" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & (Map(GetPlayerMap(Index)).Owner) & SEP_CHAR & (Map(GetPlayerMap(Index)).Name) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
                If Item(GetPlayerInvItemNum(Index, i)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(Index, i) >= Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
                        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2)
                        If GetPlayerInvItemValue(Index, i) <= 0 Then
                            Call SetPlayerInvItemNum(Index, i, 0)
                        End If
                        Call GiveItem(Index, 2, 1)
                        Call QuestMsg(Index, "You have bought a new house!", BrightGreen)
                        Call QuestMsg(Index, "House Keys have been Recieved !", Yellow)
            Map(GetPlayerMap(Index)).Owner = GetPlayerName(Index)
            Map(GetPlayerMap(Index)).Name = GetPlayerName(Index) & "'s House"
                    Map(GetPlayerMap(Index)).Revision = Map(GetPlayerMap(Index)).Revision + 1
            Call SaveMap(GetPlayerMap(Index))
            Call SendDataToAll("maphouseupdate" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & (Map(GetPlayerMap(Index)).Owner) & SEP_CHAR & (Map(GetPlayerMap(Index)).Name) & SEP_CHAR & END_CHAR)
            Call SendInventory(Index)
                    End If
                Else
                    If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i Then
                        Call SetPlayerInvItemNum(Index, i, 0)
                        Call GiveItem(Index, 2, 1)
                        Call PlayerMsg(Index, "You have boughte a new house!", BrightGreen)
                        Call QuestMsg(Index, "House Keys have been Recieved !", Yellow)
            Map(GetPlayerMap(Index)).Owner = GetPlayerName(Index)
            Map(GetPlayerMap(Index)).Name = GetPlayerName(Index) & "'s House"
                    Map(GetPlayerMap(Index)).Revision = Map(GetPlayerMap(Index)).Revision + 1
            Call SaveMap(GetPlayerMap(Index))
            Call SendDataToAll("maphouseupdate" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & (Map(GetPlayerMap(Index)).Owner) & SEP_CHAR & (Map(GetPlayerMap(Index)).Name) & SEP_CHAR & END_CHAR)
            Call SendInventory(Index)
                    End If
                End If
                If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i Then
                    Exit Sub
                End If
            End If
        Next i
        
        Call QuestMsg(Index, "You dont have enough to buy this house!", BrightRed)
End Sub

Sub PerformUseItem(ByVal Index As Long, ByVal InvNum As Long, ByVal CharNum As Long)
Dim Packet As String
Dim N As Long
Dim X As Long
Dim y As Long
Dim i As Long
            ' Prevent hacking
            If InvNum < 1 Or InvNum > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid InvNum")
                Exit Sub
            End If

            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If

            If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
                N = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
                Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, n5 As Long, n6 As Long, n7 As Long, n8 As Long, n9 As Long, n10 As Long, n11 As Long, n12 As Long, n13 As Long, n14 As Long, n15 As Long, n16 As Long, n17 As Long, n18 As Long

                n1 = Item(GetPlayerInvItemNum(Index, InvNum)).StrReq
                n2 = Item(GetPlayerInvItemNum(Index, InvNum)).DefReq
                n3 = Item(GetPlayerInvItemNum(Index, InvNum)).SpeedReq
                n6 = Item(GetPlayerInvItemNum(Index, InvNum)).MagicReq
                n4 = Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq
                n5 = Item(GetPlayerInvItemNum(Index, InvNum)).AccessReq
                n7 = Item(GetPlayerInvItemNum(Index, InvNum)).LevelReq
                n8 = Item(GetPlayerInvItemNum(Index, InvNum)).BowsReq
                n10 = Item(GetPlayerInvItemNum(Index, InvNum)).LargeBladesReq
                n11 = Item(GetPlayerInvItemNum(Index, InvNum)).SmallBladesReq
                n12 = Item(GetPlayerInvItemNum(Index, InvNum)).BluntWeaponsReq
                n13 = Item(GetPlayerInvItemNum(Index, InvNum)).PoleArmsReq
                n14 = Item(GetPlayerInvItemNum(Index, InvNum)).AxesReq
                n15 = Item(GetPlayerInvItemNum(Index, InvNum)).ThrownReq
                n16 = Item(GetPlayerInvItemNum(Index, InvNum)).XbowsReq

                ' Find out what kind of item it is
                Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type

                    Case ITEM_TYPE_ARMOR

                        If InvNum <> GetPlayerArmorSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerArmorSlot(Index, InvNum)
                        Else
                            Call SetPlayerArmorSlot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                        
                        

                    Case ITEM_TYPE_WEAPON

                        If InvNum <> GetPlayerWeaponSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerBowsLevel(Index) < n8 Then
                                Call PlayerMsg(Index, "Your Bows Level needs to be higher then " & n8 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLargeBladesLevel(Index) < n10 Then
                                Call PlayerMsg(Index, "Your Large Blades Level needs to be higher then " & n10 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerSmallBladesLevel(Index) < n11 Then
                                Call PlayerMsg(Index, "Your Small Blades Level needs to be higher then " & n11 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerBluntWeaponsLevel(Index) < n12 Then
                                Call PlayerMsg(Index, "Your Blunt Weapons Level needs to be higher then " & n12 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerPolesLevel(Index) < n13 Then
                                Call PlayerMsg(Index, "Your Polearms Level needs to be higher then " & n13 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerAxesLevel(Index) < n14 Then
                                Call PlayerMsg(Index, "Your Axes Level needs to be higher then " & n14 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerThrownLevel(Index) < n15 Then
                                Call PlayerMsg(Index, "Your Thrown Weapons Level needs to be higher then " & n15 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerXbowsLevel(Index) < n16 Then
                                Call PlayerMsg(Index, "Your Xbows Level needs to be higher then " & n16 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerWeaponSlot(Index, InvNum)
                        Else
                            Call SetPlayerWeaponSlot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_HELMET

                        If InvNum <> GetPlayerHelmetSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerHelmetSlot(Index, InvNum)
                        Else
                            Call SetPlayerHelmetSlot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_SHIELD

                        If InvNum <> GetPlayerShieldSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerShieldSlot(Index, InvNum)
                        Else
                            Call SetPlayerShieldSlot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                        
                    Case ITEM_TYPE_LEGS

                        If InvNum <> GetPlayerLegsSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerLegsSlot(Index, InvNum)
                        Else
                            Call SetPlayerLegsSlot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                        
                            Case ITEM_TYPE_BOOTS

                        If InvNum <> GetPlayerBootsSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerBootsSlot(Index, InvNum)
                        Else
                            Call SetPlayerBootsSlot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                        
                        Case ITEM_TYPE_GLOVES

                        If InvNum <> GetPlayerGlovesSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerGlovesSlot(Index, InvNum)
                        Else
                            Call SetPlayerGlovesSlot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                        
                     Case ITEM_TYPE_RING1

                        If InvNum <> GetPlayerRing1Slot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerRing1Slot(Index, InvNum)
                        Else
                            Call SetPlayerRing1Slot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                        
                        Case ITEM_TYPE_RING2

                        If InvNum <> GetPlayerRing2Slot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerRing2Slot(Index, InvNum)
                        Else
                            Call SetPlayerRing2Slot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)
                        
                        Case ITEM_TYPE_AMULET

                        If InvNum <> GetPlayerAmuletSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If
                            
                            If GetPlayerLevel(Index) < n7 Then
                                Call PlayerMsg(Index, "Your Level needs to be higher then " & n7 & " to Equip This Item !", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If
                            Call SetPlayerAmuletSlot(Index, InvNum)
                        Else
                            Call SetPlayerAmuletSlot(Index, 0)
                        End If
                        Call SendInvSlots(Index)
                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_POTIONADDHP
                        Call SetPlayerHP(Index, GetPlayerHP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
                    Else
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    End If
                        Call SendHP(Index)

                    Case ITEM_TYPE_POTIONADDMP
                        Call SetPlayerMP(Index, GetPlayerMP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
                    Else
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    End If
                        Call SendMP(Index)

                    Case ITEM_TYPE_POTIONADDSP
                        Call SetPlayerSP(Index, GetPlayerSP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
                    Else
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    End If
                        Call SendSP(Index)

                    Case ITEM_TYPE_POTIONSUBHP
                        Call SetPlayerHP(Index, GetPlayerHP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
                    Else
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    End If
                        Call SendHP(Index)

                    Case ITEM_TYPE_POTIONSUBMP
                        Call SetPlayerMP(Index, GetPlayerMP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
                    Else
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    End If
                        Call SendMP(Index)

                    Case ITEM_TYPE_POTIONSUBSP
                        Call SetPlayerSP(Index, GetPlayerSP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
                    Else
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    End If
                        Call SendSP(Index)

                    Case ITEM_TYPE_KEY
                        X = DirToX(GetPlayerX(Index), GetPlayerDir(Index))
                        y = DirToY(GetPlayerY(Index), GetPlayerDir(Index))

                        If Not IsValid(X, y) Then Exit Sub

                        ' Check if a key exists
                        If Map(GetPlayerMap(Index)).Tile(X, y).Type = TILE_TYPE_KEY Then

                            ' Check if the key they are using matches the map key
                            If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(X, y).Data1 Then
                                TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = YES
                                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)

                                If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = "" Then
                                    Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", White)
                                Else
                                    Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
                                End If
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Key" & SEP_CHAR & END_CHAR)

                                ' Check if we are supposed to take away the item
                                If Map(GetPlayerMap(Index)).Tile(X, y).Data2 = 1 Then
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                    Call PlayerMsg(Index, "The key disolves.", Yellow)
                                End If
                            End If
                        End If

                        If Map(GetPlayerMap(Index)).Tile(X, y).Type = TILE_TYPE_DOOR Then
                            TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = YES
                            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Key" & SEP_CHAR & END_CHAR)
                        End If

                    Case ITEM_TYPE_SPELL

                        ' Get the spell num
                        N = Item(GetPlayerInvItemNum(Index, InvNum)).Data1

                        If N > 0 Then

                            ' Make sure they are the right class
                            If Spell(N).ClassReq = GetPlayerClass(Index) Or Spell(N).ClassReq = 0 Then
                                If Spell(N).LevelReq = 0 And Player(Index).Char(Player(Index).CharNum).Access < 1 Then
                                    Call PlayerMsg(Index, "This spell can only be used by admins!", BrightRed)
                                    Exit Sub
                                End If

                                ' Make sure they are the right level
                                i = GetSpellReqLevel(N)

                                If n6 > i Then i = n6
                                If i <= GetPlayerLevel(Index) Then
                                    i = FindOpenSpellSlot(Index)

                                    ' Make sure they have an open spell slot
                                    If i > 0 Then

                                        ' Make sure they dont already have the spell
                                        If Not HasSpell(Index, N) Then
                                            Call SetPlayerSpell(Index, i, N)
                                            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                            Call PlayerMsg(Index, "You study the spell carefully...", Yellow)
                                            Call PlayerMsg(Index, "You have learned a new spell!", White)
                                        Else
                                            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                            Call PlayerMsg(Index, "You have already learned this spell!  The spells crumbles into dust.", BrightRed)
                                        End If
                                    Else
                                        Call PlayerMsg(Index, "You have learned all that you can learn!", BrightRed)
                                    End If
                                Else
                                    Call PlayerMsg(Index, "You must be level " & i & " to learn this spell.", White)
                                End If
                            Else
                                Call PlayerMsg(Index, "This spell can only be learned by a " & GetClassName(Spell(N).ClassReq) & ".", White)
                            End If
                        Else
                            Call PlayerMsg(Index, "This scroll is not connected to a spell, please inform an admin!", White)
                        End If

                    Case ITEM_TYPE_PET
                        Player(Index).Pet.Alive = YES
                        Player(Index).Pet.Sprite = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                        Player(Index).Pet.Dir = DIR_UP
                        Player(Index).Pet.Map = GetPlayerMap(Index)
                        Player(Index).Pet.MapToGo = 0
                        Player(Index).Pet.X = GetPlayerX(Index) + Int(Rnd * 3 - 1)

                        If Player(Index).Pet.X < 0 Or Player(Index).Pet.X > MAX_MAPX Then Player(Index).Pet.X = GetPlayerX(Index)
                        Player(Index).Pet.XToGo = -1
                        Player(Index).Pet.y = GetPlayerY(Index) + Int(Rnd * 3 - 1)

                        If Player(Index).Pet.y < 0 Or Player(Index).Pet.y > MAX_MAPY Then Player(Index).Pet.y = GetPlayerY(Index)
                        Player(Index).Pet.YToGo = -1
                        Player(Index).Pet.Level = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
                        Player(Index).Pet.HP = Player(Index).Pet.Level * 5
                        Call AddToGrid(Player(Index).Pet.Map, Player(Index).Pet.X, Player(Index).Pet.y)
                        Packet = "PETDATA" & SEP_CHAR
                        Packet = Packet & Index & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.X & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.y & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
                        Packet = Packet & END_CHAR
                        Call SendDataToMap(GetPlayerMap(Index), Packet)

                        ' Excuse the ugly code, I'm rushing
                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                        Call PlayerMsg(Index, "You got a pet!", White)
                        
                        Case ITEM_TYPE_SCRIPTED
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedItem " & Index & "," & Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1
                
                        Case ITEM_TYPE_HOUSEKEY
                        'Call CaseItemTypeKey(Index)
                        X = DirToX(GetPlayerX(Index), GetPlayerDir(Index))
                        y = DirToY(GetPlayerY(Index), GetPlayerDir(Index))

                        If Not IsValid(X, y) Then Exit Sub

                        ' Check if a key exists
                        If Map(GetPlayerMap(Index)).Tile(X, y).Type = TILE_TYPE_KEY Then
                              If GetPlayerName(Index) = Map(GetPlayerMap(Index)).Owner Then
                                
                            ' Check if the key they are using matches the map key
                            If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(X, y).Data1 Then
                                TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = YES
                                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)

                                If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = "" Then
                                    Call MapMsg(GetPlayerMap(Index), "The door to your House has been unlocked!", White)
                                Else
                                    Call MapMsg(GetPlayerMap(Index), "The door to your House has been unlocked!", White)
                                End If
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Key" & SEP_CHAR & END_CHAR)

                                ' Check if we are supposed to take away the item
                                If Map(GetPlayerMap(Index)).Tile(X, y).Data2 = 1 Then
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                    Call PlayerMsg(Index, "The key disolves.", Yellow)
                                End If
                            End If
                        End If
                    End If

                        If Map(GetPlayerMap(Index)).Tile(X, y).Type = TILE_TYPE_DOOR Then
                            TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = YES
                            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Key" & SEP_CHAR & END_CHAR)
                        End If
                        
                  Case ITEM_TYPE_FOOD
                  If GetPlayerFP(Index) >= GetPlayerMaxFP(Index) Then
                Call PlayerMsg(Index, "You are Not Hungry !", BrightCyan)
                Exit Sub
                End If
                        Call SetPlayerFP(Index, GetPlayerFP(Index) + 5)
                        Call PlayerMsg(Index, "You Eat The Food ! You Regenerate 5 Hunger Points !", BrightGreen)
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
                    Else
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    End If
                        Call SendFP(Index)
                 Case ITEM_TYPE_ARROWS
                 If GetPlayerArrowsAmount(Index) > 499 Then
                Call PlayerMsg(Index, "You Cannot Carry More than 500 Arrows At A Time !", BrightRed)
                Exit Sub
                End If
                        Call SetPlayerArrowsAmount(Index, GetPlayerArrowsAmount(Index) + 50)
                        Call PlayerMsg(Index, "You Have Equipped 50 More Arrows !", BrightGreen)
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
                    Else
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    End If
                        Call SendPlayerData(Index)
                End Select
                Call SendStats(Index)
                Call SendHP(Index)
                Call SendMP(Index)
                Call SendSP(Index)
            End If
            Exit Sub
End Sub

Sub VaultVerify(ByVal Index As Long, ByVal VaultPass As String)

If VaultPass = GetPlayerVaultCode(Index) Then
Call SendDataTo(Index, "openbank" & SEP_CHAR & END_CHAR)
Call PlayerMsg(Index, "Welcome To The Bank !", Yellow)
Else
Call QuestMsg(Index, "Your Vault Code was Incorrect.", BrightRed)
Exit Sub
End If
End Sub

Sub OnDeath(ByVal Index As Long)
Dim i As Long
Dim X As Long
Dim y As Long
Dim Victim As Long
Dim MapNum As Long

Victim = Index

               If GetPlayerPK(Victim) = 0 Then
                   MapNum = GetPlayerSpawnGateMap(Victim) 'GetVar("scripts\spawngate.ini", GetPlayerName(Victim), "map")
                   y = GetPlayerSpawnGateY(Victim) 'GetVar("scripts\spawngate.ini", GetPlayerName(Victim), "y")
                   X = GetPlayerSpawnGateX(Victim) 'GetVar("scripts\spawngate.ini", GetPlayerName(Victim), "x")
                   Call PlayerWarp(Victim, MapNum, X, y)
                   Call SendPlayerData(Victim)
                   Call SendInvSlots(Victim)
                   Call SendWornEquipment(Victim)
                   Call SendStats(Victim)
                   Call PlayerMsg(Victim, "You Awaken At Your Marked Spawn Gate  !", BrightRed)
                   'Call PlayerMsg(Victim, "You Have Died " & GetPlayerDeaths(Victim) & " Time(s) !", Cyan)
                End If
              If GetPlayerPK(Victim) >= 1 Then
                 Call PlayerWarp(Victim, 203, 19, 13)
                 'frmServer.tmrJail.Enabled = True
                 Call PlayerMsg(Victim, "You Have Been Jailed for 10 Minutes !", BrightRed)
                 Call SendPlayerData(Victim)
                 Call SendInvSlots(Victim)
                 Call SendWornEquipment(Victim)
                 Call SendStats(Victim)
             End If
End Sub

Sub HungerActive(ByVal Index As Long)
If GetPlayerFP(Index) < 1 Then
Exit Sub
End If

If GetPlayerFP(Index) < 30 Then
Call PlayerMsg(Index, "Your Feeling Weak, Pherhaps You Should Eat Soon !", Cyan)
End If

Call SetPlayerFP(Index, GetPlayerFP(Index) - 1)
Call SendFP(Index)

End Sub

Sub ReplaceOneInvItem(ByVal Index As Long, olditem As Integer, newitem As Integer)
Dim N
N = 1
Do
If GetPlayerInvItemNum(Index, N) = olditem Then
Call SetPlayerInvItemNum(Index, N, newitem)
Call SendInventoryUpdate(Index, N)
Exit Do
End If
N = N + 1
Loop Until N > 24
End Sub

Sub GoMining(ByVal Index As Long, Item As Integer, maxlevel As Integer, Name As String)
Dim c As Integer
Dim Level As Integer
Level = 11

If InvGotSpace(Index, Item) = False Then
Call PlayerMsg(Index, "Sorry, You Cannot Mine, inventory full.", BrightRed)
Exit Sub
End If

If frmServer.tmrSkills.Enabled = True Then
        Call PlayerMsg(Index, "You cannot Perform This action Again So Soon !", BrightRed)
    End If

If frmServer.tmrSkills.Enabled = False Then
If GetPlayerMineLevel(Index) < 100 Then
c = Int(Rnd * Int(Level - Int(GetPlayerMineLevel(Index) / 10)))
If c = 1 Then
Call PlayerMsg(Index, GetPlayerName(Index) & " caught a " & Name, 2)
Call ReplaceOneInvItem(Index, 0, Item)
If Item <> 205 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 10)
    ElseIf Item <> 206 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 25)
    ElseIf Item <> 207 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 50)
    ElseIf Item <> 208 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 100)
    ElseIf Item <> 209 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 225)
    End If
    Call SendPlayerData(Index)
    Call CheckPlayerMineLevelUp(Index)
    frmServer.tmrSkills.Enabled = True
Else
Call PlayerMsg(Index, GetPlayerName(Index) & " found nothing!", 12)
End If
Else
Call PlayerMsg(Index, GetPlayerName(Index) & " caught a " & Name, 2)
Call ReplaceOneInvItem(Index, 0, Item)
If Item <> 205 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 10)
    ElseIf Item <> 206 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 25)
    ElseIf Item <> 207 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 50)
    ElseIf Item <> 208 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 100)
    ElseIf Item <> 209 Then
        Call SetPlayerMineExp(Index, GetPlayerMineExp(Index) + 225)
    End If
    Call SendPlayerData(Index)
    Call CheckPlayerMineLevelUp(Index)
    frmServer.tmrSkills.Enabled = True
End If
End If
End Sub

Sub GoLJacking(ByVal Index As Long, Item As Integer, maxlevel As Integer, Name As String)
Dim c As Integer
Dim Level As Integer
Level = 11


If InvGotSpace(Index, Item) = False Then
Call PlayerMsg(Index, "Sorry, You cannot Chop anymore Wood, inventory full.", BrightRed)
Exit Sub
End If

If frmServer.tmrSkills.Enabled = True Then
        Call PlayerMsg(Index, "You cannot Perform This action Again So Soon !", BrightRed)
    End If

If frmServer.tmrSkills.Enabled = False Then
If GetPlayerLJackingLevel(Index) < 100 Then
c = Int(Rnd * Int(Level - Int(GetPlayerLJackingLevel(Index) / 10)))
If c = 1 Then
Call PlayerMsg(Index, GetPlayerName(Index) & " Obtained a " & Name, 2)
Call ReplaceOneInvItem(Index, 0, Item)
If Item <> 205 Then
        Call RequestLJackingExp(Index)
    ElseIf Item <> 206 Then
        Call RequestLJackingExp(Index)
    ElseIf Item <> 207 Then
        Call RequestLJackingExp(Index)
    ElseIf Item <> 208 Then
        Call RequestLJackingExp(Index)
    ElseIf Item <> 209 Then
        Call RequestLJackingExp(Index)
    End If
    Call SendPlayerData(Index)
    Call CheckPlayerLJackingLevelUp(Index)
    frmServer.tmrSkills.Enabled = True
Else
Call PlayerMsg(Index, GetPlayerName(Index) & " found nothing!", 12)
End If
Else
Call PlayerMsg(Index, GetPlayerName(Index) & " Obtained a " & Name, 2)
Call ReplaceOneInvItem(Index, 0, Item)
If Item <> 205 Then
        Call RequestLJackingExp(Index)
    ElseIf Item <> 206 Then
        Call RequestLJackingExp(Index)
    ElseIf Item <> 207 Then
        Call RequestLJackingExp(Index)
    ElseIf Item <> 208 Then
        Call RequestLJackingExp(Index)
    ElseIf Item <> 209 Then
        Call RequestLJackingExp(Index)
    End If
    Call SendPlayerData(Index)
    Call CheckPlayerLJackingLevelUp(Index)
    frmServer.tmrSkills.Enabled = True
End If
End If
End Sub

Sub GoFishing(ByVal Index As Long, Item As Integer, maxlevel As Integer, Name As String)
Dim c As Integer
Dim Level As Integer
Level = 11

If InvGotSpace(Index, Item) = False Then
Call PlayerMsg(Index, "Sorry, cant catch fish, inventory full.", BrightRed)
Exit Sub
End If

If frmServer.tmrSkills.Enabled = True Then
        Call PlayerMsg(Index, "You cannot Perform This action Again So Soon !", BrightRed)
    End If
    
If frmServer.tmrSkills.Enabled = False Then
If GetPlayerFishLevel(Index) < 100 Then
c = Int(Rnd * Int(Level - Int(GetPlayerFishLevel(Index) / 10)))
If c = 1 Then
Call PlayerMsg(Index, GetPlayerName(Index) & " caught a " & Name, 2)
Call ReplaceOneInvItem(Index, 0, Item)
If Item <> 205 Then
         Call RequestFishingExp(Index)
        'Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 10)
    ElseIf Item <> 206 Then
         Call RequestFishingExp(Index)
        'Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 25)
    ElseIf Item <> 207 Then
         Call RequestFishingExp(Index)
        'Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 50)
    ElseIf Item <> 208 Then
         Call RequestFishingExp(Index)
        'Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 100)
    ElseIf Item <> 209 Then
         Call RequestFishingExp(Index)
        'Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 225)
    End If
    Call SendPlayerData(Index)
    Call CheckPlayerFishLevelUp(Index)
    frmServer.tmrSkills.Enabled = True
Else
Call PlayerMsg(Index, GetPlayerName(Index) & " found nothing!", 12)
End If
Else
Call PlayerMsg(Index, GetPlayerName(Index) & " caught a " & Name, 2)
Call ReplaceOneInvItem(Index, 0, Item)
If Item <> 205 Then
        Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 10)
    ElseIf Item <> 206 Then
        Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 25)
    ElseIf Item <> 207 Then
        Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 50)
    ElseIf Item <> 208 Then
        Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 100)
    ElseIf Item <> 209 Then
        Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 225)
    End If
    Call SendPlayerData(Index)
    Call CheckPlayerFishLevelUp(Index)
    frmServer.tmrSkills.Enabled = True
End If
End If
End Sub

Sub CheckPlayerFishLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerFishExp(Index) >= GetPlayerNextFishLevel(Index) Then
        If GetPlayerFishLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerFishExp(Index) < GetPlayerNextFishLevel(Index)
                    DoEvents

                    If GetPlayerFishLevel(Index) < MAX_LEVEL Then
                        If GetPlayerFishExp(Index) >= GetPlayerNextFishLevel(Index) Then
                            d = GetPlayerFishExp(Index) - GetPlayerNextFishLevel(Index)
                            Call SetPlayerFishLevel(Index, GetPlayerFishLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerFishExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " fishing levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a fishing level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerFishLevel(Index) = MAX_LEVEL Then
            Call SetPlayerFishExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub CheckPlayerMineLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerMineExp(Index) >= GetPlayerNextMineLevel(Index) Then
        If GetPlayerMineLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerMineExp(Index) < GetPlayerNextMineLevel(Index)
                    DoEvents

                    If GetPlayerMineLevel(Index) < MAX_LEVEL Then
                        If GetPlayerMineExp(Index) >= GetPlayerNextMineLevel(Index) Then
                            d = GetPlayerMineExp(Index) - GetPlayerNextMineLevel(Index)
                            Call SetPlayerMineLevel(Index, GetPlayerMineLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerMineExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " mining levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a mining level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerMineLevel(Index) = MAX_LEVEL Then
            Call SetPlayerMineExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub CheckPlayerLJackingLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerLJackingExp(Index) >= GetPlayerNextLJackingLevel(Index) Then
        If GetPlayerLJackingLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerLJackingExp(Index) < GetPlayerNextLJackingLevel(Index)
                    DoEvents

                    If GetPlayerFishLevel(Index) < MAX_LEVEL Then
                        If GetPlayerLJackingExp(Index) >= GetPlayerNextLJackingLevel(Index) Then
                            d = GetPlayerLJackingExp(Index) - GetPlayerNextLJackingLevel(Index)
                            Call SetPlayerLJackingLevel(Index, GetPlayerLJackingLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerLJackingExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " Lumber Jacking levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a Lumber Jacking level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerLJackingLevel(Index) = MAX_LEVEL Then
            Call SetPlayerLJackingExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub RequestLJackingExp(ByVal Index As Long)

If GetPlayerLJackingLevel(Index) = 1 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 2 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 3 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 4 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 5 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 6 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 7 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 8 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 9 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 10 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 11 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 12 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 13 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 14 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 15 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 16 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 17 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 18 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 19 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 20 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 21 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 22 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 23 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 24 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 25 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 26 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 27 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 28 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 29 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 30 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 31 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 32 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 33 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 34 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 35 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 36 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 37 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 38 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 39 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 40 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 41 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 42 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 43 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 44 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 45 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 46 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 47 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 48 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 49 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 50 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 51 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 52 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 53 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 54 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 55 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 56 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 57 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 58 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 59 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 60 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 61 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 62 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 63 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 64 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 65 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 66 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 67 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 68 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 69 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 70 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 71 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 72 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 73 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 74 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 75 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 76 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 77 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 78 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 79 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 80 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 81 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 82 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 83 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 84 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 85 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 86 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 87 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 88 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 89 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 90 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 91 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 92 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 93 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 94 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 95 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 96 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 97 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 98 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 99 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 100 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 101 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 102 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 103 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 104 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 105 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 106 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 107 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 108 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 109 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Lumber Jacking Level !", BrightGreen)
ElseIf GetPlayerLJackingLevel(Index) = 110 Then
Call SetPlayerLJackingExp(Index, GetPlayerLJackingExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Lumber Jacking Level !", BrightGreen)
End If
End Sub

Sub RequestFishingExp(ByVal Index As Long)

If GetPlayerFishLevel(Index) = 1 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 2 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 3 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 4 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 5 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 6 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 7 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 8 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 9 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 10 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 11 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 12 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 13 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 14 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 15 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 16 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 17 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 18 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 19 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 20 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 21 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 22 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 23 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 24 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 25 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 26 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 27 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 28 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 29 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 30 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 31 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 32 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 33 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 34 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 35 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 36 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 37 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 38 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 39 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 40 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 41 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 42 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 43 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 44 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 45 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 46 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 47 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 48 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 49 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 50 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 51 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 52 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 53 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 54 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 55 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 56 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 57 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 58 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 59 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 60 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 61 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 62 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 63 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 64 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 65 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 66 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 67 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 68 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 69 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 70 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 71 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 72 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 73 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 74 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 75 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 76 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 77 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 78 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 79 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 80 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 81 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 82 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 83 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 84 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 85 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 86 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 87 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 88 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 89 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 90 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 91 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 92 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 93 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 94 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 95 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 96 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 97 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 98 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 99 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 100 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 101 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 102 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 103 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 104 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 105 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 106 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 107 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 108 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 109 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Fishing Level !", BrightGreen)
ElseIf GetPlayerFishLevel(Index) = 110 Then
Call SetPlayerFishExp(Index, GetPlayerFishExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Fishing Level !", BrightGreen)
End If
End Sub

Function InvGotSpace(ByVal Index As Long, Item As Integer) As Boolean
Dim X
X = FindOpenInvSlot(Index, Item)

If Not X <> 0 Then
InvGotSpace = False
Exit Function
End If

InvGotSpace = True

End Function

'Generates random numbers
Public Function RandomNo(Max As Long, Optional Last As Integer) As Long
Dim a, b
If Val(Last) < 1 Then Last = 100


If Max < 1 Then
RandomNo = 0
Exit Function
End If

a = Rnd(Last)
b = Mid(a, InStr(1, a, ".", vbTextCompare) + 1, Len(STR(Max)) - 1)


If b > Max Then
b = b - ((b \ Max) * Max)
End If

If b < 1 Then b = 0
RandomNo = b
End Function

Sub CheckPlayerAxesLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerAxesExp(Index) >= GetPlayerNextAxesLevel(Index) Then
        If GetPlayerAxesLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerAxesExp(Index) < GetPlayerNextAxesLevel(Index)
                    DoEvents

                    If GetPlayerAxesLevel(Index) < MAX_LEVEL Then
                        If GetPlayerAxesExp(Index) >= GetPlayerNextAxesLevel(Index) Then
                            d = GetPlayerAxesExp(Index) - GetPlayerNextAxesLevel(Index)
                            Call SetPlayerAxesLevel(Index, GetPlayerAxesLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerAxesExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " Axes levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a Axes level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerAxesLevel(Index) = MAX_LEVEL Then
            Call SetPlayerAxesExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub GoAxes(ByVal Index As Long)
Dim c As Integer
Dim Level As Integer
Dim xp As Integer
Level = 11
xp = 100
    
If GetPlayerAxesLevel(Index) <= 100 Then
c = 1
If c = 1 Then

        Call RequestAxesExp(Index)

    Call SendPlayerData(Index)
    Call CheckPlayerAxesLevelUp(Index)
End If
End If
End Sub
Sub RequestAxesExp(ByVal Index As Long)

If GetPlayerAxesLevel(Index) = 1 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 2 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 3 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 4 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 5 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 6 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 7 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 8 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 9 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 10 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 11 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 12 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 13 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 14 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 15 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 16 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 17 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 18 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 19 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 20 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 21 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 22 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 23 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 24 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 25 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 26 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 27 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 28 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 29 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 30 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 31 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 32 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 33 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 34 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 35 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 36 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 37 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 38 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 39 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 40 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 41 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 42 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 43 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 44 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 45 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 46 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 47 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 48 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 49 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 50 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 51 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 52 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 53 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 54 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 55 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 56 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 57 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 58 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 59 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 60 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 61 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 62 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 63 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 64 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 65 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 66 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 67 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 68 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 69 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 70 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 71 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 72 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 73 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 74 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 75 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 76 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 77 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 78 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 79 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 80 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 81 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 82 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 83 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 84 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 85 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 86 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 87 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 88 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 89 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 90 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 91 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 92 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 93 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 94 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 95 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 96 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 97 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 98 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 99 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 100 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 101 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 102 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 103 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 104 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 105 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 106 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 107 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 108 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 109 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Axes Level !", BrightGreen)
ElseIf GetPlayerAxesLevel(Index) = 110 Then
Call SetPlayerAxesExp(Index, GetPlayerAxesExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Axes Level !", BrightGreen)
End If
End Sub

Sub CheckPlayerThrownLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerThrownExp(Index) >= GetPlayerNextThrownLevel(Index) Then
        If GetPlayerThrownLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerThrownExp(Index) < GetPlayerNextThrownLevel(Index)
                    DoEvents

                    If GetPlayerThrownLevel(Index) < MAX_LEVEL Then
                        If GetPlayerThrownExp(Index) >= GetPlayerNextThrownLevel(Index) Then
                            d = GetPlayerThrownExp(Index) - GetPlayerNextThrownLevel(Index)
                            Call SetPlayerThrownLevel(Index, GetPlayerThrownLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerThrownExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " Thrown levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a Thrown level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerThrownLevel(Index) = MAX_LEVEL Then
            Call SetPlayerThrownExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub GoThrown(ByVal Index As Long)
Dim c As Integer
Dim Level As Integer
Dim xp As Integer
Level = 11
xp = 174
    
If GetPlayerThrownLevel(Index) <= 100 Then
c = 1
If c = 1 Then

        Call RequestThrownExp(Index)

    Call SendPlayerData(Index)
    Call CheckPlayerThrownLevelUp(Index)
End If
End If
End Sub

Sub RequestThrownExp(ByVal Index As Long)

If GetPlayerThrownLevel(Index) = 1 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 2 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 3 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 4 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 5 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 6 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 7 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 8 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 9 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 10 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 11 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 12 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 13 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 14 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 15 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 16 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 17 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 18 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 19 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 20 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 21 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 22 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 23 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 24 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 25 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 26 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 27 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 28 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 29 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 30 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 31 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 32 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 33 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 34 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 35 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 36 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 37 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 38 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 39 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 40 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 41 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 42 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 43 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 44 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 45 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 46 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 47 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 48 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 49 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 50 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 51 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 52 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 53 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 54 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 55 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 56 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 57 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 58 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 59 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 60 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 61 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 62 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 63 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 64 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 65 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 66 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 67 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 68 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 69 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 70 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 71 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 72 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 73 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 74 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 75 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 76 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 77 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 78 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 79 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 80 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 81 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 82 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 83 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 84 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 85 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 86 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 87 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 88 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 89 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 90 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 91 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 92 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 93 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 94 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 95 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 96 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 97 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 98 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 99 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 100 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 101 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 102 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 103 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 104 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 105 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 106 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 107 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 108 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 109 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Thrown Level !", BrightGreen)
ElseIf GetPlayerThrownLevel(Index) = 110 Then
Call SetPlayerThrownExp(Index, GetPlayerThrownExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Thrown Level !", BrightGreen)
End If
End Sub

Sub CheckPlayerXbowsLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerXbowsExp(Index) >= GetPlayerNextXbowsLevel(Index) Then
        If GetPlayerXbowsLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerXbowsExp(Index) < GetPlayerNextXbowsLevel(Index)
                    DoEvents

                    If GetPlayerXbowsLevel(Index) < MAX_LEVEL Then
                        If GetPlayerXbowsExp(Index) >= GetPlayerNextXbowsLevel(Index) Then
                            d = GetPlayerXbowsExp(Index) - GetPlayerNextXbowsLevel(Index)
                            Call SetPlayerXbowsLevel(Index, GetPlayerXbowsLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerXbowsExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " Xbows levels!", 6)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a Xbows level!", 6)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerXbowsLevel(Index) = MAX_LEVEL Then
            Call SetPlayerXbowsExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub GoXbows(ByVal Index As Long)
Dim c As Integer
Dim Level As Integer
Dim xp As Integer
Level = 11
xp = 159

If GetPlayerXbowsLevel(Index) <= 100 Then
c = 1
If c = 1 Then

        Call RequestXbowsExp(Index)

    Call SendPlayerData(Index)
    Call CheckPlayerXbowsLevelUp(Index)
End If
End If
End Sub

Sub RequestXbowsExp(ByVal Index As Long)

If GetPlayerXbowsLevel(Index) = 1 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 2 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 3 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 4 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 5 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 6 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 7 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 8 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 9 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 10 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 11 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 12 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 13 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 14 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 15 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 16 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 17 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 18 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 19 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 20 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 21 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 22 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 23 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 24 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 25 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 26 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 27 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 28 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 29 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 30 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 31 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 32 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 33 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 34 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 35 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 36 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 37 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 38 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 39 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 40 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 41 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 42 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 43 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 44 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 45 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 46 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 47 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 48 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 49 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 50 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 51 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 52 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 53 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 54 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 55 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 56 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 57 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 58 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 59 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 60 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 61 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 62 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 63 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 64 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 65 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 66 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 67 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 68 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 69 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 70 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 71 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 72 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 73 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 74 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 75 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 76 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 77 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 78 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 79 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 80 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 81 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 82 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 83 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 84 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 85 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 86 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 87 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 88 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 89 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 90 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 91 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 92 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 93 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 94 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 95 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 96 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 97 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 98 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 99 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 100 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 101 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 102 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 103 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 104 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 105 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 106 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 107 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 108 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 109 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Xbows Level !", BrightGreen)
ElseIf GetPlayerXbowsLevel(Index) = 110 Then
Call SetPlayerXbowsExp(Index, GetPlayerXbowsExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Xbows Level !", BrightGreen)
End If
End Sub

Sub CheckPlayerBowsLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerBowsExp(Index) >= GetPlayerNextBowsLevel(Index) Then
        If GetPlayerBowsLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerBowsExp(Index) < GetPlayerNextBowsLevel(Index)
                    DoEvents

                    If GetPlayerBowsLevel(Index) < MAX_LEVEL Then
                        If GetPlayerBowsExp(Index) >= GetPlayerNextBowsLevel(Index) Then
                            d = GetPlayerBowsExp(Index) - GetPlayerNextBowsLevel(Index)
                            Call SetPlayerBowsLevel(Index, GetPlayerBowsLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerBowsExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " Bows levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a Bows level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerBowsLevel(Index) = MAX_LEVEL Then
            Call SetPlayerBowsExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub GoBows(ByVal Index As Long)
Dim c As Integer
Dim Level As Integer
Dim xp As Integer
Level = 11
xp = 192
    
If GetPlayerBowsLevel(Index) <= 100 Then
c = 1
If c = 1 Then

        Call RequestBowsExp(Index)
        
    Call SendPlayerData(Index)
    Call CheckPlayerBowsLevelUp(Index)
End If
End If
End Sub

Sub RequestBowsExp(ByVal Index As Long)

If GetPlayerBowsLevel(Index) = 1 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 2 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 3 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 4 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 5 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 6 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 7 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 8 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 9 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 10 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 11 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 12 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 13 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 14 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 15 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 16 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 17 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 18 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 19 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 20 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 21 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 22 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 23 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 24 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 25 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 26 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 27 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 28 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 29 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 30 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 31 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 32 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 33 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 34 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 35 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 36 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 37 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 38 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 39 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 40 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 41 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 42 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 43 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 44 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 45 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 46 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 47 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 48 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 49 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 50 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 51 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 52 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 53 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 54 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 55 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 56 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 57 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 58 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 59 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 60 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 61 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 62 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 63 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 64 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 65 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 66 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 67 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 68 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 69 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 70 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 71 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 72 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 73 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 74 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 75 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 76 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 77 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 78 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 79 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 80 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 81 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 82 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 83 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 84 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 85 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 86 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 87 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 88 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 89 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 90 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 91 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 92 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 93 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 94 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 95 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 96 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 97 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 98 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 99 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 100 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 101 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 102 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 103 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 104 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 105 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 106 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 107 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 108 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 109 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Bows Level !", BrightGreen)
ElseIf GetPlayerBowsLevel(Index) = 110 Then
Call SetPlayerBowsExp(Index, GetPlayerBowsExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Bows Level !", BrightGreen)
End If
End Sub

Sub GoLargeBlades(ByVal Index As Long)
Dim c As Integer
Dim Level As Integer
Dim xp As Integer
Level = 11
xp = 100


If GetPlayerLargeBladesLevel(Index) <= 100 Then
c = 1
If c = 1 Then

       Call RequestLargeBladesExp(Index)

    Call SendPlayerData(Index)
    Call CheckPlayerLargeBladesLevelUp(Index)
End If
End If
End Sub

Sub RequestLargeBladesExp(ByVal Index As Long)

If GetPlayerLargeBladesLevel(Index) = 1 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 2 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 3 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 4 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 5 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 6 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 7 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 8 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 9 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 10 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 11 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 12 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 13 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 14 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 15 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 16 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 17 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 18 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 19 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 20 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 21 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 22 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 23 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 24 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 25 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 26 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 27 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 28 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 29 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 30 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 31 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 32 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 33 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 34 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 35 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 36 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 37 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 38 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 39 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 40 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 41 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 42 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 43 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 44 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 45 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 46 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 47 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 48 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 49 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 50 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 51 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 52 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 53 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 54 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 55 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 56 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 57 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 58 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 59 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 60 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 61 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 62 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 63 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 64 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 65 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 66 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 67 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 68 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 69 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 70 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 71 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 72 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 73 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 74 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 75 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 76 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 77 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 78 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 79 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 80 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 81 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 82 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 83 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 84 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 85 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 86 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 87 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 88 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 89 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 90 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 91 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 92 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 93 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 94 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 95 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 96 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 97 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 98 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 99 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 100 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 101 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 102 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 103 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 104 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 105 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 106 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 107 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 108 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 109 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Large Blades Level !", BrightGreen)
ElseIf GetPlayerLargeBladesLevel(Index) = 110 Then
Call SetPlayerLargeBladesExp(Index, GetPlayerLargeBladesExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Large Blades Level !", BrightGreen)
End If
End Sub

Sub CheckPlayerSmallBladesLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerSmallBladesExp(Index) >= GetPlayerNextSmallBladesLevel(Index) Then
        If GetPlayerSmallBladesLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerSmallBladesExp(Index) < GetPlayerNextSmallBladesLevel(Index)
                    DoEvents

                    If GetPlayerSmallBladesLevel(Index) < MAX_LEVEL Then
                        If GetPlayerSmallBladesExp(Index) >= GetPlayerNextSmallBladesLevel(Index) Then
                            d = GetPlayerSmallBladesExp(Index) - GetPlayerNextSmallBladesLevel(Index)
                            Call SetPlayerSmallBladesLevel(Index, GetPlayerSmallBladesLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerSmallBladesExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " SmallBlades levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a SmallBlades level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerSmallBladesLevel(Index) = MAX_LEVEL Then
            Call SetPlayerSmallBladesExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub GoSmallBlades(ByVal Index As Long)
Dim c As Integer
Dim Level As Integer
Dim xp As Integer
Level = 11
xp = 131
    
If GetPlayerSmallBladesLevel(Index) <= 100 Then
c = 1
If c = 1 Then

        Call RequestSmallBladesExp(Index)

    Call SendPlayerData(Index)
    Call CheckPlayerSmallBladesLevelUp(Index)
End If
End If
End Sub

Sub RequestSmallBladesExp(ByVal Index As Long)

If GetPlayerSmallBladesLevel(Index) = 1 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 2 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 3 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 4 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 5 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 6 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 7 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 8 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 9 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 10 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 11 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 12 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 13 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 14 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 15 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 16 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 17 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 18 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 19 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 20 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 21 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 22 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 23 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 24 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 25 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 26 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 27 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 28 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 29 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 30 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 31 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 32 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 33 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 34 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 35 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 36 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 37 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 38 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 39 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 40 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 41 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 42 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 43 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 44 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 45 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 46 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 47 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 48 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 49 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 50 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 51 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 52 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 53 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 54 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 55 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 56 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 57 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 58 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 59 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 60 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 61 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 62 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 63 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 64 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 65 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 66 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 67 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 68 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 69 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 70 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 71 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 72 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 73 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 74 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 75 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 76 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 77 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 78 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 79 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 80 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 81 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 82 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 83 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 84 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 85 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 86 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 87 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 88 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 89 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 90 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 91 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 92 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 93 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 94 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 95 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 96 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 97 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 98 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 99 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 100 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 101 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 102 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 103 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 104 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 105 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 106 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 107 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 108 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 109 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Small Blades Level !", BrightGreen)
ElseIf GetPlayerSmallBladesLevel(Index) = 110 Then
Call SetPlayerSmallBladesExp(Index, GetPlayerSmallBladesExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Small Blades Level !", BrightGreen)
End If
End Sub

Sub CheckPlayerBluntWeaponsLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerBluntWeaponsExp(Index) >= GetPlayerNextBluntWeaponsLevel(Index) Then
        If GetPlayerBluntWeaponsLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerBluntWeaponsExp(Index) < GetPlayerNextBluntWeaponsLevel(Index)
                    DoEvents

                    If GetPlayerBluntWeaponsLevel(Index) < MAX_LEVEL Then
                        If GetPlayerBluntWeaponsExp(Index) >= GetPlayerNextBluntWeaponsLevel(Index) Then
                            d = GetPlayerBluntWeaponsExp(Index) - GetPlayerNextBluntWeaponsLevel(Index)
                            Call SetPlayerBluntWeaponsLevel(Index, GetPlayerBluntWeaponsLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerBluntWeaponsExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " BluntWeapons levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a BluntWeapons level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerBluntWeaponsLevel(Index) = MAX_LEVEL Then
            Call SetPlayerBluntWeaponsExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub GoBluntWeapons(ByVal Index As Long)
Dim c As Integer
Dim Level As Integer
Dim xp As Integer
Level = 11
xp = 153
    
If GetPlayerBluntWeaponsLevel(Index) <= 100 Then
c = 1
If c = 1 Then

        Call RequestBluntWeaponsExp(Index)

    Call SendPlayerData(Index)
    Call CheckPlayerBluntWeaponsLevelUp(Index)
End If
End If
End Sub

Sub RequestBluntWeaponsExp(ByVal Index As Long)

If GetPlayerBluntWeaponsLevel(Index) = 1 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 2 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 3 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 4 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 5 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 6 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 7 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 8 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 9 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 10 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 11 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 12 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 13 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 14 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 15 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 16 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 17 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 18 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 19 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 20 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 21 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 22 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 23 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 24 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 25 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 26 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 27 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 28 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 29 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 30 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 31 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 32 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 33 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 34 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 35 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 36 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 37 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 38 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 39 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 40 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 41 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 42 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 43 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 44 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 45 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 46 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 47 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 48 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 49 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 50 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 51 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 52 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 53 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 54 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 55 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 56 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 57 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 58 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 59 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 60 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 61 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 62 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 63 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 64 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 65 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 66 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 67 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 68 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 69 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 70 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 71 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 72 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 73 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 74 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 75 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 76 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 77 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 78 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 79 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 80 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 81 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 82 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 83 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 84 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 85 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 86 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 87 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 88 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 89 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 90 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 91 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 92 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 93 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 94 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 95 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 96 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 97 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 98 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 99 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 100 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 101 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 102 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 103 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 104 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 105 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 106 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 107 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 108 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 109 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Blunt Weapons Level !", BrightGreen)
ElseIf GetPlayerBluntWeaponsLevel(Index) = 110 Then
Call SetPlayerBluntWeaponsExp(Index, GetPlayerBluntWeaponsExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Blunt Weapons Level !", BrightGreen)
End If
End Sub

Sub CheckPlayerPolesLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerPolesExp(Index) >= GetPlayerNextPolesLevel(Index) Then
        If GetPlayerPolesLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerPolesExp(Index) < GetPlayerNextPolesLevel(Index)
                    DoEvents

                    If GetPlayerPolesLevel(Index) < MAX_LEVEL Then
                        If GetPlayerPolesExp(Index) >= GetPlayerNextPolesLevel(Index) Then
                            d = GetPlayerPolesExp(Index) - GetPlayerNextPolesLevel(Index)
                            Call SetPlayerPolesLevel(Index, GetPlayerPolesLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerPolesExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " Poles levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a Poles level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerPolesLevel(Index) = MAX_LEVEL Then
            Call SetPlayerPolesExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub

Sub GoPoles(ByVal Index As Long)
Dim c As Integer
Dim Level As Integer
Dim xp As Integer
Level = 11
xp = 100
    
If GetPlayerPolesLevel(Index) <= 100 Then
c = 1
If c = 1 Then

        Call RequestPolesExp(Index)

    Call SendPlayerData(Index)
    Call CheckPlayerPolesLevelUp(Index)
End If
End If
End Sub

Sub RequestPolesExp(ByVal Index As Long)

If GetPlayerPolesLevel(Index) = 1 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 50)
Call PlayerMsg(Index, "You have Gained 50 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 2 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 54)
Call PlayerMsg(Index, "You have Gained 54 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 3 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 58)
Call PlayerMsg(Index, "You have Gained 58 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 4 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 62)
Call PlayerMsg(Index, "You have Gained 62 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 5 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 66)
Call PlayerMsg(Index, "You have Gained 66 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 6 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 73)
Call PlayerMsg(Index, "You have Gained 73 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 7 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 77)
Call PlayerMsg(Index, "You have Gained 77 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 8 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 81)
Call PlayerMsg(Index, "You have Gained 81 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 9 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 86)
Call PlayerMsg(Index, "You have Gained 86 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 10 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 91)
Call PlayerMsg(Index, "You have Gained 91 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 11 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 96)
Call PlayerMsg(Index, "You have Gained 96 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 12 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 101)
Call PlayerMsg(Index, "You have Gained 101 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 13 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 107)
Call PlayerMsg(Index, "You have Gained 107 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 14 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 112)
Call PlayerMsg(Index, "You have Gained 112 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 15 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 117)
Call PlayerMsg(Index, "You have Gained 117 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 16 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 121)
Call PlayerMsg(Index, "You have Gained 121 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 17 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 126)
Call PlayerMsg(Index, "You have Gained 126 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 18 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 132)
Call PlayerMsg(Index, "You have Gained 132 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 19 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 137)
Call PlayerMsg(Index, "You have Gained 137 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 20 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 140)
Call PlayerMsg(Index, "You have Gained 140 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 21 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 144)
Call PlayerMsg(Index, "You have Gained 144 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 22 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 149)
Call PlayerMsg(Index, "You have Gained 149 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 23 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 153)
Call PlayerMsg(Index, "You have Gained 153 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 24 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 158)
Call PlayerMsg(Index, "You have Gained 158 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 25 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 162)
Call PlayerMsg(Index, "You have Gained 162 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 26 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 167)
Call PlayerMsg(Index, "You have Gained 167 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 27 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 172)
Call PlayerMsg(Index, "You have Gained 172 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 28 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 176)
Call PlayerMsg(Index, "You have Gained 176 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 29 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 180)
Call PlayerMsg(Index, "You have Gained 180 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 30 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 184)
Call PlayerMsg(Index, "You have Gained 184 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 31 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 189)
Call PlayerMsg(Index, "You have Gained 189 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 32 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 193)
Call PlayerMsg(Index, "You have Gained 193 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 33 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 197)
Call PlayerMsg(Index, "You have Gained 197 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 34 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 201)
Call PlayerMsg(Index, "You have Gained 201 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 35 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 204)
Call PlayerMsg(Index, "You have Gained 204 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 36 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 209)
Call PlayerMsg(Index, "You have Gained 209 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 37 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 213)
Call PlayerMsg(Index, "You have Gained 213 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 38 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 218)
Call PlayerMsg(Index, "You have Gained 218 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 39 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 221)
Call PlayerMsg(Index, "You have Gained 221 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 40 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 225)
Call PlayerMsg(Index, "You have Gained 225 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 41 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 228)
Call PlayerMsg(Index, "You have Gained 228 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 42 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 232)
Call PlayerMsg(Index, "You have Gained 232 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 43 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 236)
Call PlayerMsg(Index, "You have Gained 236 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 44 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 239)
Call PlayerMsg(Index, "You have Gained 239 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 45 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 244)
Call PlayerMsg(Index, "You have Gained 244 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 46 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 249)
Call PlayerMsg(Index, "You have Gained 249 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 47 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 253)
Call PlayerMsg(Index, "You have Gained 253 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 48 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 257)
Call PlayerMsg(Index, "You have Gained 257 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 49 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 261)
Call PlayerMsg(Index, "You have Gained 261 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 50 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 266)
Call PlayerMsg(Index, "You have Gained 266 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 51 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 270)
Call PlayerMsg(Index, "You have Gained 270 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 52 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 273)
Call PlayerMsg(Index, "You have Gained 273 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 53 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 277)
Call PlayerMsg(Index, "You have Gained 277 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 54 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 281)
Call PlayerMsg(Index, "You have Gained 281 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 55 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 285)
Call PlayerMsg(Index, "You have Gained 285 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 56 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 289)
Call PlayerMsg(Index, "You have Gained 289 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 57 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 293)
Call PlayerMsg(Index, "You have Gained 293 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 58 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 297)
Call PlayerMsg(Index, "You have Gained 297 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 59 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 298)
Call PlayerMsg(Index, "You have Gained 298 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 60 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 303)
Call PlayerMsg(Index, "You have Gained 303 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 61 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 306)
Call PlayerMsg(Index, "You have Gained 306 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 62 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 310)
Call PlayerMsg(Index, "You have Gained 310 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 63 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 314)
Call PlayerMsg(Index, "You have Gained 314 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 64 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 320)
Call PlayerMsg(Index, "You have Gained 320 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 65 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 323)
Call PlayerMsg(Index, "You have Gained 323 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 66 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 327)
Call PlayerMsg(Index, "You have Gained 327 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 67 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 331)
Call PlayerMsg(Index, "You have Gained 331 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 68 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 335)
Call PlayerMsg(Index, "You have Gained 335 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 69 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 339)
Call PlayerMsg(Index, "You have Gained 339 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 70 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 343)
Call PlayerMsg(Index, "You have Gained 343 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 71 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 347)
Call PlayerMsg(Index, "You have Gained 347 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 72 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 351)
Call PlayerMsg(Index, "You have Gained 351 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 73 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 355)
Call PlayerMsg(Index, "You have Gained 355 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 74 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 358)
Call PlayerMsg(Index, "You have Gained 358 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 75 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 362)
Call PlayerMsg(Index, "You have Gained 362 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 76 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 366)
Call PlayerMsg(Index, "You have Gained 366 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 77 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 370)
Call PlayerMsg(Index, "You have Gained 370 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 78 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 373)
Call PlayerMsg(Index, "You have Gained 373 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 79 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 377)
Call PlayerMsg(Index, "You have Gained 377 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 80 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 382)
Call PlayerMsg(Index, "You have Gained 382 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 81 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 386)
Call PlayerMsg(Index, "You have Gained 386 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 82 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 390)
Call PlayerMsg(Index, "You have Gained 390 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 83 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 393)
Call PlayerMsg(Index, "You have Gained 393 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 84 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 397)
Call PlayerMsg(Index, "You have Gained 397 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 85 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 401)
Call PlayerMsg(Index, "You have Gained 401 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 86 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 404)
Call PlayerMsg(Index, "You have Gained 404 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 87 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 408)
Call PlayerMsg(Index, "You have Gained 408 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 88 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 411)
Call PlayerMsg(Index, "You have Gained 411 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 89 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 416)
Call PlayerMsg(Index, "You have Gained 416 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 90 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 419)
Call PlayerMsg(Index, "You have Gained 419 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 91 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 423)
Call PlayerMsg(Index, "You have Gained 423 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 92 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 426)
Call PlayerMsg(Index, "You have Gained 426 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 93 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 429)
Call PlayerMsg(Index, "You have Gained 429 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 94 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 433)
Call PlayerMsg(Index, "You have Gained 433 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 95 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 438)
Call PlayerMsg(Index, "You have Gained 438 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 96 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 443)
Call PlayerMsg(Index, "You have Gained 443 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 97 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 447)
Call PlayerMsg(Index, "You have Gained 447 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 98 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 451)
Call PlayerMsg(Index, "You have Gained 451 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 99 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 456)
Call PlayerMsg(Index, "You have Gained 456 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 100 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 460)
Call PlayerMsg(Index, "You have Gained 460 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 101 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 463)
Call PlayerMsg(Index, "You have Gained 463 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 102 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 467)
Call PlayerMsg(Index, "You have Gained 467 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 103 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 471)
Call PlayerMsg(Index, "You have Gained 471 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 104 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 476)
Call PlayerMsg(Index, "You have Gained 476 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 105 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 480)
Call PlayerMsg(Index, "You have Gained 480 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 106 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 485)
Call PlayerMsg(Index, "You have Gained 485 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 107 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 489)
Call PlayerMsg(Index, "You have Gained 489 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 108 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 493)
Call PlayerMsg(Index, "You have Gained 493 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 109 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 498)
Call PlayerMsg(Index, "You have Gained 498 Experience in your Polearms Level !", BrightGreen)
ElseIf GetPlayerPolesLevel(Index) = 110 Then
Call SetPlayerPolesExp(Index, GetPlayerPolesExp(Index) + 504)
Call PlayerMsg(Index, "You have Gained 504 Experience in your Polearms Level !", BrightGreen)
End If
End Sub

Sub CheckPlayerLargeBladesLevelUp(ByVal Index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long

    c = 0

    If GetPlayerLargeBladesExp(Index) >= GetPlayerNextLargeBladesLevel(Index) Then
        If GetPlayerLargeBladesLevel(Index) < MAX_LEVEL Then
            
                Do Until GetPlayerLargeBladesExp(Index) < GetPlayerNextLargeBladesLevel(Index)
                    DoEvents

                    If GetPlayerLargeBladesLevel(Index) < MAX_LEVEL Then
                        If GetPlayerLargeBladesExp(Index) >= GetPlayerNextLargeBladesLevel(Index) Then
                            d = GetPlayerLargeBladesExp(Index) - GetPlayerNextLargeBladesLevel(Index)
                            Call SetPlayerLargeBladesLevel(Index, GetPlayerLargeBladesLevel(Index) + 1)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerLargeBladesExp(Index, d)
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & c & " Large Blades levels!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " has gained a Large Blades level!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        End If

        If GetPlayerLargeBladesLevel(Index) = MAX_LEVEL Then
            Call SetPlayerLargeBladesExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendStats(Index)
End Sub
