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
Dim MapNum As Long, NpcNum As Long

' Drop the SP
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
        N = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    Else
        N = 0
    End If

    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Attacker), "ATTACKNPC" & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim$(Npc(NpcNum).Name)
    MapNpc(MapNum, MapNpcNum).LastAttack = GetTickCount

    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then

        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "You killed a " & Name, BrightRed, 0)
        If GetPlayerAlignment(Attacker) < 9989 Then
            Call SetPlayerAlignment(Attacker, GetPlayerAlignment(Attacker) + 10)
            Call BattleMsg(Attacker, "You Gain 10 Alignment Points !", BrightGreen, 0)
        End If
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
            Exp = Npc(NpcNum).Exp + (Npc(NpcNum).Exp * Val(Add))
        Else
            Exp = Npc(NpcNum).Exp
        End If

        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If

        ' Check if in party, if so divide up the exp
        If Player(Attacker).InParty = NO Then
            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "You cant gain anymore experience!", BrightBlue, 0)
            Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "You have gained " & Exp & " experience.", BrightBlue, 0)
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
            N = Int(Rnd * Npc(NpcNum).ItemNPC(i).Chance) + 1

            If N = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y)
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
            If Trim$(Npc(NpcNum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, "A " & Trim$(Npc(NpcNum).Name) & " : " & Trim$(Npc(NpcNum).AttackSay) & "", SayColor)
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
    Dim MaxiNum As Long
    Dim LSANum As Long
    Dim spellnum As Long

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
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
            Else
                Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
            End If
            Call UpdateGrid(OldMap, oldx, oldy, GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim))

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
Dim MapNum As Long, NpcNum As Long
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
    NpcNum = MapNpc(MapNum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then

            ' Check if at same coordinates
            X = DirToX(GetPlayerX(Attacker), GetPlayerDir(Attacker))
            y = DirToY(GetPlayerY(Attacker), GetPlayerDir(Attacker))

            If (MapNpc(MapNum, MapNpcNum).y = y) And (MapNpc(MapNum, MapNpcNum).X = X) Then
                If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                    CanAttackNpc = True
                Else

                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                       MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & Npc(NpcNum).SpawnSecs
                    Else
                       Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                    End If

                    If Npc(NpcNum).Speech <> 0 Then
                        Call SendDataTo(Attacker, "STARTSPEECH" & SEP_CHAR & Npc(NpcNum).Speech & SEP_CHAR & 0 & SEP_CHAR & NpcNum & SEP_CHAR & END_CHAR)
                    End If
                End If
            End If
        End If
    End If
End If
End Function

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
Dim AttackSpeed As Long
Dim Dir As Long

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
    NpcNum = MapNpc(MapNum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
   ' Make sure they are on the same map
If IsPlaying(Attacker) Then
If NpcNum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
CanAttackNpcWithArrow = True
            Call SetPlayerAlignment(Attacker, GetPlayerAlignment(Attacker) + 10)
            Call BattleMsg(Attacker, "You Gain 10 Alignment Points !", BrightGreen, 0)
Else

                If Trim$(Npc(NpcNum).AttackSay) <> "" Then
                    Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & " : " & Trim$(Npc(NpcNum).AttackSay), Green)
                End If
                
                If Npc(NpcNum).Speech <> 0 Then
                    For Dir = 0 To 3
                        If DirToX(GetPlayerX(Attacker), Dir) = MapNpc(MapNum, MapNpcNum).X And DirToY(GetPlayerY(Attacker), Dir) = MapNpc(MapNum, MapNpcNum).y Then
                            Call SendDataTo(Attacker, "STARTSPEECH" & SEP_CHAR & Npc(NpcNum).Speech & SEP_CHAR & 0 & SEP_CHAR & NpcNum & SEP_CHAR & END_CHAR)
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
Dim MapNum As Long, NpcNum As Long
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
    NpcNum = MapNpc(MapNum, MapNpcNum).num

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
        If NpcNum > 0 Then
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
Dim MapNum As Long, NpcNum As Long
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
Dim MapNum As Long, NpcNum As Long
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
    NpcNum = MapNpc(MapNum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > Player(Attacker).Pet.AttackTimer + 1000 Then
            If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
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
                                If Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
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
Dim C As Long

    C = 0

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
                            i = Int(GetPlayerSPEED(Index) / 10)

                            If i < 1 Then i = 1
                            If i > 3 Then i = 3
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + i)
                            Call SetPlayerExp(Index, d)
                            C = C + 1
                        End If
                    End If

                Loop

                If C > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " has gained " & C & " levels!", 6)
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

Function GetNpcHPRegen(ByVal NpcNum As Long)
Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcHPRegen = 0
        Exit Function
    End If
    i = Int(Npc(NpcNum).DEF / 3)

    If i < 1 Then i = 1
    GetNpcHPRegen = i
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
    GetNpcMaxSP = Npc(NpcNum).Speed * 2
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
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
        Else
            Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
        End If
        Call UpdateGrid(OldMap, oldx, oldy, GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim))

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
Dim MapNum As Long, NpcNum As Long
Dim Dir As Long, X As Long, y As Long
Dim Packet As String

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the pet attacking
    Call SendDataToMap(Player(Attacker).Pet.Map, "PETATTACKNPC" & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    MapNum = Player(Attacker).Pet.Map
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim$(Npc(NpcNum).Name)
    MapNpc(MapNum, MapNpcNum).LastAttack = GetTickCount
    For Dir = 0 To 3

        If MapNpc(MapNum, NpcNum).X = DirToX(Player(Attacker).Pet.X, Dir) And MapNpc(MapNum, NpcNum).y = DirToY(Player(Attacker).Pet.y, Dir) Then
            Packet = "CHANGEPETDIR" & SEP_CHAR & Dir & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR
            Call SendDataToMap(Player(Attacker).Pet.Map, Packet)
        End If
    Next

    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        For i = 1 To MAX_NPC_DROPS

            ' Drop the goods if they get it
            N = Int(Rnd * Npc(NpcNum).ItemNPC(i).Chance) + 1

            If N = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).y)
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
        If Map(GetPlayerMap(Index)).Owner = "" Then
        If GetPlayerName(Index) = Map(GetPlayerMap(Index)).Owner Then
            Call PlayerMsg(Index, "You already own this house!", BrightRed)
            Exit Sub
        Else
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 = 0 Then
                Call SendDataTo(Index, "housebuy" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Else
                If Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(Index, "This house will cost you " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 & " " & Trim(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Name) & "!", Yellow)
                Else
                    Call PlayerMsg(Index, "This house will cost you a " & Trim(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Name) & "!", Yellow)
                End If
                Call SendDataTo(Index, "housebuy" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            End If
        End If
            Else
    Call PlayerMsg(Index, "This house is not for sale!", BrightRed)
    Exit Sub
    End If
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
Dim NpcNum As Long
Dim i As Long, X As Long, y As Long
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
