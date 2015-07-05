Attribute VB_Name = "modGameLogic"
Option Explicit


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

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)

  Dim Name As String
  Dim Exp As Long
  Dim n As Long
  Dim I As Long
  Dim q As Integer
  Dim x As Long
  Dim MapNum As Long
  Dim NpcNum As Long

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
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), PacketID.Attack & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim(Npc(NpcNum).Name)

    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        Player(Attacker).targetnpc = 0

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
                Call BattleMsg(Attacker, Trim(GetVar(App.Path & "Lang.ini", "Lang", "CantGain")), BrightBlue, 0)
             Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "You gained " & Exp & " experience.", BrightBlue, 0)
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
                                Call BattleMsg(Player(Attacker).Party.Member(x), Trim(GetVar(App.Path & "Lang.ini", "Lang", "CantGain")), BrightBlue, 0)
                             Else
                                Call SetPlayerExp(Player(Attacker).Party.Member(x), GetPlayerExp(Player(Attacker).Party.Member(x)) + Exp)
                                Call BattleMsg(Player(Attacker).Party.Member(x), "You Gained " & Exp & " party experience.", BrightBlue, 0)
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
                                Call BattleMsg(Player(Attacker).Party.Member(x), Trim(GetVar(App.Path & "Lang.ini", "Lang", "CantGain")), BrightBlue, 0)
                             Else
                                Call SetPlayerExp(Player(Attacker).Party.Member(x), GetPlayerExp(n) + Exp)
                                Call BattleMsg(Player(Attacker).Party.Member(x), "You Gained " & Exp & " party experience.", BrightBlue, 0)
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
        Call SendDataToMap(MapNum, PacketID.NPCDead & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)

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
        Player(Attacker).targetnpc = MapNpcNum

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

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)

  Dim Exp As Long
  Dim n As Long

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
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), PacketID.Attack & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR)

    If Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA Then
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
                        Call BattleMsg(Victim, Trim(GetVar(App.Path & "Lang.ini", "Lang", "LostNo")), BrightRed, 1)
                        Call BattleMsg(Attacker, Trim(GetVar(App.Path & "Lang.ini", "Lang", "RecievedNo")), BrightBlue, 0)
                     Else
                        Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                        Call BattleMsg(Victim, Trim(GetVar(App.Path & "Lang.ini", "Lang", "YouLost")) & " " & Exp & " experience.", BrightRed, 1)
                        Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                        Call BattleMsg(Attacker, Trim(GetVar(App.Path & "Lang.ini", "Lang", "YouGot")) & " " & Exp & " experience for killing " & GetPlayerName(Victim) & ".", BrightBlue, 0)
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

                    If Not Trim(GetVar(App.Path & "Lang.ini", "Lang", "DeemedPk")) = "" Then
                        Call GlobalMsg(GetPlayerName(Attacker) & " " & Trim(GetVar(App.Path & "Lang.ini", "Lang", "DeemedPk")), BrightRed)
                    End If

                End If
             Else
                Call SetPlayerPK(Victim, NO)
                Call SendPlayerData(Victim)

                If Not Trim(GetVar(App.Path & "Lang.ini", "Lang", "DeemedPk")) = "" Then
                    Call GlobalMsg(GetPlayerName(Victim) & " " & Trim(GetVar(App.Path & "Lang.ini", "Lang", "PaidPk")), BrightRed)
                End If

            End If
         Else

            ' Player not dead, just do the damage
            Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
            Call SendHP(Victim)

            ' Check for a weapon and say damage
            'Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
            'Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)

        End If
     ElseIf Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then

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
    Call SendDataToMap(GetPlayerMap(Victim), PacketID.Sound & SEP_CHAR & "pain" & SEP_CHAR & Player(Victim).Char(Player(Victim).CharNum).Sex & SEP_CHAR & END_CHAR)

End Sub

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean

  Dim MapNum As Long
  Dim NpcNum As Long
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

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean

  Dim MapNum As Long
  Dim NpcNum As Long
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

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean

  Dim AttackSpeed As Long
  Dim MinPkLvl As Integer

    MinPkLvl = 0 + Val(ReadINI("CONFIG", "MinPkLvl", "Data.ini"))

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
                If Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                    ' Check to make sure the victim isn't an admin

                    If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
                     Else
                        ' Check if map is attackable

                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level

                            If GetPlayerLevel(Attacker) < MinPkLvl Then
                                Call PlayerMsg(Attacker, "You are below level " & MinPkLvl & ", you cannot attack another player yet!", BrightRed)
                             Else

                                If GetPlayerLevel(Victim) < MinPkLvl Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & MinPkLvl & ", you cannot attack this player yet!", BrightRed)
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
                            Call PlayerMsg(Attacker, Trim(GetVar(App.Path & "Lang.ini", "Lang", "SafeZone")), BrightRed)
                        End If

                    End If
                 ElseIf Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                    CanAttackPlayer = True
                End If

            End If

         Case DIR_DOWN

            If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                If Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
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

                                If GetPlayerLevel(Attacker) < MinPkLvl Then
                                    Call PlayerMsg(Attacker, "You are below level " & MinPkLvl & ", you cannot attack another player yet!", BrightRed)
                                 Else

                                    If GetPlayerLevel(Victim) < MinPkLvl Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & MinPkLvl & ", you cannot attack this player yet!", BrightRed)
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
                                Call PlayerMsg(Attacker, Trim(GetVar(App.Path & "Lang.ini", "Lang", "SafeZone")), BrightRed)
                            End If

                        End If
                    End If
                 ElseIf Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                    CanAttackPlayer = True
                End If

            End If

         Case DIR_LEFT

            If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                If Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
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

                                If GetPlayerLevel(Attacker) < MinPkLvl Then
                                    Call PlayerMsg(Attacker, "You are below level " & MinPkLvl & ", you cannot attack another player yet!", BrightRed)
                                 Else

                                    If GetPlayerLevel(Victim) < MinPkLvl Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & MinPkLvl & ", you cannot attack this player yet!", BrightRed)
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
                                Call PlayerMsg(Attacker, Trim(GetVar(App.Path & "Lang.ini", "Lang", "SafeZone")), BrightRed)
                            End If

                        End If
                    End If
                 ElseIf Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                    CanAttackPlayer = True
                End If

            End If

         Case DIR_RIGHT

            If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                If Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
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

                                If GetPlayerLevel(Attacker) < MinPkLvl Then
                                    Call PlayerMsg(Attacker, "You are below level " & MinPkLvl & ", you cannot attack another player yet!", BrightRed)
                                 Else

                                    If GetPlayerLevel(Victim) < MinPkLvl Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & MinPkLvl & ", you cannot attack this player yet!", BrightRed)
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
                                Call PlayerMsg(Attacker, Trim(GetVar(App.Path & "Lang.ini", "Lang", "SafeZone")), BrightRed)
                            End If

                        End If
                    End If
                 ElseIf Map(GetPlayerMap(Victim)).tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                    CanAttackPlayer = True
                End If

            End If
        End Select
    End If

End Function

Function CanAttackPlayerWithArrow(ByVal Attacker As Long, ByVal Victim As Long) As Boolean

    CanAttackPlayerWithArrow = False

    ' Check To make sure that they dont have access

    If GetPlayerAccess(Attacker) > ADMIN_MONITER Then
        Call PlayerMsg(Attacker, "You can't attack for thou art an admin.", BrightBlue)
     Else
        ' Check To make sure the victim isn't an admin

        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
            Call PlayerMsg(Attacker, "You can't attack " & GetPlayerName(Victim) & " for he is an admin!", BrightRed)
         Else
            ' Check If map Is attackable

            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                ' Make sure they are high enough level

                If GetPlayerLevel(Attacker) < 10 Then
                    Call PlayerMsg(Attacker, "Your level is below 10 you can't attack anybody until your level 10 or higher.", BrightRed)
                 Else

                    If GetPlayerLevel(Victim) < 10 Then
                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is lower then level 10 for that you can't attack him.", BrightRed)
                     Else

                        If Trim(GetPlayerGuild(Attacker)) <> "" And GetPlayerGuild(Victim) <> "" Then
                            If Trim(GetPlayerGuild(Attacker)) <> Trim(GetPlayerGuild(Victim)) Then
                                CanAttackPlayerWithArrow = True
                             Else
                                Call PlayerMsg(Attacker, "Is in the same guild as you are for that you can't attack him.", BrightRed)
                            End If

                         Else
                            CanAttackPlayerWithArrow = True
                        End If

                    End If
                End If
             Else
                Call PlayerMsg(Attacker, Trim(GetVar(App.Path & "Lang.ini", "Lang", "SafeZone")), BrightRed)
            End If

        End If
    End If

End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean

  Dim MapNum As Long
  Dim NpcNum As Long

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

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean

  Dim I As Long
  Dim n As Long
  Dim x As Long
  Dim y As Long

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
            n = Map(MapNum).tile(x, y - 1).Type

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
            n = Map(MapNum).tile(x, y + 1).Type

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
            n = Map(MapNum).tile(x - 1, y).Type

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
            n = Map(MapNum).tile(x + 1, y).Type

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

Sub canon(ByVal index As Long)

    Call SendDataTo(index, PacketID.OnCanon & SEP_CHAR & END_CHAR)

End Sub

Function CanPlayerBlockHit(ByVal index As Long) As Boolean

  Dim I As Long
  Dim n As Long
  Dim ShieldSlot As Long

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

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean

  Dim I As Long
  Dim n As Long

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

Sub CastSpell(ByVal index As Long, ByVal SpellSlot As Long)

  Dim SpellNum As Long
  Dim I As Long
  Dim n As Long
  Dim Damage As Long
  Dim Casted As Boolean

    Casted = False
    ' Prevent player from using spells if they have been script locked

    If Player(index).lockedspells = True Then
        Exit Sub
    End If

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

    I = GetSpellReqLevel(SpellNum)

    ' Check if they have enough MP

    If GetPlayerMP(index) < Spell(SpellNum).MPCost Then
        Call BattleMsg(index, Trim(GetVar(App.Path & "Lang.ini", "Lang", "NotEnoughMana")), BrightRed, 0)
        Exit Sub
    End If

    ' Make sure they are the right level

    If I > GetPlayerLevel(index) Then
        Call BattleMsg(index, "You need to be " & I & "to cast this spell.", BrightRed, 0)
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

  Dim x As Long
  Dim y As Long

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

                                If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
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
                                        Call PlayerMsg(index, Trim(GetVar(App.Path & "Lang.ini", "Lang", "NotCast")), BrightRed)
                                    End If

                                End If
                             Else
                                Player(index).TargetType = TARGET_TYPE_PLAYER

                                If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
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
                                                Call BattleMsg(index, "A deadly mix of elements harm the " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & "!", Blue, 0)
                                                Damage = Int(Damage * 1.25)
                                                If Element(Spell(SpellNum).Element).Strong = Npc(MapNpc(GetPlayerMap(index), n).num).Element And Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Weak = Spell(SpellNum).Element Then Damage = Int(Damage * 1.2)
                                            End If

                                            If Element(Spell(SpellNum).Element).Weak = Npc(MapNpc(GetPlayerMap(index), n).num).Element Or Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Strong = Spell(SpellNum).Element Then
                                                Call BattleMsg(index, "The " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & " aborbs much of the elemental damage!", Red, 0)
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
                    Call SendDataToMap(GetPlayerMap(index), PacketID.SpellAnim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & Player(index).Target & SEP_CHAR & Player(index).CastedSpell & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
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

                If GetPlayerHP(n) > 0 And GetPlayerMap(index) = GetPlayerMap(n) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(n) >= 10 And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(n) <= 0 Then
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
                Call PlayerMsg(index, Trim(GetVar(App.Path & "Lang.ini", "Lang", "NotCast")), BrightRed)
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
                                Call BattleMsg(index, "A deadly mix of elements harm the " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & "!", Blue, 0)
                                Damage = Int(Damage * 1.25)
                                If Element(Spell(SpellNum).Element).Strong = Npc(MapNpc(GetPlayerMap(index), n).num).Element And Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Weak = Spell(SpellNum).Element Then Damage = Int(Damage * 1.2)
                            End If

                            If Element(Spell(SpellNum).Element).Weak = Npc(MapNpc(GetPlayerMap(index), n).num).Element Or Element(Npc(MapNpc(GetPlayerMap(index), n).num).Element).Strong = Spell(SpellNum).Element Then
                                Call BattleMsg(index, "The " & Trim(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & " aborbs much of the elemental damage!", Red, 0)
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
        Call SendDataToMap(GetPlayerMap(index), PacketID.SpellAnim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & Player(index).Target & SEP_CHAR & Player(index).CastedSpell & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & END_CHAR)
        Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "magic" & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR)
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

Sub CheckEquippedItems(ByVal index As Long)

  Dim slot As Long
  Dim ItemNum As Long

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

Sub CheckPlayerLevelUp(ByVal index As Long)

  Dim I As Long
  Dim d As Long
  Dim c As Long

    c = 0
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
                            c = c + 1
                        End If

                    End If
                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(index) & " " & Trim(GetVar(App.Path & "Lang.ini", "Lang", "Gained")) & " " & c & " levels!", 6)
                 Else
                    Call GlobalMsg(GetPlayerName(index) & " " & Trim(GetVar(App.Path & "Lang.ini", "Lang", "GainedA")), 6)
                End If

                Call BattleMsg(index, "You have " & GetPlayerPOINTS(index) & " stat points", 9, 0)
            End If

            Call SendDataToMap(GetPlayerMap(index), PacketID.LevelUp & SEP_CHAR & index & SEP_CHAR & END_CHAR)
            Call SendPlayerLevelToAll(index)
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

Sub DoSkill(ByVal index As Integer, ByVal skil As Integer, ByVal skillsheet As Integer)

  Dim sheet As Integer
  Dim I As Integer
  Dim success_chance As Integer
  Dim a As Integer

    For I = 1 To MAX_SKILLS_SHEETS
        If skillsheet <> 0 Then
            sheet = skillsheet
            a = 1
         Else
            sheet = I
        End If

        Select Case Item(skill(skil).itemequiped(sheet)).Type
         Case ITEM_TYPE_WEAPON

            If GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) <> skill(skil).itemequiped(sheet) Then
                Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                GoTo Hell
            End If

         Case ITEM_TYPE_ARMOR

            If GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) <> skill(skil).itemequiped(sheet) Then
                Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                GoTo Hell
            End If

         Case ITEM_TYPE_HELMET

            If GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) <> skill(skil).itemequiped(sheet) Then
                Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                GoTo Hell
            End If

         Case ITEM_TYPE_SHIELD

            If GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) <> skill(skil).itemequiped(sheet) Then
                Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                GoTo Hell
            End If

         Case ITEM_TYPE_LEGS

            If GetPlayerInvItemNum(index, GetPlayerLegsSlot(index)) <> skill(skil).itemequiped(sheet) Then
                Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                GoTo Hell
            End If

         Case ITEM_TYPE_RING

            If GetPlayerInvItemNum(index, GetPlayerRingSlot(index)) <> skill(skil).itemequiped(sheet) Then
                Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                GoTo Hell
            End If

         Case ITEM_TYPE_NECKLACE

            If GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index)) <> skill(skil).itemequiped(sheet) Then
                Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                GoTo Hell
            End If

        End Select

        If Player(index).Char(Player(index).CharNum).SkillLvl(skil) < skill(skil).minlevel(sheet) Then
            Call PlayerMsg(index, "You aren't skilled enough to " & skill(skil).Name & " here. You need to be level " & skill(skil).minlevel(sheet) & ".", 4)
            GoTo Hell
        End If

        If "" & skill(skil).AttemptName <> "" Then Call PlayerMsg(index, skill(skil).AttemptName, White)

        success_chance = skill(skil).base_chance(sheet)

        If Int((Val(Player(index).Char(Player(index).CharNum).SkillLvl(skil)) - Int(skill(skil).minlevel(sheet)) + Int(success_chance) + 1) * Rnd) <= Int(Player(index).Char(Player(index).CharNum).SkillLvl(skil) + 1) - Int(skill(skil).minlevel(sheet)) + 1 Then
            If "" & skill(skil).Succes <> "" Then
                Call PlayerMsg(index, skill(skil).Succes, Green)
            End If

         Else
            If "" & skill(skil).Fail <> "" Then Call PlayerMsg(index, skill(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1).Fail, Red)
            GoTo Hell
        End If

        If skill(skil).ItemTake1num(sheet) <> 0 Then
            If TakeItemPlayer(index, skill(skil).ItemTake1num(sheet), skill(skil).ItemTake1val(sheet)) = False Then
                Call PlayerMsg(index, "You need " & Trim$(Item(skill(skil).ItemGive1num(sheet) + 1).Name) & " to " & Trim$(skill(skil).Action) & " here.", Red)
                GoTo Hell
            End If

        End If

        If skill(skil).ItemTake2num(sheet) <> 0 Then
            If TakeItemPlayer(index, skill(skil).ItemTake2num(sheet), skill(skil).ItemTake2val(sheet)) = False Then
                Call PlayerMsg(index, "You need " & Trim(Item(skill(skil).ItemGive2num(sheet)).Name) & " to " & Trim(skill(skil).Action) & " here.", Red)
                GoTo Hell
            End If

        End If

        If skill(skil).ItemGive1num(sheet) <> 0 Then
            If GiveItemPlayer(index, skill(skil).ItemGive1num(sheet), skill(skil).ItemGive1val(sheet)) = False Then
                Call PlayerMsg(index, "You don't have enough inventory space! You need to make some room in order to " & Trim(skill(skil).Action) & " here.", Red)
                GoTo Hell
            End If

            Call PlayerMsg(index, "You " & Trim$(skill(skil).Action) & " a " & Trim(Item(skill(skil).ItemGive1num(sheet)).Name) & ".", Green)
        End If

        If skill(skil).ItemGive2num(sheet) <> 0 Then
            If GiveItemPlayer(index, skill(skil).ItemGive2num(sheet), skill(skil).ItemGive2val(sheet)) = False Then
                Call PlayerMsg(index, "You don't have enough inventory space! You need to make some room in order to " & Trim(skill(skil).Action) & " here.", Red)
                GoTo Hell
            End If

            Call PlayerMsg(index, "You " & Trim$(skill(skil).Action) & " a " & Trim(Item(skill(skil).ItemGive2num(sheet)).Name) & ".", Green)
        End If

        'Add EXP and send if it's changed

        If skill(skil).ExpGiven(sheet) <> 0 Then
            Player(index).Char(Player(index).CharNum).SkillExp(skil) = Player(index).Char(Player(index).CharNum).SkillExp(skil) + skill(skil).ExpGiven(sheet)
            Call SendUpdatePlayerSkill(index, skil)
        End If

        'experience.ini

        If 0 + ReadINI("EXPERIENCE", "Exp" & Val(Player(index).Char(Player(index).CharNum).SkillLvl(skil) + 1), App.Path & "\experience.ini") <= Player(index).Char(Player(index).CharNum).SkillExp(skil) Then
            Player(index).Char(Player(index).CharNum).SkillExp(skil) = 0
            Player(index).Char(Player(index).CharNum).SkillLvl(skil) = Player(index).Char(Player(index).CharNum).SkillLvl(skil) + 1
            Call SendUpdatePlayerSkill(index, skil)
            Call SendDataToMap(GetPlayerMap(index), PacketID.LevelUp & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        End If

        Exit Sub

        ' Pfft, programmer humor :P - Pickle
Hell:
        If a = 1 Then Exit Sub
    Next I

End Sub

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

Function GetPlayerDamage(ByVal index As Long) As Long

  Dim WeaponSlot As Long
  Dim RingSlot As Long
  Dim NecklaceSlot As Long

    GetPlayerDamage = 0
    ' Check for subscript out of range

    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If

    'GetPlayerDamage in script - TODO LATER - Can't get it to work. :(
    'If Scripting = 1 Then
    '    GetPlayerDamage = MyScript.RunCodeReturn("Scripts\Main.txt", "GetPlayerDamage ", index)
    'Else
    GetPlayerDamage = Int(GetPlayerSTR(index) / 2)
    'End If

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

Function GetPlayerProtection(ByVal index As Long) As Long

  Dim ArmorSlot As Long
  Dim HelmSlot As Long
  Dim ShieldSlot As Long
  Dim LegsSlot As Long

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
                    Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, HelmSlot) & "/" & Trim(Item(GetPlayerInvItemNum(index, HelmSlot)).Data1), Yellow, 0)
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
                    Call BattleMsg(index, "Your " & Trim(Item(GetPlayerInvItemNum(index, LegsSlot)).Name) & " " & Trim(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " is about to break! Dur: " & GetPlayerInvItemDur(index, LegsSlot) & "/" & Trim(Item(GetPlayerInvItemNum(index, LegsSlot)).Data1), Yellow, 0)
                End If

            End If
        End If
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

Function GetPLeader(ByVal index As Byte) As Byte

    GetPLeader = Player(index).Party.Leader

End Function

Function GetPMember(ByVal index As Byte, ByVal Member As Byte) As Byte

    GetPMember = Player(index).Party.Member(Member)

End Function

Function GetPShare(ByVal index As Byte) As Boolean

    GetPShare = Player(index).Party.ShareExp

End Function

Sub GetRidOfTimer(ByVal Name As String)

    Call CTimers.Remove(Name)

End Sub

Function GetSpellReqLevel(ByVal SpellNum As Long)

    GetSpellReqLevel = Spell(SpellNum).LevelReq ' - Int(GetClassMAGI(GetPlayerClass(index)) / 4)

End Function

Function GetTimeLeft(ByVal Name As String) As Long

    On Error GoTo Hell
    GetTimeLeft = CTimers.Item(Name).tmrWait - GetTickCount
    Exit Function
Hell:
    GetTimeLeft = -1

End Function

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long

  Dim I As Long
  Dim n As Long

    n = 0

    For I = 1 To MAX_PLAYERS

        If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
            n = n + 1
        End If

    Next I

    GetTotalMapPlayers = n

End Function

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
        Call SendDataTo(index, PacketID.BankMsg & SEP_CHAR & "Bank full!" & SEP_CHAR & END_CHAR)
    End If

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
        Call PlayerMsg(index, Trim(GetVar(App.Path & "Lang.ini", "Lang", "FullInv")), BrightRed)
    End If

End Sub

Function GiveItemPlayer(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean

  Dim I As Long

    GiveItemPlayer = False

    ' Check for subscript out of range

    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
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
        GiveItemPlayer = True
     Else
        Call PlayerMsg(index, Trim(GetVar(App.Path & "Lang.ini", "Lang", "FullInv")), BrightRed)
    End If

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

Sub JoinGame(ByVal index As Long)

    'On Error GoTo ErrorHandler
  Dim MOTD As String
  Dim f As Long

    ' Check for the dreaded Subscript Out Of Range
    If index < 0 Or index > MAX_PLAYERS Then Call ReportError("modGameLogic", "JoinGame - index RTE9", Err.number, Err.Description)
    ' Set the flag so we know the person is in the game
    Player(index).InGame = True

    ' Send an ok to client to start receiving in game data
    Call SendDataTo(index, PacketID.LoginOK & SEP_CHAR & index & SEP_CHAR & END_CHAR)

    ReDim Player(index).Party.Member(1 To MAX_PARTY_MEMBERS)

    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendEmoticons(index)
    Call SendElements(index)
    Call SendSkills(index)
    Call SendQuests(index)
    Call SendArrows(index)
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
    Call SendGameClockTo(index)
    Call SendNewsTo(index)
    Call DisabledTimeTo(index)
    Call Sendsprite(index, index)
    Call SendActionNames(index)

    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    Call SendPlayerData(index)

    Player(index).HookShotX = 0
    Player(index).HookShotY = 0

    If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "JoinGame " & index
     Else
        MOTD = GetVar("motd.ini", "MOTD", "Msg")

        ' Send a global message that he/she joined

        If GetPlayerAccess(index) <= ADMIN_MONITER Then
            Call GlobalMsg(GetPlayerName(index) & " " & Trim(GetVar(App.Path & "Lang.ini", "Lang", "Joined")) & " " & GAME_NAME & "!", 7)
         Else
            Call GlobalMsg(GetPlayerName(index) & " " & Trim(GetVar(App.Path & "Lang.ini", "Lang", "Joined")) & " " & GAME_NAME & "!", 15)
        End If

        ' Send them welcome

        If Not MOTD + "" = "" Then
            MOTD = STR(MOTD)
        End If

        Call SendDataTo(index, MOTD)

        Call PlayerMsg(index, Trim(GetVar(App.Path & "Lang.ini", "Lang", "Welcome")) & " " & GAME_NAME & "!", 15)

        ' Send motd

        If Trim(MOTD) <> "" Then
            Call PlayerMsg(index, Trim(GetVar(App.Path & "Lang.ini", "Lang", "Motd")) & " " & MOTD, 11)
        End If

        ' Send whos online
        Call SendWhosOnline(index)
    End If

    Call ShowPLR(index)

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(index, PacketID.InGame & SEP_CHAR & year & SEP_CHAR & month & SEP_CHAR & day & SEP_CHAR & weekday & SEP_CHAR & END_CHAR)

    Exit Sub
ErrorHandler:
    Call ReportError("modGameLogic", "JoinGame", Err.number, Err.Description)

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

            If Map(GetPlayerMap(index)).BootMap > 0 Then
                Call SetPlayerX(index, Map(GetPlayerMap(index)).BootX)
                Call SetPlayerY(index, Map(GetPlayerMap(index)).BootY)
                Call SetPlayerMap(index, Map(GetPlayerMap(index)).BootMap)
            End If

            ' Send a global message that he/she left

            If GetPlayerAccess(index) <= 1 Then
                Call GlobalMsg(GetPlayerName(index) & " " & Trim(GetVar(App.Path & "Lang.ini", "Lang", "Left")) & " " & GAME_NAME & "!", 7)
             Else
                Call GlobalMsg(GetPlayerName(index) & " " & Trim(GetVar(App.Path & "Lang.ini", "Lang", "Left")) & " " & GAME_NAME & "!", 15)
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


'ASGARD
Public Sub LoadWordfilter()

  Dim I

    On Error GoTo LoadWordfilter_Error
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

    On Error GoTo 0
    Exit Sub

LoadWordfilter_Error:

    MsgBox "Error loading word filter: " & Err.number & " (" & Err.Description & ") in procedure LoadWordfilter of Module modGameLogic. Check your word filter files."

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
    Call SendDataToMap(GetPlayerMap(Victim), PacketID.NPCAttack & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)

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
    'Call SendDataToMap(GetPlayerMap(Victim), PacketID.ChangeDir & SEP_CHAR & GetPlayerDir(Victim) & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    'End If
    ':: END AUTO TURN ::

    Name = Trim(Npc(MapNpc(MapNum, MapNpcNum).num).Name)

    If Damage >= GetPlayerHP(Victim) Then
        ' Say damage
        'Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)

        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)

        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " was kille by " & Name, BrightRed)

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
                Call BattleMsg(Victim, Trim(GetVar(App.Path & "Lang.ini", "Lang", "LostNo")), BrightRed, 0)
             Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call BattleMsg(Victim, Trim(GetVar(App.Path & "Lang.ini", "Lang", "YouLost")) & " " & Exp & " experience.", BrightRed, 0)
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

    Call SendDataTo(Victim, PacketID.BlitNPCDmg & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
    Call SendDataToMap(GetPlayerMap(Victim), PacketID.Sound & SEP_CHAR & "pain" & SEP_CHAR & Player(Victim).Char(Player(Victim).CharNum).Sex & SEP_CHAR & END_CHAR)

End Sub

Sub NPCDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)

  Dim packet As String

    ' Check for subscript out of range

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(MapNum, MapNpcNum).Dir = Dir
    packet = PacketID.NPCDir & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, packet)

End Sub

Sub NPCMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)

  Dim packet As String

    ' Check for subscript out of range

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    MapNpc(MapNum, MapNpcNum).Dir = Dir

    Select Case Dir
     Case DIR_UP
        MapNpc(MapNum, MapNpcNum).y = MapNpc(MapNum, MapNpcNum).y - 1
        packet = PacketID.NPCMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)

     Case DIR_DOWN
        MapNpc(MapNum, MapNpcNum).y = MapNpc(MapNum, MapNpcNum).y + 1
        packet = PacketID.NPCMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)

     Case DIR_LEFT
        MapNpc(MapNum, MapNpcNum).x = MapNpc(MapNum, MapNpcNum).x - 1
        packet = PacketID.NPCMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)

     Case DIR_RIGHT
        MapNpc(MapNum, MapNpcNum).x = MapNpc(MapNum, MapNpcNum).x + 1
        packet = PacketID.NPCMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)
    End Select

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
                    Call SendIndexWornEquipment(index)
                End If

                MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)

             Case ITEM_TYPE_WEAPON

                If InvNum = GetPlayerWeaponSlot(index) Then
                    Call SetPlayerWeaponSlot(index, 0)
                    Call SendWornEquipment(index)
                    Call SendIndexWornEquipment(index)
                End If

                MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)

             Case ITEM_TYPE_HELMET

                If InvNum = GetPlayerHelmetSlot(index) Then
                    Call SetPlayerHelmetSlot(index, 0)
                    Call SendWornEquipment(index)
                    Call SendIndexWornEquipment(index)
                End If

                MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)

             Case ITEM_TYPE_SHIELD

                If InvNum = GetPlayerShieldSlot(index) Then
                    Call SetPlayerShieldSlot(index, 0)
                    Call SendWornEquipment(index)
                    Call SendIndexWornEquipment(index)
                End If

                MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)

             Case ITEM_TYPE_LEGS

                If InvNum = GetPlayerLegsSlot(index) Then
                    Call SetPlayerLegsSlot(index, 0)
                    Call SendWornEquipment(index)
                    Call SendIndexWornEquipment(index)
                End If

                MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)

             Case ITEM_TYPE_RING

                If InvNum = GetPlayerRingSlot(index) Then
                    Call SetPlayerRingSlot(index, 0)
                    Call SendWornEquipment(index)
                    Call SendIndexWornEquipment(index)
                End If

                MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)

             Case ITEM_TYPE_NECKLACE

                If InvNum = GetPlayerNecklaceSlot(index) Then
                    Call SetPlayerNecklaceSlot(index, 0)
                    Call SendWornEquipment(index)
                    Call SendIndexWornEquipment(index)
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

                'Normally messages for item drops would go here but it's scripted now

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

Sub PlayerMapGetItem(ByVal index As Long)

  Dim I As Long
  Dim n As Long
  Dim MapNum As Long
  Dim Msg As String

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
                        Msg = "You pickup " & MapItem(MapNum, I).Value & " " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
                     Else
                        Call SetPlayerInvItemValue(index, n, 0)
                        Msg = "You pickup " & Trim(Item(GetPlayerInvItemNum(index, n)).Name) & "."
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
                    Call PlayerMsg(index, Msg, Yellow)
                    Exit Sub
                 Else
                    Call PlayerMsg(index, Trim(GetVar(App.Path & "Lang.ini", "Lang", "FullInv")), BrightRed)
                    Exit Sub
                End If

            End If
        End If
    Next I

End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long)

  Dim packet As String
  Dim MapNum As Long
  Dim x As Long
  Dim y As Long
  Dim I As Long
  Dim Moved As Byte
  Dim skil As Long
  Dim sheet As Long
  Dim a As Long

    ' They tried to hack
    'If Moved = NO Then
    'Call HackingAttempt(index, "Position Modification")
    'Exit Sub
    'End If

    ' Check for subscript out of range

    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    ' Check for scrolling to prevent RTE 9

    If GetPlayerX(index) > MAX_MAPX Or GetPlayerY(index) > MAX_MAPY Then
        Call PlayerWarp(index, GetPlayerMap(index), 0, 0)
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)

    Moved = NO

    Select Case Dir
     Case DIR_UP
        ' Check to make sure not outside of boundries

        If GetPlayerY(index) > 0 Then
            ' Check to make sure that the tile is walkable

            If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_ROOFBLOCK Then
                If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_GUILDBLOCK And Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).String1) <> Trim(GetPlayerGuild(index)) Then
                    Exit Sub
                End If

                ' Check to see if the tile is a skill tile

                If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_SKILL Then
                    ' Check to see if the tile is a key and if it is check if its opened

                    If (Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) - 1)

                        packet = PacketID.PlayerMove & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), packet)
                        Moved = YES
                    End If

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

            If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_ROOFBLOCK Then
                If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_GUILDBLOCK And Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).String1 <> GetPlayerGuild(index) Then
                    Exit Sub
                End If

                If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_SKILL Then
                    ' Check to see if the tile is a key and if it is check if its opened

                    If (Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                        Call SetPlayerY(index, GetPlayerY(index) + 1)

                        packet = PacketID.PlayerMove & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), packet)
                        Moved = YES
                    End If

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

            If Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_ROOFBLOCK Then
                If Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_GUILDBLOCK And Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).String1 <> GetPlayerGuild(index) Then
                    Exit Sub
                End If

                'Check to see if the tile is a skill tile

                If Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_SKILL Then
                    ' Check to see if the tile is a key and if it is check if its opened

                    If (Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                        'BARON DEBUG TIME - CAUSE OF RTE 9'S ABOVE ?
                        Call SetPlayerX(index, GetPlayerX(index) - 1)

                        packet = PacketID.PlayerMove & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), packet)
                        Moved = YES
                    End If

                End If
            End If
         Else
            ' Check to see if we can move them to the another map

            If Map(GetPlayerMap(index)).left > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).left, MAX_MAPX, GetPlayerY(index))
                Moved = YES
            End If

        End If

     Case DIR_RIGHT
        ' Check to make sure not outside of boundries

        If GetPlayerX(index) < MAX_MAPX Then
            ' Check to make sure that the tile is walkable

            If Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_ROOFBLOCK Then
                If Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_GUILDBLOCK And Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).String1 <> GetPlayerGuild(index) Then
                    Exit Sub
                End If

                ' Check for skill tile

                If Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_SKILL Then
                    ' Check to see if the tile is a key and if it is check if its opened

                    If (Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                        Call SetPlayerX(index, GetPlayerX(index) + 1)

                        packet = PacketID.PlayerMove & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                        Call SendDataToMapBut(index, GetPlayerMap(index), packet)
                        Moved = YES
                    End If

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

    If GetPlayerX(index) < 0 Or GetPlayerY(index) < 0 Or GetPlayerX(index) > MAX_MAPX Or GetPlayerY(index) > MAX_MAPY Or GetPlayerMap(index) <= 0 Then
        Call HackingAttempt(index, "")
        Exit Sub
    End If

    'healing tiles code

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SendHP(index)
        Call SendMP(index)
        Call PlayerMsg(index, "You feel a sudden rush through your body as you regain strength!", BrightGreen)
    End If

    'Check for kill tile, and if so kill them

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KILL Then
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
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)

            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(index), PacketID.MapKey & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If

        End If
    End If

    If GetPlayerX(index) - 1 >= 0 Then
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)

            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(index), PacketID.MapKey & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If

        End If
    End If

    If GetPlayerY(index) - 1 >= 0 Then
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1

            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(index), PacketID.MapKey & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If

        End If
    End If

    If GetPlayerY(index) + 1 <= MAX_MAPY Then
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1

            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(index), PacketID.MapKey & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "door" & SEP_CHAR & END_CHAR)
            End If

        End If
    End If

    ' Check to see if the tile is a warp tile, and if so warp them

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_WARP Then
        MapNum = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1
        x = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2
        y = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3

        Call PlayerWarp(index, MapNum, x, y)
        Moved = YES
    End If

    ' Check for key trigger open

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KEYOPEN Then
        x = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1
        y = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2

        If Map(GetPlayerMap(index)).tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
            TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

            Call SendDataToMap(GetPlayerMap(index), PacketID.MapKey & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)

            If Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1) = "" Then
                Call MapMsg(GetPlayerMap(index), "A door has been unlocked!", White)
             Else
                Call MapMsg(GetPlayerMap(index), Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1), White)
            End If

            Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
        End If

    End If

    ' Check for shop

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SHOP Then
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1 > 0 Then
            Call SendTrade(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1)
         Else
            Call PlayerMsg(index, "There is no shop here.", BrightRed)
        End If

    End If

    ' Check if player stepped on sprite changing tile

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(index) = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            Call PlayerMsg(index, "You already have this sprite!", BrightRed)
            Exit Sub
         Else

            If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
                Call SendDataTo(index, PacketID.SpriteChange & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
             Else

                If Item(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(index, "This sprite will cost you " & Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3 & " " & Trim(Item(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", Yellow)
                 Else
                    Call PlayerMsg(index, "This sprite will cost you a " & Trim(Item(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", Yellow)
                End If

                Call SendDataTo(index, PacketID.SpriteChange & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            End If

        End If
    End If
    ' Check if player stepped on house buying tile

    ' Check if player stepped on house buying tile

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_HOUSE Then
        If Len(Map(GetPlayerMap(index)).owner) < 2 Then
            If GetPlayerName(index) = Map(GetPlayerMap(index)).owner Then
                Call PlayerMsg(index, "You already own this house!", BrightRed)
                Exit Sub
             Else

                If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 0 Then
                    Call SendDataTo(index, PacketID.HouseBuy & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                 Else

                    If Item(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1).Type = ITEM_TYPE_CURRENCY Then
                        Call PlayerMsg(index, "This house will cost you " & Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2 & " " & Trim(Item(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1).Name) & "!", Yellow)
                     Else
                        Call PlayerMsg(index, "This house will cost you a " & Trim(Item(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1).Name) & "!", Yellow)
                    End If

                    Call SendDataTo(index, PacketID.HouseBuy & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                End If

            End If
         Else
            Call PlayerMsg(index, "This house is not for sale!", BrightRed)
            Exit Sub
        End If

    End If

    ' Check if player stepped on sprite changing tile

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2 > -1 Then
            If GetPlayerClass(index) <> Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
                Call PlayerMsg(index, "You arent the required class!", BrightRed)
                Exit Sub
            End If

        End If

        If GetPlayerClass(index) = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            Call PlayerMsg(index, "You are already this class!", BrightRed)
         Else

            If Player(index).Char(Player(index).CharNum).Sex = 0 Then
                If GetPlayerSprite(index) = Class(GetPlayerClass(index)).MaleSprite Then
                    Call SetPlayerSprite(index, Class(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1).MaleSprite)
                End If

             Else

                If GetPlayerSprite(index) = Class(GetPlayerClass(index)).FemaleSprite Then
                    Call SetPlayerSprite(index, Class(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1).FemaleSprite)
                End If

            End If

            Call SetPlayerSTR(index, (Player(index).Char(Player(index).CharNum).STR - Class(GetPlayerClass(index)).STR))
            Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).DEF - Class(GetPlayerClass(index)).DEF))
            Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).Magi - Class(GetPlayerClass(index)).Magi))
            Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed - Class(GetPlayerClass(index)).Speed))

            Call SetPlayerClass(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1)

            Call SetPlayerSTR(index, (Player(index).Char(Player(index).CharNum).STR + Class(GetPlayerClass(index)).STR))
            Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).DEF + Class(GetPlayerClass(index)).DEF))
            Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).Magi + Class(GetPlayerClass(index)).Magi))
            Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed + Class(GetPlayerClass(index)).Speed))

            Call PlayerMsg(index, "Your new class is a " & Trim(Class(GetPlayerClass(index)).Name) & "!", BrightGreen)

            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
            Call SendDataToMap(GetPlayerMap(index), PacketID.CheckSprite & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
        End If

    End If

    ' Check if player stepped on notice tile

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_NOTICE Then
        If Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1) <> "" Then
            Call PlayerMsg(index, Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1), Black)
        End If

        If Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String2) <> "" Then
            Call PlayerMsg(index, Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String2), Grey)
        End If

        If Not Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String3 = "" Or Not Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String3 = vbNullString Then
            Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String3 & SEP_CHAR & END_CHAR)
        End If

    End If

    'Check if player steppted on minus stat tile

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_LOWER_STAT Then
        If Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1) <> "" Then
            Call PlayerMsg(index, Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1), Black)
        End If

        If Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1) <> 0 Then
            Call SetPlayerHP(index, GetPlayerHP(index) - Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1))
        End If

        If Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2) <> 0 Then
            Call SetPlayerMP(index, GetPlayerMP(index) - Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2))
        End If

        If Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3) <> 0 Then
            Call SetPlayerSP(index, GetPlayerSP(index) - Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3))
        End If

    End If

    ' Check if player stepped on sound tile

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1 & SEP_CHAR & END_CHAR)
    End If

    If Scripting = 1 Then
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & index & "," & Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1
        End If

    End If

    ' Check if player stepped on Bank tile

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_BANK Then
        Call SendDataTo(index, PacketID.OpenBank & SEP_CHAR & END_CHAR)
    End If

    ' Check if player stepped on canon tile

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_CANON Then
        Call canon(index)
     Else
        Call SendDataTo(index, PacketID.CanonOff & SEP_CHAR & END_CHAR)
    End If

    ' Check if player stepped on skill tile

    If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SKILL Then
        skil = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1 + 1

        For I = 1 To MAX_SKILLS_SHEETS

            If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2 <> 0 Then
                sheet = Val(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2)
                a = 1
             Else
                sheet = I
            End If

            Select Case Item(skill(skil).itemequiped(sheet)).Type
             Case ITEM_TYPE_WEAPON

                If GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) <> skill(skil).itemequiped(sheet) Then
                    Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                    GoTo Hell
                End If

             Case ITEM_TYPE_ARMOR

                If GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) <> skill(skil).itemequiped(sheet) Then
                    Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                    GoTo Hell
                End If

             Case ITEM_TYPE_HELMET

                If GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) <> skill(skil).itemequiped(sheet) Then
                    Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                    GoTo Hell
                End If

             Case ITEM_TYPE_SHIELD

                If GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) <> skill(skil).itemequiped(sheet) Then
                    Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                    GoTo Hell
                End If

             Case ITEM_TYPE_LEGS

                If GetPlayerInvItemNum(index, GetPlayerLegsSlot(index)) <> skill(skil).itemequiped(sheet) Then
                    Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                    GoTo Hell
                End If

             Case ITEM_TYPE_RING

                If GetPlayerInvItemNum(index, GetPlayerRingSlot(index)) <> skill(skil).itemequiped(sheet) Then
                    Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                    GoTo Hell
                End If

             Case ITEM_TYPE_NECKLACE

                If GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index)) <> skill(skil).itemequiped(sheet) Then
                    Call PlayerMsg(index, "You need to have a " & Trim(Item(skill(skil).itemequiped(sheet)).Name) & " equiped to " & Trim(skill(skil).Name) & " here.", 4)
                    GoTo Hell
                End If

            End Select

            Exit Sub
Hell:
            If a = 1 Then Exit Sub
        Next I

    End If

End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)

  Dim packet As String
  Dim OldMap As Long
  Dim npcn As Long

    'Error handling
    On Error GoTo WarpErr

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

    If Not OldMap = MapNum Then
        Call SendLeaveMap(index, OldMap)
    End If

    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    'Do we need to spawn NPCs on the new map?

    If PlayersOnMap(MapNum) = NO Then
        SpawnMapNpcs (MapNum)
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "warp" & SEP_CHAR & END_CHAR)

    Player(index).GettingMap = YES
    Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)

    Call SendInventory(index)
    'Call SendWornEquipment(index)
    Call SendIndexWornEquipmentFromMap(index)

    packet = PacketID.ForceHouseClose & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

    If Player(index).pet <> 0 Then

        Do While npcn <= MAX_MAP_NPCS

            If 0 + Map(MapNum).Npc(npcn) = 0 Then
                Call ScriptSpawnNpc(npcn, MapNum, GetPlayerX(index), GetPlayerY(index), Player(index).pet)
                MapNpc(MapNum, npcn).owner = index
                npcn = MAX_MAP_NPCS + 1
             Else
                npcn = npcn + 1
            End If

        Loop
    End If

    'Fixes an OnMapLoad problem

    If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & index & OldMap
    End If

    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & index & " on map " & GetPlayerMap(index) & ".", "logs\ErrorLog.txt")

End Sub

Public Sub RemovePLR()

    frmServer.lvUsers.ListItems.Clear

End Sub

Public Sub RemovePMember(ByVal index As Byte)

  Dim I
  Dim b
  Dim q As Integer

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
            Call PlayerMsg(b, "The party has been disbaned", White)
            Player(b).InParty = False

            For I = 1 To MAX_PARTY_MEMBERS ' clears player's party
                Player(b).Party.Member(I) = 0
            Next I

        End If
    End If

    For I = 1 To MAX_PARTY_MEMBERS

        If Player(index).Party.Member(I) > 0 And Player(index).Party.Member(I) = index Then
            Call SendDataTo(index, PacketID.RemoveMembers & SEP_CHAR & SEP_CHAR & END_CHAR)
        End If

        If Player(index).Party.Member(I) > 0 And Player(index).Party.Member(I) <> index Then
            Call SendDataTo(index, PacketID.UpdateMembers & SEP_CHAR & I & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        End If

    Next I

End Sub

Sub ReportError(ByVal module As String, ByVal subroutine As String, ByVal ErrNum As Integer, ByVal errDesc As String)

    Call MsgBox("Run-time error " & STR(ErrNum) & ": " & errDesc & vbFormFeed & "Module: " & module & "Routine: " & subroutine, vbCritical, "Error!")
    Call DestroyServer

End Sub

Sub ScriptSetAttribute(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)

  Dim packet As String

    With Map(mapper).tile(x, y)
        .Type = Attrib
        .Data1 = Data1
        .Data2 = Data2
        .Data3 = Data3
        .String1 = String1
        .String2 = String2
        .String3 = String3
    End With

    packet = PacketID.TileCheckAttribute & SEP_CHAR & mapper & SEP_CHAR & CStr(x) & SEP_CHAR & CStr(y) & SEP_CHAR

    With Map(mapper).tile(x, y)
        packet = packet & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR
    End With

    Call SendDataToAll(packet & END_CHAR)

End Sub

Sub ScriptSetTile(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal setx As Long, ByVal sety As Long, ByVal tileset As Long, ByVal layer As Long)

  Dim packet As String

    packet = PacketID.TileCheck & SEP_CHAR & mapper & SEP_CHAR & CStr(x) & SEP_CHAR & CStr(y) & SEP_CHAR & CStr(layer) & SEP_CHAR
    Select Case layer

     Case 0
        Map(mapper).tile(x, y).Ground = sety * 14 + setx
        Map(mapper).tile(x, y).GroundSet = tileset
        packet = packet & Map(mapper).tile(x, y).Ground & SEP_CHAR & Map(mapper).tile(x, y).GroundSet

     Case 1
        Map(mapper).tile(x, y).Mask = sety * 14 + setx
        Map(mapper).tile(x, y).MaskSet = tileset
        packet = packet & Map(mapper).tile(x, y).Mask & SEP_CHAR & Map(mapper).tile(x, y).MaskSet

     Case 2
        Map(mapper).tile(x, y).Anim = sety * 14 + setx
        Map(mapper).tile(x, y).AnimSet = tileset
        packet = packet & Map(mapper).tile(x, y).Anim & SEP_CHAR & Map(mapper).tile(x, y).AnimSet

     Case 3
        Map(mapper).tile(x, y).Mask2 = sety * 14 + setx
        Map(mapper).tile(x, y).Mask2Set = tileset
        packet = packet & Map(mapper).tile(x, y).Mask2 & SEP_CHAR & Map(mapper).tile(x, y).Mask2Set

     Case 4
        Map(mapper).tile(x, y).M2Anim = sety * 14 + setx
        Map(mapper).tile(x, y).M2AnimSet = tileset
        packet = packet & Map(mapper).tile(x, y).M2Anim & SEP_CHAR & Map(mapper).tile(x, y).M2AnimSet

     Case 5
        Map(mapper).tile(x, y).Fringe = sety * 14 + setx
        Map(mapper).tile(x, y).FringeSet = tileset
        packet = packet & Map(mapper).tile(x, y).Fringe & SEP_CHAR & Map(mapper).tile(x, y).FringeSet

     Case 6
        Map(mapper).tile(x, y).FAnim = sety * 14 + setx
        Map(mapper).tile(x, y).FAnimSet = tileset
        packet = packet & Map(mapper).tile(x, y).FAnim & SEP_CHAR & Map(mapper).tile(x, y).FAnimSet

     Case 7
        Map(mapper).tile(x, y).Fringe2 = sety * 14 + setx
        Map(mapper).tile(x, y).Fringe2Set = tileset
        packet = packet & Map(mapper).tile(x, y).Fringe2 & SEP_CHAR & Map(mapper).tile(x, y).Fringe2Set

     Case 8
        Map(mapper).tile(x, y).F2Anim = sety * 14 + setx
        Map(mapper).tile(x, y).F2AnimSet = tileset
        packet = packet & Map(mapper).tile(x, y).F2Anim & SEP_CHAR & Map(mapper).tile(x, y).F2AnimSet
    End Select

    Call SaveMap(mapper)
    Call SendDataToAll(packet & END_CHAR)

End Sub

Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, ByVal spawn_x As Long, ByVal spawn_y As Long, ByVal NpcNum As Long)

    '                         NPC_index               map_number          X spawn          y spawn            NPC_number
  Dim packet As String
  Dim I As Long

    ' Check for subscript out of range

    If MapNpcNum < 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    If NpcNum = 0 Then
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        MapNpc(MapNum, MapNpcNum).num = 0
        Map(MapNum).Npc(MapNpcNum) = 0
        MapNpc(MapNum, MapNpcNum).Target = 0
        MapNpc(MapNum, MapNpcNum).HP = 0
        MapNpc(MapNum, MapNpcNum).MP = 0
        MapNpc(MapNum, MapNpcNum).SP = 0
        MapNpc(MapNum, MapNpcNum).Dir = 0
        MapNpc(MapNum, MapNpcNum).x = 0
        MapNpc(MapNum, MapNpcNum).y = 0

        'Packet = PacketID.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(mapnum, MapNpcNum).num & SEP_CHAR & MapNpc(mapnum, MapNpcNum).x & SEP_CHAR & MapNpc(mapnum, MapNpcNum).y & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(mapnum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
        'Call SendDataToMap(mapnum, Packet)
        Call SaveMap(MapNum)
    End If

    'MapNpc(mapnum, MapNpcNum).num = 0
    'MapNpc(mapnum, MapNpcNum).SpawnWait = GetTickCount
    'MapNpc(mapnum, MapNpcNum).HP = 0
    'Call SendDataToMap(mapnum, PacketID.NPCDead & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)

    Map(MapNum).Revision = Map(MapNum).Revision + 1

    MapNpc(MapNum, MapNpcNum).num = NpcNum
    Map(MapNum).Npc(MapNpcNum) = NpcNum

    MapNpc(MapNum, MapNpcNum).Target = 0

    MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxhp(NpcNum)
    MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
    MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)

    MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

    MapNpc(MapNum, MapNpcNum).x = spawn_x
    MapNpc(MapNum, MapNpcNum).y = spawn_y

    packet = PacketID.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, packet)

    Call SaveMap(MapNum)

    For I = 1 To MAX_PLAYERS

        If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
            Call SendDataTo(I, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(I) & SEP_CHAR & Map(GetPlayerMap(I)).Revision & SEP_CHAR & END_CHAR)
        End If

    Next I

End Sub

Sub SendIndexWornEquipment(ByVal index As Long)

  Dim packet As String
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

    packet = PacketID.ItemWorn & SEP_CHAR & index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Sub SendIndexWornEquipmentFromMap(ByVal index As Long)

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

                packet = PacketID.ItemWorn & SEP_CHAR & index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & SEP_CHAR & END_CHAR
                Call SendDataTo(index, packet)
            End If

        End If
    Next I

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

Public Sub SetPShare(ByVal index As Byte, ByVal share As Boolean)

    Player(index).Party.ShareExp = share

End Sub

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

Sub SpawnAllMapNpcs()

  Dim I As Long

    For I = 1 To MAX_MAPS

        If PlayersOnMap(I) = YES Then
            Call SpawnMapNpcs(I)
        End If

    Next I

End Sub

Sub SpawnAllMapsItems()

  Dim I As Long

    For I = 1 To MAX_MAPS
        Call SpawnMapItems(I)
    Next I

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

  Dim packet As String
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

        packet = PacketID.SpawnItem & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)
    End If

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)

  Dim x As Integer
  Dim y As Integer

    ' Check for subscript out of range

    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            ' Check if the tile type is an item or a saved tile incase someone drops something

            If IS_SCROLLING = 1 And Map(MapNum).scrolling = 1 Then
                If x = 20 Then
                    x = 0
                    y = y + 1
                End If

                If y = 15 Then Exit Sub
            End If

            If (Map(MapNum).tile(x, y).Type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically

                If (Item(Map(MapNum).tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(MapNum).tile(x, y).Data1).Stackable = 1) And Map(MapNum).tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).tile(x, y).Data1, 1, MapNum, x, y)
                 Else
                    Call SpawnItem(Map(MapNum).tile(x, y).Data1, Map(MapNum).tile(x, y).Data2, MapNum, x, y)
                End If

            End If
        Next x

    Next y

End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)

  Dim I As Long

    For I = 1 To MAX_MAP_NPCS

        If Map(MapNum).Npc(I) > 0 Then
            Call SpawnNPC(I, MapNum)
        End If

    Next I

End Sub

Sub SpawnNPC(ByVal MapNpcNum As Long, ByVal MapNum As Long)

  Dim packet As String
  Dim NpcNum As Long
  Dim I As Long
  Dim x As Long
  Dim y As Long
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
                Call SendDataToMap(MapNum, PacketID.NPCDead & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
                Exit Sub
            End If

         Else

            If Npc(NpcNum).SpawnTime = 2 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, PacketID.NPCDead & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
                Exit Sub
            End If

        End If

        MapNpc(MapNum, MapNpcNum).num = NpcNum
        MapNpc(MapNum, MapNpcNum).Target = 0

        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxhp(NpcNum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)

        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

        ' Try to find an NPC Spawn tile first

        For x = 1 To MAX_MAPX
            For y = 1 To MAX_MAPY

                If Map(MapNum).tile(x, y).Type = TILE_TYPE_NPC_SPAWN And Map(MapNum).tile(x, y).Data1 = MapNpcNum Then
                    MapNpc(MapNum, MapNpcNum).x = x
                    MapNpc(MapNum, MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If

            Next y
            If Spawned = True Then Exit For
        Next x

        ' We'll try 100 times to randomly place the sprite

        If Not Spawned Then

            For I = 1 To 100
                x = Int(Rnd * MAX_MAPX)
                y = Int(Rnd * MAX_MAPY)

                ' Check if the tile is walkable

                If Map(MapNum).tile(x, y).Type = TILE_TYPE_WALKABLE Then
                    MapNpc(MapNum, MapNpcNum).x = x
                    MapNpc(MapNum, MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If

            Next I
        End If

        ' Didn't spawn, so now we'll just try to find a free tile

        If Not Spawned Then

            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX

                    If Map(MapNum).tile(x, y).Type = TILE_TYPE_WALKABLE Then
                        MapNpc(MapNum, MapNpcNum).x = x
                        MapNpc(MapNum, MapNpcNum).y = y
                        Spawned = True
                    End If

                Next x
            Next y
        End If

        ' If we suceeded in spawning then send it to everyone

        If Spawned Then
            packet = PacketID.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
            Call SendDataToMap(MapNum, packet)
        End If

    End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)

End Sub

Public Function SwearCheck(TextToSay As String) As String

  Dim I As Integer
  Dim j As Integer 'no variants y'hear?
  Dim Asterisk As String

    SwearCheck = TextToSay
    If WordList <= 0 Then Exit Function

    For I = 1 To WordList
        Asterisk = ""

        For j = 1 To Len(Wordfilter(I))
            Asterisk = Asterisk & "*"
        Next j

        SwearCheck = Replace$(SwearCheck, Wordfilter(I), Asterisk, 1, -1, vbTextCompare)
    Next I

End Function

Sub TakeBankItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)

  Dim I As Long
  Dim n As Long
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

Sub TakeItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)

  Dim I As Long
  Dim n As Long
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

Function TakeItemPlayer(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean

  Dim I As Long
  Dim n As Long

    TakeItemPlayer = False

    ' Check for subscript out of range

    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For I = 1 To MAX_INV
        ' Check to see if the player has the item

        If GetPlayerInvItemNum(index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero

                If ItemVal >= GetPlayerInvItemValue(index, I) Then
                    TakeItemPlayer = True
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
                            TakeItemPlayer = True
                         Else
                            ' Check if the item we are taking isn't already equipped

                            If ItemNum <> GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) Then
                                TakeItemPlayer = True
                            End If

                        End If
                     Else
                        TakeItemPlayer = True
                    End If

                 Case ITEM_TYPE_ARMOR

                    If GetPlayerArmorSlot(index) > 0 Then
                        If I = GetPlayerArmorSlot(index) Then
                            Call SetPlayerArmorSlot(index, 0)
                            Call SendWornEquipment(index)
                            TakeItemPlayer = True
                         Else
                            ' Check if the item we are taking isn't already equipped

                            If ItemNum <> GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) Then
                                TakeItemPlayer = True
                            End If

                        End If
                     Else
                        TakeItemPlayer = True
                    End If

                 Case ITEM_TYPE_HELMET

                    If GetPlayerHelmetSlot(index) > 0 Then
                        If I = GetPlayerHelmetSlot(index) Then
                            Call SetPlayerHelmetSlot(index, 0)
                            Call SendWornEquipment(index)
                            TakeItemPlayer = True
                         Else
                            ' Check if the item we are taking isn't already equipped

                            If ItemNum <> GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) Then
                                TakeItemPlayer = True
                            End If

                        End If
                     Else
                        TakeItemPlayer = True
                    End If

                 Case ITEM_TYPE_SHIELD

                    If GetPlayerShieldSlot(index) > 0 Then
                        If I = GetPlayerShieldSlot(index) Then
                            Call SetPlayerShieldSlot(index, 0)
                            Call SendWornEquipment(index)
                            TakeItemPlayer = True
                         Else
                            ' Check if the item we are taking isn't already equipped

                            If ItemNum <> GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) Then
                                TakeItemPlayer = True
                            End If

                        End If
                     Else
                        TakeItemPlayer = True
                    End If

                 Case ITEM_TYPE_LEGS

                    If GetPlayerLegsSlot(index) > 0 Then
                        If I = GetPlayerLegsSlot(index) Then
                            Call SetPlayerLegsSlot(index, 0)
                            Call SendWornEquipment(index)
                            TakeItemPlayer = True
                         Else
                            ' Check if the item we are taking isn't already equipped

                            If ItemNum <> GetPlayerInvItemNum(index, GetPlayerLegsSlot(index)) Then
                                TakeItemPlayer = True
                            End If

                        End If
                     Else
                        TakeItemPlayer = True
                    End If

                 Case ITEM_TYPE_RING

                    If GetPlayerRingSlot(index) > 0 Then
                        If I = GetPlayerRingSlot(index) Then
                            Call SetPlayerRingSlot(index, 0)
                            Call SendWornEquipment(index)
                            TakeItemPlayer = True
                         Else
                            ' Check if the item we are taking isn't already equipped

                            If ItemNum <> GetPlayerInvItemNum(index, GetPlayerRingSlot(index)) Then
                                TakeItemPlayer = True
                            End If

                        End If
                     Else
                        TakeItemPlayer = True
                    End If

                 Case ITEM_TYPE_NECKLACE

                    If GetPlayerNecklaceSlot(index) > 0 Then
                        If I = GetPlayerNecklaceSlot(index) Then
                            Call SetPlayerNecklaceSlot(index, 0)
                            Call SendWornEquipment(index)
                            TakeItemPlayer = True
                         Else
                            ' Check if the item we are taking isn't already equipped

                            If ItemNum <> GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index)) Then
                                TakeItemPlayer = True
                            End If

                        End If
                     Else
                        TakeItemPlayer = True
                    End If

                End Select

                n = Item(GetPlayerInvItemNum(index, I)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away

                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) And (n <> ITEM_TYPE_LEGS) And (n <> ITEM_TYPE_RING) And (n <> ITEM_TYPE_NECKLACE) Then
                    TakeItemPlayer = True
                End If

            End If

            If TakeItemPlayer = True Then
                Call SetPlayerInvItemNum(index, I, 0)
                Call SetPlayerInvItemValue(index, I, 0)
                Call SetPlayerInvItemDur(index, I, 0)

                ' Send the inventory update
                Call SendInventoryUpdate(index, I)
                Exit Function
            End If

        End If
    Next I

End Function

Function TotalOnlinePlayers() As Long

    On Error GoTo ErrorHandler
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
    ReportError "modGameLogic.bas", "TotalOnlinePlayers", Err.number, Err.Description

End Function

Public Sub UpdateParty(ByVal index As Byte)

    Player(index).Party = Player(Player(index).Party.Leader).Party

End Sub


'Function TotalOnlinePlayers() As Long
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

