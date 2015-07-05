Attribute VB_Name = "modGameLogic"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.
Option Explicit

Private Const GameLoopTime As Long = 10
Private tmrCheckSockets As Long
Private tmrChatLogsTime As Long
Private tmrCheckSpawnMapItems As Long
Private tmrPlayerSaveTime As Long
Private PlayerTimerTime As Long
Private tmrGameAI As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ServerLoop()
On Error GoTo ErrHandler

Dim LoopStartTime As Long
Dim Elapsed As Long
Dim i As Long

    ServerOn = YES
    
    Do While ServerOn
    
    'Thanks to Spodi's vbGORE
    'Make sure that the system's clock didn't reset
    '(check the sub for more details)
    ValidateTime
    
    LoopStartTime = GetTickCount()
    
    If tmrGameAI < GetTickCount Then
        Call GameAI
        tmrGameAI = GetTickCount + 500
    End If
    
    'No more need for this!
    'ServerLogic
    
    If GiveHPTimer < GetTickCount Then

        For i = 1 To MAX_PLAYERS

            If IsPlaying(i) Then
                If GetPlayerHP(i) < GetPlayerMaxHP(i) And GetPlayerHP(i) >= 0 Then
                    Call SetPlayerHP(i, GetPlayerHP(i) + GetPlayerHPRegen(i))
                    Call SendHP(i)
                End If

                If GetPlayerMP(i) < GetPlayerMaxMP(i) And GetPlayerMP(i) >= 0 Then
                    Call SetPlayerMP(i, GetPlayerMP(i) + GetPlayerMPRegen(i))
                    Call SendMP(i)
                End If

                If GetPlayerSP(i) < GetPlayerMaxSP(i) And GetPlayerSP(i) >= 0 Then
                    Call SetPlayerSP(i, GetPlayerSP(i) + GetPlayerSPRegen(i))
                    Call SendSP(i)
                End If
            End If

            DoEvents
        Next

        GiveHPTimer = GetTickCount + 10000
    End If
    
    If tmrCheckSockets < GetTickCount Then

        ' Check for disconnections, just in case
        For i = 1 To MAX_PLAYERS

            If frmServer.Socket(i).State > 7 Then
                Call CloseSocket(i)
            End If

        Next i
        tmrCheckSockets = GetTickCount + 300000
    End If
    
    ' Since the chat logs and check spawn map items have the same time, we put them together
    If tmrChatLogsTime < GetTickCount Then
        LogChats
        CheckSpawnMapItems
        tmrChatLogsTime = GetTickCount + 1000
    End If
    
    'If tmrCheckSpawnMapItems < GetTickCount Then
    '    CheckSpawnMapItems
    '    tmrCheckSpawnMapItems = GetTickCount + 1000
    'End If
    
    If tmrPlayerSave Then
        If tmrPlayerSaveTime < GetTickCount Then
            PlayerSaveTimer2
            tmrPlayerSaveTime = GetTickCount + 60000
        End If
    End If
    
    If PlayerTimer Then
        If PlayerTimerTime < GetTickCount Then
            PlayerSaveTimer2
            PlayerTimerTime = GetTickCount + 5000
        End If
    End If
    
    If tmrPlayerSave Then
        If tmrPlayerSaveTime < GetTickCount Then
            PlayerSaveTimer
            tmrPlayerSaveTime = GetTickCount + 60000
        End If
    End If
    
    DoEvents
    
    'Check if we have enough time to sleep
    'Thanks to Spodi's vbGORE for the code
    Elapsed = GetTickCount - LoopStartTime
        
    If Elapsed < GameLoopTime Then
        If Elapsed >= 0 Then    'Make sure nothing weird happens, causing for a huge sleep time
            Sleep Int(GameLoopTime - Elapsed)
        End If
    End If
    
    Loop
    
    Call DestroyServer
    Exit Sub
    
ErrHandler:
    Call AddLog("There was an error in ServerLoop()!", "errorlist.txt")
End Sub

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
    Call SendDataToMap(GetPlayerMap(Attacker), ATTACKNPC_CHAR & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & END_CHAR)
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim$(Npc(NpcNum).Name)
    MapNpc(MapNum, MapNpcNum).LastAttack = GetTickCount

    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then

        ' Check for a weapon and say damage
        Call BattleMsg(Attacker, "You killed a " & Name, BrightRed, 0)
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

        If add > 0 Then
            If add < 100 Then
                If add < 10 Then
                    add = 0 & ".0" & Right$(add, 2)
                Else
                    add = 0 & "." & Right$(add, 2)
                End If

            Else
                add = Mid$(add, 1, 1) & "." & Right$(add, 2)
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
                Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
            End If

        Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, NPCDEAD_CHAR & SEP_CHAR & MapNpcNum & END_CHAR)

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

        Call TakeFromGrid(MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)

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
            If Trim$(Npc(NpcNum).AttackSay) <> vbNullString Then
                Call PlayerMsg(Attacker, "A " & Trim$(Npc(NpcNum).Name) & ": " & Trim$(Npc(NpcNum).AttackSay) & vbNullString, SayColor)
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

    'Call SendDataToMap(MapNum, npchp_CHAR & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & END_CHAR)
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub AttackPlayer(ByVal Attacker As Long, _
   ByVal Victim As Long, _
   ByVal Damage As Long)
    Dim Exp As Long
    Dim N As Long
    Dim OldMap, oldx, oldy As Long

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
    Call SendDataToMap(GetPlayerMap(Attacker), ATTACKPLAYER_CHAR & SEP_CHAR & Attacker & SEP_CHAR & Victim & END_CHAR)

    If Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA Then
        If Damage >= GetPlayerHP(Victim) Then

            ' Set HP to nothing
            Call SetPlayerHP(Victim, 0)

            ' Check for a weapon and say damage
            Call BattleMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " damage.", White, 0)
            Call BattleMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " damage.", BrightRed, 1)

            ' Player is dead
            Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
            'Call SendDataToMap(GetPlayerMap(Victim), SOUND_CHAR & SEP_CHAR & "Dead" & END_CHAR)
            Call SendSound(Victim, DEAD_SOUND, SDTM)

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

    ' Reset timer for attacking
    Player(Attacker).AttackTimer = GetTickCount
    'Call SendDataToMap(GetPlayerMap(Victim), SOUND_CHAR & SEP_CHAR & "Pain" & END_CHAR)
    Call SendSound(Victim, PAIN_SOUND, SDTM)
End Sub

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim MapNum As Long, NpcNum As Long
    Dim AttackSpeed As Long
    Dim X As Long
    Dim Y As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 0
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
            X = DirToX(GetPlayerX(Attacker), GetPlayerDir(Attacker))
            Y = DirToY(GetPlayerY(Attacker), GetPlayerDir(Attacker))

            If (MapNpc(MapNum, MapNpcNum).Y = Y) And (MapNpc(MapNum, MapNpcNum).X = X) Then
                If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    CanAttackNpc = True
                Else

                    If Trim$(Npc(NpcNum).AttackSay) <> vbNullString Then
                        Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & ": " & Trim$(Npc(NpcNum).AttackSay), Green)
                    End If

                    If Npc(NpcNum).Speech <> 0 Then
                        Call SendDataTo(Attacker, STARTSPEECH_CHAR & SEP_CHAR & Npc(NpcNum).Speech & SEP_CHAR & 0 & SEP_CHAR & NpcNum & END_CHAR)
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
    If IsPlaying(Attacker) Then
        If NpcNum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
            If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                CanAttackNpcWithArrow = True
            Else

                If Trim$(Npc(NpcNum).AttackSay) <> vbNullString Then
                    Call PlayerMsg(Attacker, Trim$(Npc(NpcNum).Name) & ": " & Trim$(Npc(NpcNum).AttackSay), Green)
                End If

                If Npc(NpcNum).Speech <> 0 Then

                    For Dir = 0 To 3

                        If DirToX(GetPlayerX(Attacker), Dir) = MapNpc(MapNum, MapNpcNum).X And DirToY(GetPlayerY(Attacker), Dir) = MapNpc(MapNum, MapNpcNum).Y Then
                            Call SendDataTo(Attacker, STARTSPEECH_CHAR & SEP_CHAR & Npc(NpcNum).Speech & SEP_CHAR & 0 & SEP_CHAR & NpcNum & END_CHAR)
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
    Dim Y As Long

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
        Y = DirToY(GetPlayerY(Attacker), GetPlayerDir(Attacker))

        If (GetPlayerY(Victim) = Y) And (GetPlayerX(Victim) = X) Then
            If Map(GetPlayerMap(Victim)).Tile(X, Y).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then

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

                                    If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                        If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
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

            ElseIf Map(GetPlayerMap(Victim)).Tile(X, Y).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
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

                                If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                    If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
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
    Dim Y As Long

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
            Y = DirToY(MapNpc(MapNum, MapNpcNum).Y, MapNpc(MapNum, MapNpcNum).Dir)

            ' Check if at same coordinates
            If (Player(Index).Pet.Y = Y) And (Player(Index).Pet.X = X) Then
                CanNpcAttackPet = True
            End If
        End If
    End If

End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long, NpcNum As Long
    Dim X As Long
    Dim Y As Long

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
            Y = DirToY(MapNpc(MapNum, MapNpcNum).Y, MapNpc(MapNum, MapNpcNum).Dir)

            ' Check if at same coordinates
            If (GetPlayerY(Index) = Y) And (GetPlayerX(Index) = X) Then
                CanNpcAttackPlayer = True
            End If
        End If
    End If

End Function

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim X As Long, Y As Long

    CanNpcMove = False

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Function
    X = DirToX(MapNpc(MapNum, MapNpcNum).X, Dir)
    Y = DirToY(MapNpc(MapNum, MapNpcNum).Y, Dir)

    If Not IsValid(X, Y) Then Exit Function
    If Grid(MapNum).Loc(X, Y).Blocked = True Then Exit Function
    If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_WALKABLE And Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_ITEM Then Exit Function
    CanNpcMove = True
End Function

Function CanPetAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim MapNum As Long, NpcNum As Long
    Dim X As Long
    Dim Y As Long
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
                    Y = DirToY(Player(Attacker).Pet.Y, Dir)

                    If (MapNpc(MapNum, MapNpcNum).Y = Y) And (MapNpc(MapNum, MapNpcNum).X = X) Then
                        CanPetAttackNpc = True
                    End If

                Next

            End If
        End If
    End If

End Function

Function CanPetMove(ByVal PetNum As Long, ByVal Dir) As Boolean
    Dim X As Long, Y As Long
    Dim i As Long, Packet As String

    CanPetMove = False

    If PetNum <= 0 Or PetNum > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Function
    X = DirToX(Player(PetNum).Pet.X, Dir)
    Y = DirToY(Player(PetNum).Pet.Y, Dir)

    If Not IsValid(X, Y) Then
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
                'Packet = PETDATA_CHAR & SEP_CHAR
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

    If Grid(Player(PetNum).Pet.Map).Loc(X, Y).Blocked = True Then Exit Function
    CanPetMove = True
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim i As Long, N As Long, ShieldSlot As Long

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
    Dim SpellNum As Long, i As Long, N As Long, Damage As Long
    Dim Casted As Boolean
    Dim X As Long, Y As Long
    Dim Packet As String

    Casted = False
    Call SendPlayerXY(Index)

    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    SpellNum = GetPlayerSpell(Index, SpellSlot)

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then
        Call BattleMsg(Index, "You do not have this spell!", BrightRed, 0)
        Exit Sub
    End If

    i = GetSpellReqLevel(SpellNum)

    ' Check if they have enough MP
    If GetPlayerMP(Index) < Spell(SpellNum).MPCost Then
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
    If Spell(SpellNum).Type = SPELL_TYPE_PET Then
        Player(Index).Pet.Alive = YES
        Player(Index).Pet.Sprite = Spell(SpellNum).Data1
        Player(Index).Pet.Dir = DIR_UP
        Player(Index).Pet.Map = GetPlayerMap(Index)
        Player(Index).Pet.MapToGo = 0
        Player(Index).Pet.X = GetPlayerX(Index) + Int(Rnd * 3 - 1)

        If Player(Index).Pet.X < 0 Or Player(Index).Pet.X > MAX_MAPX Then Player(Index).Pet.X = GetPlayerX(Index)
        Player(Index).Pet.XToGo = -1
        Player(Index).Pet.Y = GetPlayerY(Index) + Int(Rnd * 3 - 1)

        If Player(Index).Pet.Y < 0 Or Player(Index).Pet.Y > MAX_MAPY Then Player(Index).Pet.Y = GetPlayerY(Index)
        Player(Index).Pet.YToGo = -1
        Player(Index).Pet.Level = Spell(SpellNum).Range
        Player(Index).Pet.HP = Player(Index).Pet.Level * 5
        Call AddToGrid(Player(Index).Pet.Map, Player(Index).Pet.X, Player(Index).Pet.Y)
        Packet = PETDATA_CHAR & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.X & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Y & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR

        ' Excuse the messy code, I'm rushing
        Call PlayerMsg(Index, "You summon a beast", White)
        Call SendDataToMap(GetPlayerMap(Index), Packet)
        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
        Call SendMP(Index)
        Casted = True
        Exit Sub
    End If

    If Spell(SpellNum).AE = 1 Then

        For Y = GetPlayerY(Index) - Spell(SpellNum).Range To GetPlayerY(Index) + Spell(SpellNum).Range
            For X = GetPlayerX(Index) - Spell(SpellNum).Range To GetPlayerX(Index) + Spell(SpellNum).Range
                N = -1

                If IsValid(X, Y) Then

                    For i = 1 To MAX_PLAYERS

                        If IsPlaying(i) = True Then
                            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                                If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                                    If i = Index Then
                                        If Spell(SpellNum).Type = SPELL_TYPE_ADDHP Or Spell(SpellNum).Type = SPELL_TYPE_ADDMP Or Spell(SpellNum).Type = SPELL_TYPE_ADDSP Then
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
                            If MapNpc(GetPlayerMap(Index), i).X = X And MapNpc(GetPlayerMap(Index), i).Y = Y Then
                                Player(Index).Target = i
                                Player(Index).TargetType = TARGET_TYPE_NPC
                                N = Player(Index).Target
                            End If
                        End If

                    Next

                    If N < 0 Then
                        Player(Index).Target = MakeLoc(X, Y)
                        Player(Index).TargetType = TARGET_TYPE_LOCATION
                        N = MakeLoc(X, Y)
                    End If

                    Casted = False

                    If N > 0 Then
                        If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
                            If IsPlaying(N) Then
                                If N <> Index Then
                                    Player(Index).TargetType = TARGET_TYPE_PLAYER

                                    If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then

                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Select Case Spell(SpellNum).Type

                                            Case SPELL_TYPE_SUBHP
                                                Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(N) + (Rnd * 5) - 2

                                                If Damage > 0 Then
                                                    Call AttackPlayer(Index, N, Damage)
                                                Else
                                                    Call BattleMsg(Index, "The spell was to weak to hurt " & GetPlayerName(N) & "!", BrightRed, 0)
                                                End If

                                            Case SPELL_TYPE_SUBMP
                                                Call SetPlayerMP(N, GetPlayerMP(N) - Spell(SpellNum).Data1)
                                                Call SendMP(N)

                                            Case SPELL_TYPE_SUBSP
                                                Call SetPlayerSP(N, GetPlayerSP(N) - Spell(SpellNum).Data1)
                                                Call SendSP(N)
                                        End Select

                                        Casted = True
                                    Else

                                        If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then

                                            Select Case Spell(SpellNum).Type

                                                Case SPELL_TYPE_ADDHP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerHP(N, GetPlayerHP(N) + Spell(SpellNum).Data1)
                                                    Call SendHP(N)

                                                Case SPELL_TYPE_ADDMP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerMP(N) + Spell(SpellNum).Data1)
                                                    Call SendMP(N)

                                                Case SPELL_TYPE_ADDSP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerSP(N) + Spell(SpellNum).Data1)
                                                    Call SendMP(N)
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

                                        If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then

                                            Select Case Spell(SpellNum).Type

                                                Case SPELL_TYPE_ADDHP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerHP(N, GetPlayerHP(N) + Spell(SpellNum).Data1)
                                                    Call SendHP(N)

                                                Case SPELL_TYPE_ADDMP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerMP(N) + Spell(SpellNum).Data1)
                                                    Call SendMP(N)

                                                Case SPELL_TYPE_ADDSP

                                                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                    Call SetPlayerMP(N, GetPlayerSP(N) + Spell(SpellNum).Data1)
                                                    Call SendMP(N)
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
                                    If Spell(SpellNum).Type >= SPELL_TYPE_SUBHP And Spell(SpellNum).Type <= SPELL_TYPE_SUBSP Then

                                        'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                                        Select Case Spell(SpellNum).Type

                                            Case SPELL_TYPE_SUBHP
                                                Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(Index), N).num).DEF / 2) + (Rnd * 5) - 2

                                                If Damage > 0 Then
                                                    Call AttackNpc(Index, N, Damage)
                                                Else
                                                    Call BattleMsg(Index, "The spell was to weak to hurt " & Trim$(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & "!", BrightRed, 0)
                                                End If

                                            Case SPELL_TYPE_SUBMP
                                                MapNpc(GetPlayerMap(Index), N).MP = MapNpc(GetPlayerMap(Index), N).MP - Spell(SpellNum).Data1

                                            Case SPELL_TYPE_SUBSP
                                                MapNpc(GetPlayerMap(Index), N).SP = MapNpc(GetPlayerMap(Index), N).SP - Spell(SpellNum).Data1
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
                                    Call BattleMsg(Index, "Could not cast spell!", BrightRed, 0)
                                End If

                            Else
                                Player(Index).TargetType = TARGET_TYPE_LOCATION
                                Casted = True
                            End If
                        End If
                    End If

                    If Casted = True Then
                        Call SendDataToMap(GetPlayerMap(Index), SPELLANIM_CHAR & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & END_CHAR)

                        'Call SendDataToMap(GetPlayerMap(index), sound_CHAR & SEP_CHAR & "magic" & Spell(SpellNum).Sound & END_CHAR)
                    End If
                End If

            Next
        Next

        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
        Call SendMP(Index)
    Else
        N = Player(Index).Target

        If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
            If IsPlaying(N) Then
                If GetPlayerName(N) <> GetPlayerName(Index) Then
                    If CInt(Sqr((GetPlayerX(Index) - GetPlayerX(N)) ^ 2 + ((GetPlayerY(Index) - GetPlayerY(N)) ^ 2))) > Spell(SpellNum).Range Then
                        Call BattleMsg(Index, "You are too far away to hit the target.", BrightRed, 0)
                        Exit Sub
                    End If
                End If

                Player(Index).TargetType = TARGET_TYPE_PLAYER

                If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then

                    'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                    Select Case Spell(SpellNum).Type

                        Case SPELL_TYPE_SUBHP
                            Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - GetPlayerProtection(N) + (Rnd * 5) - 2

                            If Damage > 0 Then
                                Call AttackPlayer(Index, N, Damage)
                            Else
                                Call BattleMsg(Index, "The spell was to weak to hurt " & GetPlayerName(N) & "!", BrightRed, 0)
                            End If

                        Case SPELL_TYPE_SUBMP
                            Call SetPlayerMP(N, GetPlayerMP(N) - Spell(SpellNum).Data1)
                            Call SendMP(N)

                        Case SPELL_TYPE_SUBSP
                            Call SetPlayerSP(N, GetPlayerSP(N) - Spell(SpellNum).Data1)
                            Call SendSP(N)
                    End Select

                    ' Take away the mana points
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
                    Call SendMP(Index)
                    Casted = True
                Else

                    If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(SpellNum).Type >= SPELL_TYPE_ADDHP And Spell(SpellNum).Type <= SPELL_TYPE_ADDSP Then

                        Select Case Spell(SpellNum).Type

                            Case SPELL_TYPE_ADDHP

                                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerHP(N, GetPlayerHP(N) + Spell(SpellNum).Data1)
                                Call SendHP(N)

                            Case SPELL_TYPE_ADDMP

                                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(N, GetPlayerMP(N) + Spell(SpellNum).Data1)
                                Call SendMP(N)

                            Case SPELL_TYPE_ADDSP

                                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(N, GetPlayerSP(N) + Spell(SpellNum).Data1)
                                Call SendMP(N)
                        End Select

                        ' Take away the mana points
                        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
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

            If CInt(Sqr((GetPlayerX(Index) - MapNpc(GetPlayerMap(Index), N).X) ^ 2 + ((GetPlayerY(Index) - MapNpc(GetPlayerMap(Index), N).Y) ^ 2))) > Spell(SpellNum).Range Then
                Call BattleMsg(Index, "You are too far away to hit the target.", BrightRed, 0)
                Exit Sub
            End If

            Player(Index).TargetType = TARGET_TYPE_NPC

            If Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(MapNpc(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then

                'Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                Select Case Spell(SpellNum).Type

                    Case SPELL_TYPE_ADDHP
                        MapNpc(GetPlayerMap(Index), N).HP = MapNpc(GetPlayerMap(Index), N).HP + Spell(SpellNum).Data1

                    Case SPELL_TYPE_SUBHP
                        Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(SpellNum).Data1) - Int(Npc(MapNpc(GetPlayerMap(Index), N).num).DEF / 2 + (Rnd * 5) - 2)

                        If Damage > 0 Then
                            Call AttackNpc(Index, N, Damage)
                        Else
                            Call BattleMsg(Index, "The spell was to weak to hurt " & Trim$(Npc(MapNpc(GetPlayerMap(Index), N).num).Name) & "!", BrightRed, 0)
                        End If

                    Case SPELL_TYPE_ADDMP
                        MapNpc(GetPlayerMap(Index), N).MP = MapNpc(GetPlayerMap(Index), N).MP + Spell(SpellNum).Data1

                    Case SPELL_TYPE_SUBMP
                        MapNpc(GetPlayerMap(Index), N).MP = MapNpc(GetPlayerMap(Index), N).MP - Spell(SpellNum).Data1

                    Case SPELL_TYPE_ADDSP
                        MapNpc(GetPlayerMap(Index), N).SP = MapNpc(GetPlayerMap(Index), N).SP + Spell(SpellNum).Data1

                    Case SPELL_TYPE_SUBSP
                        MapNpc(GetPlayerMap(Index), N).SP = MapNpc(GetPlayerMap(Index), N).SP - Spell(SpellNum).Data1
                End Select

                ' Take away the mana points
                Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(SpellNum).MPCost)
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
        Call SendDataToMap(GetPlayerMap(Index), SPELLANIM_CHAR & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & Player(Index).CastedSpell & END_CHAR)

        'If Spell(SpellNum).sound > 0 Then Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Magic" & Spell(SpellNum).sound & END_CHAR)
        If Spell(SpellNum).sound > 0 Then Call SendSound(Index, MAGIC_SOUND, SDTM, Spell(SpellNum).sound)
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
                            'Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & "CongratulationsNewLevelAchieved" & END_CHAR)
                            Call SendSound(Index, NEWLEVEL_SOUND, SDT)
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

            Call SendDataToMap(GetPlayerMap(Index), LEVELUP_CHAR & SEP_CHAR & Index & END_CHAR)
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

Public Function DirToY(ByVal Y As Long, _
   ByVal Dir As Byte) As Long
    DirToY = Y

    If Dir = DIR_LEFT Or Dir = DIR_RIGHT Then Exit Function

    ' UP = 0, DOWN = 1
    ' 0 * 2 = 0, 0 - 1 = -1
    ' 1 * 2 = 2, 2 - 1 = 1
    DirToY = Y + ((Dir * 2) - 1)
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

    If HPRegenOn = YES Then

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

    If MPRegenOn = YES Then

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

        If GetPlayerInvItemDur(Index, ArmorSlot) > 0 Then
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

        If GetPlayerInvItemDur(Index, HelmSlot) > 0 Then
            Call SetPlayerInvItemDur(Index, HelmSlot, GetPlayerInvItemDur(Index, HelmSlot) - 1)

            If GetPlayerInvItemDur(Index, HelmSlot) <= 0 Then
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

        If GetPlayerInvItemDur(Index, ShieldSlot) > 0 Then
            Call SetPlayerInvItemDur(Index, ShieldSlot, GetPlayerInvItemDur(Index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(Index, ShieldSlot) <= 0 Then
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

    If SPRegenOn = YES Then

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

Function GetSpellReqLevel(ByVal SpellNum As Long)
    GetSpellReqLevel = Spell(SpellNum).LevelReq ' - Int(GetClassMAGI(GetPlayerClass(index)) / 4)
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

        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
        End If

        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
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
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    HasSpell = False

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Public Function IsValid(ByVal X As Long, _
   ByVal Y As Long) As Boolean
    IsValid = True

    If X < 0 Or X > MAX_MAPX Or Y < 0 Or Y > MAX_MAPY Then IsValid = False
End Function

Sub JoinGame(ByVal Index As Long)
    Dim MOTD As String

    ' Set the flag so we know the person is in the game
    Player(Index).InGame = "YES"
    Call SpecialPutVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "GENERAL", "InGame", Player(Index).InGame)
    
    ' Send an ok to client to start receiving in game data
    Call SendDataTo(Index, LOGINOK_CHAR & SEP_CHAR & Index & END_CHAR)
    'Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & "LoggingIntoServer" & END_CHAR)
    Call SendSound(Index, LOGINTOSERVER_SOUND, SDT)

    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendEmoticons(Index)
    Call SendSpeech(Index)
    Call SendArrows(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendStats(Index)
    Call SendWeatherTo(Index)
    Call SendTimeTo(Index)
    Call SendOnlineList
    Call SendFriendListTo(Index)
    Call SendFriendListToNeeded(GetPlayerName(Index))

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

        'Call SendDataToAllBut(Index, SOUND_CHAR & SEP_CHAR & "ANewPlayerHasJoined" & END_CHAR)
        Call SendSound(Index, PLAYERJOINED_SOUND, SDTAB)

        ' Send them welcome
        Call PlayerMsg(Index, "Welcome to " & GAME_NAME & "!", 15)

        ' Send motd
        If Trim$(MOTD) <> vbNullString Then
            Call PlayerMsg(Index, "MOTD: " & MOTD, 11)
        End If
    End If

    ' Send whos online
    Call SendWhosOnline(Index)
    'Call ShowPLR(Index)

    ' Send the flag so they know they can start doing stuff
    Call SendDataTo(Index, INGAME_CHAR & END_CHAR)
End Sub

Sub LeftGame(ByVal Index As Long)
    Dim N As Long
    Dim i As Long

    If Player(Index).InGame = "YES" Then
        Player(Index).InGame = "NO"
        Call SpecialPutVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "GENERAL", "InGame", Player(Index).InGame)
        'Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & "LoggingOutOfServer" & END_CHAR)
        Call SendSound(Index, LOGOUTOFSERVER_SOUND, SDT)
        'Call SendDataToAllBut(Index, SOUND_CHAR & SEP_CHAR & "APlayerHasLeft" & END_CHAR)
        Call SendSound(Index, PLAYERHASLEFT_SOUND, SDTAB)

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) = 0 Then
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
        
        If Player(Index).Pet.Alive = YES Then
            Call TakeFromGrid(GetPlayerMap(Index), Player(Index).Pet.X, Player(Index).Pet.Y)
        End If
        
        Call TakeFromGrid(GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

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
        Call AddLog(GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".", "serverlog.txt")
        'Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " has disconnected from " & GAME_NAME & ".", True)
        Call SendLeftGame(Index)
        Call CloseSocket(Index)
        'Call RemovePLR

        'For N = 1 To MAX_PLAYERS
        '    Call ShowPLR(N)
        'Next

    End If

    Call SendFriendListToNeeded(GetPlayerName(Index))
    Call ClearPlayer(Index)
    Call SendOnlineList
End Sub

' I want to start using the loc system. Instead of two variables...
' (x and y), you can store both as a "loc" and extract them back
Public Function MakeLoc(ByVal X As Long, _
   ByVal Y As Long) As Long
    MakeLoc = (Y * MAX_MAPX) + X
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
    Call SendDataToMap(Player(Victim).Pet.Map, NPCATTACKPET_CHAR & SEP_CHAR & MapNpcNum & SEP_CHAR & Victim & END_CHAR)
    MapNum = Player(Victim).Pet.Map
    Name = Trim$(Npc(MapNpc(MapNum, MapNpcNum).num).Name)

    If Damage >= Player(Victim).Pet.HP Then
        Call BattleMsg(Victim, "Your pet died!", Red, 1)
        Player(Victim).Pet.Alive = NO
        Call TakeFromGrid(Player(Victim).Pet.Map, Player(Victim).Pet.X, Player(Victim).Pet.Y)
        MapNpc(MapNum, MapNpcNum).Target = 0
        Packet = PETDATA_CHAR & SEP_CHAR
        Packet = Packet & Victim & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.X & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Y & SEP_CHAR
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
        Packet = PETHP_CHAR & SEP_CHAR & Player(Victim).Pet.Level * 5 & SEP_CHAR & Player(Victim).Pet.HP & END_CHAR
        Call SendDataTo(Victim, Packet)
    End If

    'Call SendDataTo(Victim, BLITNPCDMGPET_CHAR & SEP_CHAR & Damage & END_CHAR)
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
    Call SendDataToMap(GetPlayerMap(Victim), NPCATTACK_CHAR & SEP_CHAR & MapNpcNum & SEP_CHAR & Victim & END_CHAR)
    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(MapNpc(MapNum, MapNpcNum).num).Name)

    If Damage >= GetPlayerHP(Victim) Then

        ' Say damage
        Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)

        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by a " & Name & ".", BrightRed)
        'Call SendDataToMap(GetPlayerMap(Victim), SOUND_CHAR & SEP_CHAR & "Dead" & END_CHAR)
        Call SendSound(Victim, DEAD_SOUND, SDTM)

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

    Call SendDataTo(Victim, BLITNPCDMG_CHAR & SEP_CHAR & Damage & END_CHAR)
    'Call SendDataToMap(GetPlayerMap(Victim), SOUND_CHAR & SEP_CHAR & "Pain" & END_CHAR)
    Call SendSound(Victim, PAIN_SOUND, SDTM)
End Sub

Sub NpcDir(ByVal MapNum As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Dir As Long)
    Dim Packet As String

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    Packet = NPCDIR_CHAR & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub NpcMove(ByVal MapNum As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Dir As Long, _
   ByVal Movement As Long)
    Dim Packet As String
    Dim X As Long
    Dim Y As Long

    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub
    MapNpc(MapNum, MapNpcNum).Dir = Dir
    X = DirToX(MapNpc(MapNum, MapNpcNum).X, Dir)
    Y = DirToY(MapNpc(MapNum, MapNpcNum).Y, Dir)
    Call UpdateGrid(MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y, MapNum, X, Y)
    MapNpc(MapNum, MapNpcNum).Y = Y
    MapNpc(MapNum, MapNpcNum).X = X
    Packet = NPCMOVE_CHAR & SEP_CHAR & MapNpcNum & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub PetAttackNpc(ByVal Attacker As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Damage As Long)
    Dim Name As String
    Dim N As Long, i As Long
    Dim MapNum As Long, NpcNum As Long
    Dim Dir As Long, X As Long, Y As Long
    Dim Packet As String

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the pet attacking
    Call SendDataToMap(Player(Attacker).Pet.Map, PETATTACKNPC_CHAR & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & END_CHAR)
    MapNum = Player(Attacker).Pet.Map
    NpcNum = MapNpc(MapNum, MapNpcNum).num
    Name = Trim$(Npc(NpcNum).Name)
    MapNpc(MapNum, MapNpcNum).LastAttack = GetTickCount

    For Dir = 0 To 3

        If MapNpc(MapNum, NpcNum).X = DirToX(Player(Attacker).Pet.X, Dir) And MapNpc(MapNum, NpcNum).Y = DirToY(Player(Attacker).Pet.Y, Dir) Then
            Packet = CHANGEPETDIR_CHAR & SEP_CHAR & Dir & SEP_CHAR & Attacker & END_CHAR
            Call SendDataToMap(Player(Attacker).Pet.Map, Packet)
        End If

    Next

    If Damage >= MapNpc(MapNum, MapNpcNum).HP Then

        For i = 1 To MAX_NPC_DROPS

            ' Drop the goods if they get it
            N = Int(Rnd * Npc(NpcNum).ItemNPC(i).Chance) + 1

            If N = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)
            End If

        Next

        Call BattleMsg(Attacker, "Your pet killed a " & Name & ".", Red, 1)

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum, MapNpcNum).num = 0
        MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum, MapNpcNum).HP = 0
        Call SendDataToMap(MapNum, NPCDEAD_CHAR & SEP_CHAR & MapNpcNum & END_CHAR)
        Call TakeFromGrid(MapNum, MapNpc(MapNum, MapNpcNum).X, MapNpc(MapNum, MapNpcNum).Y)

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

    'Call SendDataToMap(MapNum, npchp_CHAR & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & END_CHAR)
    ' Reset attack timer
    Player(Attacker).Pet.AttackTimer = GetTickCount
End Sub

Sub PetMove(ByVal PetNum As Long, _
   ByVal Dir As Long, _
   ByVal Movement As Long)
    Dim Packet As String
    Dim X As Long
    Dim Y As Long
    Dim i As Long

    If GetPlayerMap(PetNum) <= 0 Or GetPlayerMap(PetNum) > MAX_MAPS Or PetNum <= 0 Or PetNum > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub
    Player(PetNum).Pet.Dir = Dir
    X = DirToX(Player(PetNum).Pet.X, Dir)
    Y = DirToY(Player(PetNum).Pet.Y, Dir)

    If IsValid(X, Y) Then
        If Grid(Player(PetNum).Pet.Map).Loc(X, Y).Blocked = True Then
            Packet = CHANGEPETDIR_CHAR & SEP_CHAR & Dir & SEP_CHAR & PetNum & END_CHAR
            Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
            Exit Sub
        End If

        Call UpdateGrid(Player(PetNum).Pet.Map, Player(PetNum).Pet.X, Player(PetNum).Pet.Y, Player(PetNum).Pet.Map, X, Y)
        Player(PetNum).Pet.Y = Y
        Player(PetNum).Pet.X = X
        Packet = PETMOVE_CHAR & SEP_CHAR & PetNum & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
        Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
    Else
        i = Player(PetNum).Pet.Map

        If Dir = DIR_UP Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Up
            Player(PetNum).Pet.Y = MAX_MAPY
        End If

        If Dir = DIR_DOWN Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Down
            Player(PetNum).Pet.Y = 0
        End If

        If Dir = DIR_LEFT Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Left
            Player(PetNum).Pet.X = MAX_MAPX
        End If

        If Dir = DIR_RIGHT Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Right
            Player(PetNum).Pet.X = 0
        End If

        Packet = PETDATA_CHAR & SEP_CHAR
        Packet = Packet & PetNum & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Map & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.X & SEP_CHAR
        Packet = Packet & Player(PetNum).Pet.Y & SEP_CHAR
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

            MapItem(GetPlayerMap(Index), i).num = GetPlayerInvItemNum(Index, InvNum)
            MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
            MapItem(GetPlayerMap(Index), i).Y = GetPlayerY(Index)

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
            If (MapItem(MapNum, i).X = GetPlayerX(Index)) And (MapItem(MapNum, i).Y = GetPlayerY(Index)) Then

                ' Find open slot
                N = FindOpenInvSlot(Index, MapItem(MapNum, i).num)

                ' Open slot available?
                If N <> 0 Then

                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(Index, N, MapItem(MapNum, i).num)

                    If Item(GetPlayerInvItemNum(Index, N)).Type = ITEM_TYPE_CURRENCY Then
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
                    MapItem(MapNum, i).Y = 0
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
    Dim Y As Long
    Dim oldx As Long
    Dim oldy As Long
    Dim OldMap As Long
    Dim Moved As Byte

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
    Y = DirToY(GetPlayerY(Index), Dir)
    Call TakeFromGrid(GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

    ' Move the player's pet out of the way if we need to
    If Player(Index).Pet.Alive = YES Then
        If Player(Index).Pet.Map = GetPlayerMap(Index) And Player(Index).Pet.X = X And Player(Index).Pet.Y = Y Then
            If Grid(GetPlayerMap(Index)).Loc(DirToX(X, Dir), DirToY(Y, Dir)).Blocked = False Then
                Call UpdateGrid(Player(Index).Pet.Map, Player(Index).Pet.X, Player(Index).Pet.Y, Player(Index).Pet.Map, DirToX(X, Dir), DirToY(Y, Dir))
                Player(Index).Pet.Y = DirToY(Y, Dir)
                Player(Index).Pet.X = DirToX(X, Dir)
                Packet = PETMOVE_CHAR & SEP_CHAR & Index & SEP_CHAR & DirToX(X, Dir) & SEP_CHAR & DirToY(Y, Dir) & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
                Call SendDataToMap(Player(Index).Pet.Map, Packet)
            End If
        End If
    End If

    ' Check to make sure not outside of boundries
    If IsValid(X, Y) Then

        ' Check to make sure that the tile is walkable
        If Grid(GetPlayerMap(Index)).Loc(X, Y).Blocked = False Then

            ' Check to see if the tile is a key and if it is check if its opened
            If (Map(GetPlayerMap(Index)).Tile(X, Y).Type <> TILE_TYPE_KEY Or Map(GetPlayerMap(Index)).Tile(X, Y).Type <> TILE_TYPE_DOOR) Or ((Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_DOOR Or Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY) And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES) Then
                Call SetPlayerX(Index, X)
                Call SetPlayerY(Index, Y)
                Packet = PLAYERMOVE_CHAR & SEP_CHAR & Index & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
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
        Call HackingAttempt(Index, vbNullString)
        Exit Sub
    End If

    'healing tiles code
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SendHP(Index)
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SendMP(Index)
        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
        Call SendSP(Index)
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

    If IsValid(X, Y) Then
        If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_DOOR Then
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                Call SendDataToMap(GetPlayerMap(Index), MAPKEY_CHAR & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Key" & END_CHAR)
                Call SendSound(Index, KEY_SOUND, SDTM)
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

    Call AddToGrid(GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

    ' Check for key trigger open
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KEYOPEN Then
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2

        If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
            TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
            Call SendDataToMap(GetPlayerMap(Index), MAPKEY_CHAR & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)

            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", White)
            Else
                Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
            End If

            'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Key" & END_CHAR)
            Call SendSound(Index, KEY_SOUND, SDTM)
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
                Call SendDataTo(Index, SPRITECHANGE_CHAR & SEP_CHAR & 0 & END_CHAR)
            Else

                If Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(Index, "This sprite will cost you " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 & " " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", Yellow)
                Else
                    Call PlayerMsg(Index, "This sprite will cost you a " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", Yellow)
                End If

                Call SendDataTo(Index, SPRITECHANGE_CHAR & SEP_CHAR & 1 & END_CHAR)
            End If
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
            Call SendDataToMap(GetPlayerMap(Index), CHECKSPRITE_CHAR & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
        End If
    End If

    ' Check if player stepped on notice tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_NOTICE Then
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), Black)
        End If

        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2), Grey)
        End If

        Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 & END_CHAR)
    End If

    ' Check if player stepped on sound tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1 & END_CHAR)
    End If

    If SCRIPTING = 1 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        End If
    End If

End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional sound As Boolean = True)
    Dim OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    'If Trim$(Shop(ShopNum).LeaveSay) <> vbNullString Then
    'Call PlayerMsg(Index, Trim$(Shop(ShopNum).Name) & ": " & Trim$(Shop(ShopNum).LeaveSay) & vbNullString, SayColor)
    'End If
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    Call SendLeaveMap(Index, OldMap)
    Call UpdateGrid(OldMap, GetPlayerX(Index), GetPlayerY(Index), MapNum, X, Y)
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    If Player(Index).Pet.Alive = YES Then
        Player(Index).Pet.MapToGo = -1
        Player(Index).Pet.XToGo = -1
        Player(Index).Pet.YToGo = -1
        Player(Index).Pet.Map = MapNum
        Player(Index).Pet.X = X
        Player(Index).Pet.Y = Y
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    Player(Index).GettingMap = YES

    'If sound Then Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Warp" & END_CHAR)
    If sound Then Call SendSound(Index, WARP_SOUND, SDTM)
    Call SendDataTo(Index, CHECKFORMAP_CHAR & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
End Sub

'Public Sub RemovePLR()
'    frmServer.lvUsers.ListItems.Clear
'End Sub

Sub SetUpGrid()
    Dim i As Long
    Dim X As Long
    Dim Y As Long

    'Call ClearGrid

    For i = 1 To MAX_MAPS
        For X = 0 To MAX_MAPX
            For Y = 0 To MAX_MAPY

                If Map(i).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then Grid(i).Loc(X, Y).Blocked = True
            Next
        Next

        For X = 1 To MAX_MAP_NPCS

            If MapNpc(i, X).num > 0 Then
                Grid(i).Loc(MapNpc(i, X).X, MapNpc(i, X).Y).Blocked = True
            End If

        Next
    Next

End Sub

'Public Sub ShowPLR(ByVal Index As Long)
    'Dim ls As ListItem

    'On Error Resume Next

    'If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(Index) = True Then
    '    frmServer.lvUsers.ListItems.Remove Index
    'End If

    'Set ls = frmServer.lvUsers.ListItems.add(Index, , Index)

    'If IsPlaying(Index) = False Then
    '    ls.SubItems(1) = vbNullString
    '    ls.SubItems(2) = vbNullString
    '    ls.SubItems(3) = vbNullString
    '    ls.SubItems(4) = vbNullString
    '    ls.SubItems(5) = vbNullString
    'Else
    '    ls.SubItems(1) = GetPlayerLogin(Index)
    '    ls.SubItems(2) = GetPlayerName(Index)
    '    ls.SubItems(3) = GetPlayerLevel(Index)
    '    ls.SubItems(4) = GetPlayerSprite(Index)
    '    ls.SubItems(5) = GetPlayerAccess(Index)
    'End If

'End Sub

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
   ByVal Y As Long)
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, X, Y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, _
   ByVal ItemNum As Long, _
   ByVal ItemVal As Long, _
   ByVal ItemDur As Long, _
   ByVal MapNum As Long, _
   ByVal X As Long, _
   ByVal Y As Long)
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
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                MapItem(MapNum, i).Dur = ItemDur
            Else
                MapItem(MapNum, i).Dur = 0
            End If

        Else
            MapItem(MapNum, i).Dur = 0
        End If

        MapItem(MapNum, i).X = X
        MapItem(MapNum, i).Y = Y
        Packet = SPAWNITEM_CHAR & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim X As Long
    Dim Y As Long

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
    Dim i As Long, X As Long, Y As Long
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
                Call SendDataToMap(MapNum, NPCDEAD_CHAR & SEP_CHAR & MapNpcNum & END_CHAR)
                Exit Sub
            End If

        Else

            If Npc(NpcNum).SpawnTime = 2 Then
                MapNpc(MapNum, MapNpcNum).num = 0
                MapNpc(MapNum, MapNpcNum).SpawnWait = GetTickCount
                MapNpc(MapNum, MapNpcNum).HP = 0
                Call SendDataToMap(MapNum, NPCDEAD_CHAR & SEP_CHAR & MapNpcNum & END_CHAR)
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
                Y = Int(Rnd * MAX_MAPY)

                ' Check if the tile is walkable
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                    MapNpc(MapNum, MapNpcNum).X = X
                    MapNpc(MapNum, MapNpcNum).Y = Y
                    Spawned = True
                    Exit For
                End If

            Next

            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then

                For Y = 0 To MAX_MAPY
                    For X = 0 To MAX_MAPX

                        If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                            MapNpc(MapNum, MapNpcNum).X = X
                            MapNpc(MapNum, MapNpcNum).Y = Y
                            Spawned = True
                            Exit For
                        End If

                    Next
                Next

            End If

        Else
            MapNpc(MapNum, MapNpcNum).X = Map(MapNum).NpcSpawn(MapNpcNum).X
            MapNpc(MapNum, MapNpcNum).Y = Map(MapNum).NpcSpawn(MapNpcNum).Y
            Spawned = True
        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = SPAWNNPC_CHAR & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & END_CHAR
            Call SendDataToMap(MapNum, Packet)
            Call ResetMapGrid(MapNum)
            'Call AddToGrid(MapNum, MapNpc(MapNum, MapNpcNum).x, MapNpc(MapNum, MapNpcNum).y)
        End If
    End If

    'Call SendDataToMap(MapNum, npchp_CHAR & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & END_CHAR)
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

                N = Item(GetPlayerInvItemNum(Index, i)).Type

                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (N <> ITEM_TYPE_WEAPON) And (N <> ITEM_TYPE_ARMOR) And (N <> ITEM_TYPE_HELMET) And (N <> ITEM_TYPE_SHIELD) Then
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
    Dim Y As Long
    Dim i As Long

    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            Grid(MapNum).Loc(X, Y).Blocked = False

            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then Grid(MapNum).Loc(X, Y).Blocked = True
        Next
    Next

    For i = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum, i).num > 0 Then
            Grid(MapNum).Loc(MapNpc(MapNum, i).X, MapNpc(MapNum, i).Y).Blocked = True
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

Public Sub ValidateTime()

'*****************************************************************
'This will validate that the timer hasn't rolled over
'If the timer does roll over, everything time-based will go out of
'wack, so we just turn off the server and let it reset
'This only happens after the server computer is on for 596.5 hours
'after turning on, then every 1193 hours after that
'*****************************************************************

    'Check if there was a roll-over (current time < last check)
    If GetTickCount < LastGetTickCount Then
        Call AddLog("The system clock has rolled-over, and will mess with anything that is timed based. Server has been forcefully shutdown to reset this.", "errorlist.txt")
        Call DestroyServer
    End If
    
    'Set the last check to now since we just checked it
    LastGetTickCount = GetTickCount
    
End Sub
