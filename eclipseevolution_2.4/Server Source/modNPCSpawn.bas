Attribute VB_Name = "modNPCSpawn"
Option Explicit

Sub AttackAttributeNpc(ByVal MapNpcNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Attacker As Long, ByVal Damage As Long)

  Dim Name As String
  Dim Exp As Long
  Dim n As Long
  Dim i As Long
  Dim q As Integer
  Dim d As Long
  Dim MapNum As Long
  Dim NpcNum As Long

    ' Check for subscript out of range

    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_ATTRIBUTE_NPCS Or Damage < 0 Then
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
    NpcNum = MapAttributeNpc(MapNum, MapNpcNum, X, Y).num
    Name = Trim(Npc(NpcNum).Name)

    If Damage >= MapAttributeNpc(MapNum, MapNpcNum, X, Y).HP Then
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
                Call BattleMsg(Attacker, "Cannot gain anymore experience.", BrightBlue, 0)
             Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                Call BattleMsg(Attacker, "You gained " & Exp & " experience.", BrightBlue, 0)
            End If

         Else
            q = 0
            'The following code will tell us how many party members we have active

            For d = 1 To MAX_PARTY_MEMBERS
                If Player(Attacker).Party.Member(d) > 0 Then q = q + 1
            Next d

            'MsgBox "in party" & q
            If q = 2 Then 'Remember, if they aren't in a party they'll only get one person, so this has to be at least two
                Exp = Exp * 0.75 ' 3/4 experience
                'MsgBox Exp

                For d = 1 To MAX_PARTY_MEMBERS

                    If Player(Attacker).Party.Member(d) > 0 Then
                        If Player(Player(Attacker).Party.Member(d)).Party.ShareExp = True Then
                            If GetPlayerLevel(Player(Attacker).Party.Member(d)) = MAX_LEVEL Then
                                Call SetPlayerExp(Player(Attacker).Party.Member(d), Experience(MAX_LEVEL))
                                Call BattleMsg(Player(Attacker).Party.Member(d), "You cannot gain anymore experience.", BrightBlue, 0)
                             Else
                                Call SetPlayerExp(Player(Attacker).Party.Member(d), GetPlayerExp(Player(Attacker).Party.Member(d)) + Exp)
                                Call BattleMsg(Player(Attacker).Party.Member(d), "You gained " & Exp & " experience.", BrightBlue, 0)
                            End If

                        End If
                    End If
                Next d

             Else 'if there are 3 or more party members..
                Exp = Exp * 0.5  ' 1/2 experience

                For d = 1 To MAX_PARTY_MEMBERS

                    If Player(Attacker).Party.Member(d) > 0 Then
                        If Player(Player(Attacker).Party.Member(d)).Party.ShareExp = True Then
                            If GetPlayerLevel(Player(Attacker).Party.Member(d)) = MAX_LEVEL Then
                                Call SetPlayerExp(Player(Attacker).Party.Member(d), Experience(MAX_LEVEL))
                                Call BattleMsg(Player(Attacker).Party.Member(d), "You cannot gain anymore experience.", BrightBlue, 0)
                             Else
                                Call SetPlayerExp(Player(Attacker).Party.Member(d), GetPlayerExp(n) + Exp)
                                Call BattleMsg(Player(Attacker).Party.Member(d), "You gained " & Exp & " experience.", BrightBlue, 0)
                            End If

                        End If
                    End If
                Next d

            End If
        End If

        For i = 1 To MAX_NPC_DROPS
            ' Drop the goods if they get it
            n = Int(Rnd * Npc(NpcNum).ItemNPC(i).Chance) + 1

            If n = 1 Then
                Call SpawnItem(Npc(NpcNum).ItemNPC(i).ItemNum, Npc(NpcNum).ItemNPC(i).ItemValue, MapNum, MapAttributeNpc(MapNum, MapNpcNum, X, Y).X, MapAttributeNpc(MapNum, MapNpcNum, X, Y).Y)
            End If

        Next i

        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)

        ' Check for level up party member

        If Player(Attacker).InParty = YES Then

            For d = 1 To MAX_PARTY_MEMBERS
                Call CheckPlayerLevelUp(Player(Attacker).Party.Member(d))
            Next d

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
        MapAttributeNpc(MapNum, MapNpcNum, X, Y).HP = MapAttributeNpc(MapNum, MapNpcNum, X, Y).HP - Damage

        ' Check for a weapon and say damage
        'Call BattleMsg(Attacker, "You hit a " & Name & " for " & Damage & " damage.", White, 0)

        If n = 0 Then
            'Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
         Else
            'Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If

        ' Check if we should send a message

        If MapAttributeNpc(MapNum, MapNpcNum, X, Y).Target = 0 And MapAttributeNpc(MapNum, MapNpcNum, X, Y).Target <> Attacker Then
            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                Call PlayerMsg(Attacker, "A " & Trim(Npc(NpcNum).Name) & " : " & Trim(Npc(NpcNum).AttackSay) & "", SayColor)
            End If

        End If

        ' Set the NPC target to the player
        MapAttributeNpc(MapNum, MapNpcNum, X, Y).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm

        If Npc(MapAttributeNpc(MapNum, MapNpcNum, X, Y).num).Behavior = NPC_BEHAVIOR_GUARD Then

            For i = 1 To MAX_ATTRIBUTE_NPCS

                If MapNpc(MapNum, i).num = MapAttributeNpc(MapNum, MapNpcNum, X, Y).num Then
                    MapNpc(MapNum, i).Target = Attacker
                End If

            Next i
        End If

    End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapAttributeNpc(MapNum, MapNpcNum, x, y).HP & SEP_CHAR & GetNpcMaxHP(MapAttributeNpc(MapNum, MapNpcNum, x, y).num) & SEP_CHAR & END_CHAR)

    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount

End Sub

Sub AttackAttributeNpcs(ByVal index As Long)

  Dim i As Long
  Dim X As Long
  Dim Y As Long
  Dim n As Long
  Dim NpcNum As Long
  Dim MapNum As Long
  Dim Damage As Long

    MapNum = GetPlayerMap(index)

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX

            If Map(MapNum).tile(X, Y).Type = TILE_TYPE_NPC_SPAWN Then

                For i = 1 To MAX_ATTRIBUTE_NPCS

                    If i <= Map(MapNum).tile(X, Y).Data2 Then

                        NpcNum = MapAttributeNpc(MapNum, i, X, Y).num

                        ' Can we attack the npc?

                        If CanAttackAttributeNpc(index, i, X, Y) Then
                            ' Get the damage we can do

                            If Not CanPlayerCriticalHit(index) Then
                                Damage = GetPlayerDamage(index) - Int(Npc(NpcNum).DEF / 2)
                                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                             Else
                                n = GetPlayerDamage(index)
                                Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(NpcNum).DEF / 2)
                                Call BattleMsg(index, Trim(GetVar(App.Path & "Lang.ini", "Lang", "Surge")), BrightCyan, 0)

                                'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                            End If

                            If Damage > 0 Then
                                Call AttackAttributeNpc(i, X, Y, index, Damage)
                                'Call SendDataTo(index, PacketID.BlitPlayerDmg & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                             Else
                                Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)

                                'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                                'Call SendDataTo(index, PacketID.BlitPlayerDmg & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                            End If

                            Exit Sub
                        End If

                    End If
                Next i

            End If
        Next X

    Next Y

End Sub

Sub AttributeNpcAttackPlayer(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal Victim As Long, ByVal Damage As Long)

  Dim Name As String
  Dim Exp As Long
  Dim MapNum As Long

    ' Check for subscript out of range

    If index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for subscript out of range

    If MapNpc(GetPlayerMap(Victim), index).num <= 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), PacketID.AttributeNPCAttack & SEP_CHAR & index & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR)

    MapNum = GetPlayerMap(Victim)

    ':: AUTO TURN ::
    'If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "AutoTurn")) = 1 Then
    'If GetPlayerX(Victim) - 1 = MapNpc(MapNum, index).X Then
    'Call SetPlayerDir(Victim, DIR_LEFT)
    'End If
    'If GetPlayerX(Victim) + 1 = MapNpc(MapNum, index).X Then
    'Call SetPlayerDir(Victim, DIR_RIGHT)
    'End If
    'If GetPlayerY(Victim) - 1 = MapNpc(MapNum, index).Y Then
    'Call SetPlayerDir(Victim, DIR_UP)
    'End If
    'If GetPlayerY(Victim) + 1 = MapNpc(MapNum, index).Y Then
    'Call SetPlayerDir(Victim, DIR_DOWN)
    'End If
    'Call SendDataToMap(GetPlayerMap(Victim), PacketID.ChangeDir & SEP_CHAR & GetPlayerDir(Victim) & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    'End If
    ':: END AUTO TURN ::

    Name = Trim(Npc(MapNpc(MapNum, index).num).Name)

    If Damage >= GetPlayerHP(Victim) Then
        ' Say damage
        'Call BattleMsg(Victim, "You were hit for " & Damage & " damage.", BrightRed, 1)

        'Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)

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

            End If

            ' Calculate exp to give attacker
            Exp = GetPlayerExp(Victim) \ 3

            ' Make sure we dont get less then 0

            If Exp < 0 Then Exp = 0

            If Exp = 0 Then
                Call BattleMsg(Victim, Trim(GetVar(App.Path & "Lang.ini", "Lang", "LostNo")), BrightRed, 0)
             Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call BattleMsg(Victim, Trim(GetVar(App.Path & "Lang.ini", "Lang", "YouLost")) & " experience.", BrightRed, 0)
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
        MapNpc(MapNum, index).Target = 0

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

Sub AttributeNPCDir(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal MapNum As Long, ByVal Dir As Long)

  Dim packet As String

    If index > Map(MapNum).tile(X, Y).Data2 Then Exit Sub

    ' Check for subscript out of range

    If MapNum <= 0 Or MapNum > MAX_MAPS Or index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapAttributeNpc(MapNum, index, X, Y).Dir = Dir
    packet = PacketID.AttributeNPCDir & SEP_CHAR & index & SEP_CHAR & Dir & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, packet)

End Sub

Sub AttributeNPCGameAI(ByVal MapNum As Long)

  Dim i As Long
  Dim X As Long
  Dim Y As Long
  Dim n As Long
  Dim d As Long
  Dim Damage As Long
  Dim DistanceX As Long
  Dim DistanceY As Long
  Dim NpcNum As Long
  Dim Target As Long
  Dim DidWalk As Boolean

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX

            If Map(MapNum).tile(X, Y).Type = TILE_TYPE_NPC_SPAWN Then

                For n = 1 To MAX_ATTRIBUTE_NPCS

                    If n <= Map(MapNum).tile(X, Y).Data2 Then
                        NpcNum = MapAttributeNpc(MapNum, n, X, Y).num

                        ' /////////////////////////////////////////
                        ' // This is used for ATTACKING ON SIGHT //
                        ' /////////////////////////////////////////
                        ' If the npc is a attack on sight, search for a player on the map

                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then

                            For d = 1 To MAX_PLAYERS

                                If IsPlaying(d) Then
                                    If GetPlayerMap(d) = Y And MapAttributeNpc(MapNum, n, X, Y).Target = 0 And GetPlayerAccess(d) <= ADMIN_MONITER Then
                                        n = Npc(NpcNum).Range

                                        DistanceX = MapAttributeNpc(MapNum, n, X, Y).X - GetPlayerX(d)
                                        DistanceY = MapAttributeNpc(MapNum, n, X, Y).Y - GetPlayerY(d)

                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1

                                        ' Are they in range?  if so GET'M!

                                        If DistanceX <= n And DistanceY <= n Then
                                            If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                If Trim(Npc(NpcNum).AttackSay) <> "" Then
                                                    Call PlayerMsg(d, "A " & Trim(Npc(NpcNum).Name) & " : " & Trim(Npc(NpcNum).AttackSay) & "", SayColor)
                                                End If

                                                MapAttributeNpc(MapNum, n, X, Y).Target = d
                                            End If

                                        End If
                                    End If
                                End If
                            Next d

                        End If

                        If 0 + MapAttributeNpc(MapNum, n, X, Y).owner <> 0 Then
                            Dim npcn As Long

                            If 0 + GetPlayerMap(MapAttributeNpc(MapNum, n, X, Y).owner) <> MapNum Then

                                Do While npcn <= MAX_MAP_NPCS
                                    
                                    '//!! Old line was:
                                    'If 0 + MapNpc(MapNum, MapNpcNum).num = 0 Then
                                    'But there is no such thing as a MapNpcNum variable in here
                                    'Use npcn instead maybe? :x
                                    If 0 + MapNpc(MapNum, npcn).num = 0 Then
                                        Call ScriptSpawnNpc(npcn, GetPlayerMap(MapAttributeNpc(MapNum, n, X, Y).owner), GetPlayerX(MapAttributeNpc(MapNum, n, X, Y).owner), GetPlayerY(MapAttributeNpc(MapNum, n, X, Y).owner), NpcNum)
                                        MapAttributeNpc(GetPlayerMap(MapAttributeNpc(MapNum, n, X, Y).owner), npcn, GetPlayerX(MapAttributeNpc(MapNum, n, X, Y).owner), GetPlayerY(MapAttributeNpc(MapNum, n, X, Y).owner)).owner = MapAttributeNpc(MapNum, n, X, Y).owner
                                        Call ScriptSpawnNpc(n, MapNum, X, Y, 0)
                                     Else
                                        npcn = npcn + 1
                                    End If

                                Loop
                            End If

                        End If

                        ' /////////////////////////////////////////////
                        ' // This is used for NPC walking/targetting //
                        ' /////////////////////////////////////////////

                        If 0 + MapAttributeNpc(MapNum, n, X, Y).owner <> 0 Then
                            Target = 0 + Player(MapAttributeNpc(MapNum, n, X, Y).owner).Target
                         Else
                            Target = MapAttributeNpc(MapNum, n, X, Y).Target
                        End If

                        ' Check to see if its time for the npc to walk

                        If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            ' Check to see if we are following a player or not

                            If Target > 0 Then
                                ' Check if the player is even playing, if so follow'm

                                If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                    DidWalk = False

                                    i = Int(Rnd * 5)

                                    ' Lets move the npc

                                    Select Case i
                                     Case 0
                                        ' Up

                                        If MapAttributeNpc(MapNum, n, X, Y).Y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_UP) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Down

                                        If MapAttributeNpc(MapNum, n, X, Y).Y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_DOWN) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Left

                                        If MapAttributeNpc(MapNum, n, X, Y).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_LEFT) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Right

                                        If MapAttributeNpc(MapNum, n, X, Y).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_RIGHT) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If

                                     Case 1
                                        ' Right

                                        If MapAttributeNpc(MapNum, n, X, Y).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_RIGHT) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Left

                                        If MapAttributeNpc(MapNum, n, X, Y).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_LEFT) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Down

                                        If MapAttributeNpc(MapNum, n, X, Y).Y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_DOWN) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Up

                                        If MapAttributeNpc(MapNum, n, X, Y).Y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_UP) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If

                                     Case 2
                                        ' Down

                                        If MapAttributeNpc(MapNum, n, X, Y).Y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_DOWN) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Up

                                        If MapAttributeNpc(MapNum, n, X, Y).Y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_UP) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Right

                                        If MapAttributeNpc(MapNum, n, X, Y).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_RIGHT) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Left

                                        If MapAttributeNpc(MapNum, n, X, Y).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_LEFT) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If

                                     Case 3
                                        ' Left

                                        If MapAttributeNpc(MapNum, n, X, Y).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_LEFT) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Right

                                        If MapAttributeNpc(MapNum, n, X, Y).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_RIGHT) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Up

                                        If MapAttributeNpc(MapNum, n, X, Y).Y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_UP) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                        ' Down

                                        If MapAttributeNpc(MapNum, n, X, Y).Y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanAttributeNPCMove(n, X, Y, MapNum, DIR_DOWN) Then
                                                Call AttributeNPCMove(n, X, Y, MapNum, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If

                                        End If
                                    End Select

                                    ' Check if we can't move and if player is behind something and if we can just switch dirs

                                    If Not DidWalk Then
                                        If MapAttributeNpc(MapNum, n, X, Y).X - 1 = GetPlayerX(Target) And MapAttributeNpc(MapNum, n, X, Y).Y = GetPlayerY(Target) Then
                                            If MapAttributeNpc(MapNum, n, X, Y).Dir <> DIR_LEFT Then
                                                Call AttributeNPCDir(n, X, Y, MapNum, DIR_LEFT)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapAttributeNpc(MapNum, n, X, Y).X + 1 = GetPlayerX(Target) And MapAttributeNpc(MapNum, n, X, Y).Y = GetPlayerY(Target) Then
                                            If MapAttributeNpc(MapNum, n, X, Y).Dir <> DIR_RIGHT Then
                                                Call AttributeNPCDir(n, X, Y, MapNum, DIR_RIGHT)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapAttributeNpc(MapNum, n, X, Y).X = GetPlayerX(Target) And MapAttributeNpc(MapNum, n, X, Y).Y - 1 = GetPlayerY(Target) Then
                                            If MapAttributeNpc(MapNum, n, X, Y).Dir <> DIR_UP Then
                                                Call AttributeNPCDir(n, X, Y, MapNum, DIR_UP)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapAttributeNpc(MapNum, n, X, Y).X = GetPlayerX(Target) And MapAttributeNpc(MapNum, n, X, Y).Y + 1 = GetPlayerY(Target) Then
                                            If MapAttributeNpc(MapNum, n, X, Y).Dir <> DIR_DOWN Then
                                                Call AttributeNPCDir(n, X, Y, MapNum, DIR_DOWN)
                                            End If

                                            DidWalk = True
                                        End If

                                        ' We could not move so player must be behind something, walk randomly.

                                        If Not DidWalk Then
                                            i = Int(Rnd * 2)

                                            If i = 1 Then
                                                i = Int(Rnd * 4)

                                                If CanAttributeNPCMove(n, X, Y, MapNum, i) Then
                                                    Call AttributeNPCMove(n, X, Y, MapNum, i, MOVING_WALKING)
                                                End If

                                            End If
                                        End If
                                    End If
                                 Else
                                    MapAttributeNpc(MapNum, n, X, Y).Target = 0
                                End If

                             Else

                                If MapAttributeNpc(MapNum, n, X, Y).owner <> 0 Then
                                    '//!! Old line was:
                                    'If GetPlayerTargetNpc(owner) <> 0 Then
                                    'Variable owner did not exist, not sure what to replace it with so
                                    'I used: MapAttributeNpc(MapNum, n, x, y).owner
                                    If GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner) <> 0 Then
                                        If MapNpc(MapNum, GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).X < X Then
                                            Call NPCMove(MapNum, n, 2, 1)
                                            Exit Sub
                                        End If

                                        If MapNpc(MapNum, GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).X > X Then
                                            Call NPCMove(MapNum, n, 3, 1)
                                            Exit Sub
                                        End If

                                        If MapNpc(MapNum, GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).Y < Y - 1 Then
                                            Call NPCMove(MapNum, n, 0, 1)
                                            Exit Sub
                                        End If

                                        If MapNpc(MapNum, GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).Y > Y + 1 Then
                                            Call NPCMove(MapNum, n, 1, 1)
                                            Exit Sub
                                        End If

                                     Else

                                        If Player(MapAttributeNpc(MapNum, n, X, Y).owner).Char(Player(MapAttributeNpc(MapNum, n, X, Y).owner).CharNum).X < X Then
                                            Call NPCMove(MapNum, n, 2, 1)
                                            Exit Sub
                                        End If

                                        If Player(MapAttributeNpc(MapNum, n, X, Y).owner).Char(Player(MapAttributeNpc(MapNum, n, X, Y).owner).CharNum).X > X Then
                                            Call NPCMove(MapNum, n, 3, 1)
                                            Exit Sub
                                        End If

                                        If Player(MapAttributeNpc(MapNum, n, X, Y).owner).Char(Player(MapAttributeNpc(MapNum, n, X, Y).owner).CharNum).Y < Y - 1 Then
                                            Call NPCMove(MapNum, n, 0, 1)
                                            Exit Sub
                                        End If

                                        If Player(MapAttributeNpc(MapNum, n, X, Y).owner).Char(Player(MapAttributeNpc(MapNum, n, X, Y).owner).CharNum).Y > Y + 1 Then
                                            Call NPCMove(MapNum, n, 1, 1)
                                            Exit Sub
                                        End If

                                    End If
                                 Else
                                    i = Int(Rnd * 4)

                                    If i = 1 Then
                                        i = Int(Rnd * 4)

                                        If CanAttributeNPCMove(n, X, Y, MapNum, i) Then
                                            Call AttributeNPCMove(n, X, Y, MapNum, i, MOVING_WALKING)
                                        End If

                                    End If
                                End If
                            End If

                            ' /////////////////////////////////////////////
                            ' // This is used for npcs to attack players //
                            ' /////////////////////////////////////////////

                            If 0 + MapAttributeNpc(MapNum, n, X, Y).owner <> 0 Then
                                Target = GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)
                             Else
                                Target = MapAttributeNpc(MapNum, n, X, Y).Target
                            End If

                            ' Check if the npc can attack the targeted player player

                            If Target > 0 Then
                                If 0 + MapAttributeNpc(MapNum, n, X, Y).owner <> 0 Then
                                    If GetPlayerMap(MapAttributeNpc(MapNum, n, X, Y).owner) = MapNum Then
                                        If MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).X = 1 Then
                                            If CanAttributeNpcAttackNpc(MapNum, n, X, Y) Then
                                                'pet attacking npc
                                                Damage = Int(Npc(n).STR * 2) - Int(Npc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).DEF / 2)

                                                If Damage > 0 Then
                                                    MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).HP = MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).HP - Damage
                                                End If

                                                'npc attacking pet
                                                Damage = Int(Npc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).STR * 2) - Int(Npc(n).DEF / 2)

                                                If Damage > 0 Then
                                                    MapNpc(n).HP = MapNpc(n).HP - Damage
                                                End If

                                            End If
                                        End If
                                    End If
                                 Else
                                    ' Is the target playing and on the same map?

                                    If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                        ' Can the npc attack the player?

                                        If CanAttributeNpcAttackPlayer(X, Target) Then
                                            If Not CanPlayerBlockHit(Target) Then
                                                Damage = Npc(NpcNum).STR - GetPlayerProtection(Target)

                                                If Damage > 0 Then
                                                    Call NpcAttackPlayer(X, Target, Damage)
                                                 Else
                                                    Call BattleMsg(Target, "The " & Trim(Npc(NpcNum).Name) & " could not hurt you.", BrightBlue, 1)

                                                    'Call PlayerMsg(Target, "The " & Trim(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                                End If

                                             Else
                                                Call BattleMsg(Target, "You blocked " & Trim(Npc(NpcNum).Name) & "'s hit.", BrightCyan, 1)

                                                'Call PlayerMsg(Target, "Your " & Trim(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            End If

                                        End If
                                     Else
                                        ' Player left map or game, set target to 0
                                        MapAttributeNpc(MapNum, n, X, Y).Target = 0
                                    End If

                                End If
                            End If

                            ' ////////////////////////////////////////////
                            ' // This is used for regenerating NPC's HP //
                            ' ////////////////////////////////////////////
                            ' Check to see if we want to regen some of the npc's hp

                            If GetTickCount > GiveNPCHPTimer + 10000 Then
                                If MapAttributeNpc(MapNum, n, X, Y).HP > 0 Then
                                    MapAttributeNpc(MapNum, n, X, Y).HP = MapAttributeNpc(MapNum, n, X, Y).HP + GetNpcHPRegen(NpcNum)

                                    ' Check if they have more then they should and if so just set it to max

                                    If MapAttributeNpc(MapNum, n, X, Y).HP > GetNpcMaxhp(NpcNum) Then
                                        MapAttributeNpc(MapNum, n, X, Y).HP = GetNpcMaxhp(NpcNum)
                                    End If

                                End If
                            End If

                            ' ////////////////////////////////////////////////////////
                            ' // This is used for checking if an NPC is dead or not //
                            ' ////////////////////////////////////////////////////////
                            ' Check if the npc is dead or not

                            If NpcNum > 0 Then
                                If MapAttributeNpc(MapNum, n, X, Y).HP <= 0 And GetNpcMaxhp(NpcNum) > 0 Then
                                    MapAttributeNpc(MapNum, n, X, Y).num = 0
                                    MapAttributeNpc(MapNum, n, X, Y).SpawnWait = GetTickCount
                                End If

                            End If

                            ' //////////////////////////////////////
                            ' // This is used for spawning an NPC //
                            ' //////////////////////////////////////
                            ' Check if we are supposed to spawn an npc or not

                            If NpcNum <= 0 Then
                                If GetTickCount > MapAttributeNpc(MapNum, n, X, Y).SpawnWait + (Npc(Map(MapNum).tile(X, Y).Data1).SpawnSecs * 1000) Then
                                    Call SpawnAttributeNPC(n, X, Y, MapNum)
                                End If

                            End If
                            Call SendDataToMap(MapNum, PacketID.AttributeNPCHP & SEP_CHAR & n & SEP_CHAR & MapAttributeNpc(MapNum, n, X, Y).HP & SEP_CHAR & GetNpcMaxhp(MapAttributeNpc(MapNum, n, X, Y).num) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR)
                        End If

                    End If
                Next n

            End If
        Next X

    Next Y

End Sub

Sub AttributeNPCMove(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal MapNum As Long, ByVal Dir As Long, ByVal Movement As Long)

  Dim packet As String

    If index > Map(MapNum).tile(X, Y).Data2 Then Exit Sub

    ' Check for subscript out of range

    If MapNum <= 0 Or MapNum > MAX_MAPS Or index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    MapAttributeNpc(MapNum, index, X, Y).Dir = Dir

    Select Case Dir
     Case DIR_UP
        MapAttributeNpc(MapNum, index, X, Y).Y = MapAttributeNpc(MapNum, index, X, Y).Y - 1
        packet = PacketID.AttributeNPCMove & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).X & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Y & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Dir & SEP_CHAR & Movement & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)

     Case DIR_DOWN
        MapAttributeNpc(MapNum, index, X, Y).Y = MapAttributeNpc(MapNum, index, X, Y).Y + 1
        packet = PacketID.AttributeNPCMove & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).X & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Y & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Dir & SEP_CHAR & Movement & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)

     Case DIR_LEFT
        MapAttributeNpc(MapNum, index, X, Y).X = MapAttributeNpc(MapNum, index, X, Y).X - 1
        packet = PacketID.AttributeNPCMove & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).X & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Y & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Dir & SEP_CHAR & Movement & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)

     Case DIR_RIGHT
        MapAttributeNpc(MapNum, index, X, Y).X = MapAttributeNpc(MapNum, index, X, Y).X + 1
        packet = PacketID.AttributeNPCMove & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).X & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Y & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Dir & SEP_CHAR & Movement & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)
    End Select

End Sub

Function CanAttackAttributeNpc(ByVal Attacker As Long, ByVal index As Long, ByVal X As Long, ByVal Y As Long) As Boolean

  Dim AttackSpeed As Long
  Dim NpcNum As Long
  Dim MapNum As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
     Else
        AttackSpeed = 1000
    End If

    CanAttackAttributeNpc = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or index <= 0 Or index > MAX_ATTRIBUTE_NPCS Then Exit Function

    MapNum = GetPlayerMap(Attacker)
    ' Check for subscript out of range
    'If MapAttributeNpc(MapNum, index, x, y).num <= 0 Then Exit Function

    NpcNum = Map(MapNum).tile(X, Y).Data1

    ' Make sure the npc isn't already dead
    'If MapAttributeNpc(MapNum, index, x, y).HP <= 0 Then Exit Function

    ' Make sure they are on the same map
    'If IsPlaying(Attacker) Then

    If GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
        ' Check if at same coordinates

        Select Case GetPlayerDir(Attacker)
         Case DIR_UP

            If (MapAttributeNpc(MapNum, index, X, Y).Y + 1 = GetPlayerY(Attacker)) And (MapAttributeNpc(MapNum, index, X, Y).X = GetPlayerX(Attacker)) Then
                If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    CanAttackAttributeNpc = True
                 Else
                    Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                End If

                Exit Function
            End If

         Case DIR_DOWN

            If (MapAttributeNpc(MapNum, index, X, Y).Y - 1 = GetPlayerY(Attacker)) And (MapAttributeNpc(MapNum, index, X, Y).X = GetPlayerX(Attacker)) Then
                If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    CanAttackAttributeNpc = True
                 Else
                    Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                End If

                Exit Function
            End If

         Case DIR_LEFT

            If (MapAttributeNpc(MapNum, index, X, Y).Y = GetPlayerY(Attacker)) And (MapAttributeNpc(MapNum, index, X, Y).X + 1 = GetPlayerX(Attacker)) Then
                If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    CanAttackAttributeNpc = True
                 Else
                    Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                End If

                Exit Function
            End If

         Case DIR_RIGHT

            If (MapAttributeNpc(MapNum, index, X, Y).Y = GetPlayerY(Attacker)) And (MapAttributeNpc(MapNum, index, X, Y).X - 1 = GetPlayerX(Attacker)) Then
                If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    CanAttackAttributeNpc = True
                 Else
                    Call PlayerMsg(Attacker, Trim(Npc(NpcNum).Name) & " :" & Trim(Npc(NpcNum).AttackSay), Green)
                End If

                Exit Function
            End If

        End Select
    End If
    'End If

End Function

Function CanAttributeNpcAttackNpc(ByVal MapNum, ByVal MapNpcNum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
Dim n As Long

    CanAttributeNpcAttackNpc = False

    ' Check for subscript out of range
    '//!! Warning, there is "NO" trace of what "n" is in this function!
    'Right now just an empty variable is being used in replace of no variable
    If MapNpcNum <= 0 Or MapNpcNum > MAX_ATTRIBUTE_NPCS Or IsPlaying(MapAttributeNpc(MapNum, n, X, Y).owner) = False Then
        Exit Function
    End If

    ' Make sure the npc isn't already dead

    If MapNpc(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second

    If GetTickCount < MapNpc(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    MapNpc(MapNum, MapNpcNum).AttackTimer = GetTickCount

    ' Check if at same coordinates

    If (MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).Y + 1 = MapNpc(MapNum, MapNpcNum).Y) And (MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).X = MapNpc(MapNum, MapNpcNum).X) Then
        CanAttributeNpcAttackNpc = True
     Else

        If (MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).Y - 1 = MapNpc(MapNum, MapNpcNum).Y) And (MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).X = MapNpc(MapNum, MapNpcNum).X) Then
            CanAttributeNpcAttackNpc = True
         Else

            If (MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).Y = MapNpc(MapNum, MapNpcNum).Y) And (MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).X + 1 = MapNpc(MapNum, MapNpcNum).X) Then
                CanAttributeNpcAttackNpc = True
             Else
                
                '//!! old line was:
                'If (MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, x, y).owner)).y = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
                If (MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).Y = MapNpc(MapNum, MapNpcNum).Y) And (MapNpc(GetPlayerTargetNpc(MapAttributeNpc(MapNum, n, X, Y).owner)).X - 1 = MapNpc(MapNum, MapNpcNum).X) Then
                    CanAttributeNpcAttackNpc = True
                End If

            End If
        End If
    End If

End Function

Function CanAttributeNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean

  Dim MapNum As Long
  Dim NpcNum As Long

    CanAttributeNpcAttackPlayer = False

    ' Check for subscript out of range

    If MapNpcNum <= 0 Or MapNpcNum > MAX_ATTRIBUTE_NPCS Or IsPlaying(index) = False Then
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

            If (GetPlayerY(index) + 1 = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).X) Then
                CanAttributeNpcAttackPlayer = True
             Else

                If (GetPlayerY(index) - 1 = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(MapNum, MapNpcNum).X) Then
                    CanAttributeNpcAttackPlayer = True
                 Else

                    If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) + 1 = MapNpc(MapNum, MapNpcNum).X) Then
                        CanAttributeNpcAttackPlayer = True
                     Else

                        If (GetPlayerY(index) = MapNpc(MapNum, MapNpcNum).Y) And (GetPlayerX(index) - 1 = MapNpc(MapNum, MapNpcNum).X) Then
                            CanAttributeNpcAttackPlayer = True
                        End If

                    End If
                End If
            End If

            '            Select Case MapNpc(MapNum, MapNpcNum).Dir
            '                Case DIR_UP
            '                    If (GetPlayerY(Index) + 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
            '                        CanAttributeNpcAttackPlayer = True
            '                    End If
            '
            '                Case DIR_DOWN
            '                    If (GetPlayerY(Index) - 1 = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum, MapNpcNum).x) Then
            '                        CanAttributeNpcAttackPlayer = True
            '                    End If
            '
            '                Case DIR_LEFT
            '                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum, MapNpcNum).x) Then
            '                        CanAttributeNpcAttackPlayer = True
            '                    End If
            '
            '                Case DIR_RIGHT
            '                    If (GetPlayerY(Index) = MapNpc(MapNum, MapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum, MapNpcNum).x) Then
            '                        CanAttributeNpcAttackPlayer = True
            '                    End If
            '            End Select
        End If

    End If

End Function

Function CanAttributeNPCMove(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal MapNum As Long, ByVal Dir) As Boolean

  Dim n As Long
  Dim BX As Long
  Dim BY As Long

    CanAttributeNPCMove = False

    ' Check for subscript out of range

    If MapNum <= 0 Or MapNum > MAX_MAPS Or index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    If index > Map(MapNum).tile(X, Y).Data2 Then Exit Function

    BX = MapAttributeNpc(MapNum, index, X, Y).X
    BY = MapAttributeNpc(MapNum, index, X, Y).Y

    CanAttributeNPCMove = True

    Select Case Dir
     Case DIR_UP
        ' Check to make sure not outside of boundries

        If BY > 0 Then
            n = Map(MapNum).tile(BX, BY - 1).Type

            ' Check to make sure that the tile is walkable

            If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            ' Check to make sure that there is not a player in the way

            If CanAttributeNPCMovePlayer(MapNum, index, X, Y, DIR_UP) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            If CanAttributeNPCMoveAttributeNPC(MapNum, index, X, Y, DIR_UP) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            If CanAttributeNPCMoveNPC(MapNum, index, X, Y, DIR_UP) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

         Else
            CanAttributeNPCMove = False
        End If

     Case DIR_DOWN
        ' Check to make sure not outside of boundries

        If BY < MAX_MAPY Then
            n = Map(MapNum).tile(BX, BY + 1).Type

            ' Check to make sure that the tile is walkable

            If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            ' Check to make sure that there is not a player in the way

            If CanAttributeNPCMovePlayer(MapNum, index, X, Y, DIR_DOWN) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            If CanAttributeNPCMoveAttributeNPC(MapNum, index, X, Y, DIR_DOWN) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            If CanAttributeNPCMoveNPC(MapNum, index, X, Y, DIR_DOWN) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

         Else
            CanAttributeNPCMove = False
        End If

     Case DIR_LEFT
        ' Check to make sure not outside of boundries

        If BX > 0 Then
            n = Map(MapNum).tile(BX - 1, BY).Type

            ' Check to make sure that the tile is walkable

            If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            ' Check to make sure that there is not a player in the way

            If CanAttributeNPCMovePlayer(MapNum, index, X, Y, DIR_LEFT) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            If CanAttributeNPCMoveAttributeNPC(MapNum, index, X, Y, DIR_LEFT) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            If CanAttributeNPCMoveNPC(MapNum, index, X, Y, DIR_LEFT) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

         Else
            CanAttributeNPCMove = False
        End If

     Case DIR_RIGHT
        ' Check to make sure not outside of boundries

        If BX < MAX_MAPX Then
            n = Map(MapNum).tile(BX + 1, BY).Type

            ' Check to make sure that the tile is walkable

            If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPC_SPAWN Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            ' Check to make sure that there is not a player in the way

            If CanAttributeNPCMovePlayer(MapNum, index, X, Y, DIR_RIGHT) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            If CanAttributeNPCMoveAttributeNPC(MapNum, index, X, Y, DIR_RIGHT) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

            If CanAttributeNPCMoveNPC(MapNum, index, X, Y, DIR_RIGHT) = False Then
                CanAttributeNPCMove = False
                Exit Function
            End If

         Else
            CanAttributeNPCMove = False
        End If

    End Select

End Function

Function CanAttributeNPCMoveAttributeNPC(ByVal MapNum As Long, ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal Dir As Long) As Boolean

  Dim i As Long
  Dim BX As Long
  Dim BY As Long

    CanAttributeNPCMoveAttributeNPC = True

    For BX = 0 To MAX_MAPX
        For BY = 0 To MAX_MAPY

            If Map(MapNum).tile(X, Y).Type = TILE_TYPE_NPC_SPAWN Then

                For i = 1 To MAX_ATTRIBUTE_NPCS

                    If i <> index Then
                        If MapAttributeNpc(MapNum, i, BX, BY).num > 0 Then

                            Select Case Dir
                             Case DIR_UP

                                If (MapAttributeNpc(MapNum, index, X, Y).X = MapAttributeNpc(MapNum, i, BX, BY).X) And (MapAttributeNpc(MapNum, index, X, Y).Y - 1 = MapAttributeNpc(MapNum, i, BX, BY).Y) Then
                                    CanAttributeNPCMoveAttributeNPC = False
                                    Exit Function
                                End If

                             Case DIR_DOWN

                                If (MapAttributeNpc(MapNum, index, X, Y).X = MapAttributeNpc(MapNum, i, BX, BY).X) And (MapAttributeNpc(MapNum, index, X, Y).Y + 1 = MapAttributeNpc(MapNum, i, BX, BY).Y) Then
                                    CanAttributeNPCMoveAttributeNPC = False
                                    Exit Function
                                End If

                             Case DIR_LEFT

                                If (MapAttributeNpc(MapNum, index, X, Y).X - 1 = MapAttributeNpc(MapNum, i, BX, BY).X) And (MapAttributeNpc(MapNum, index, X, Y).Y = MapAttributeNpc(MapNum, i, BX, BY).Y) Then
                                    CanAttributeNPCMoveAttributeNPC = False
                                    Exit Function
                                End If

                             Case DIR_RIGHT

                                If (MapAttributeNpc(MapNum, index, X, Y).X + 1 = MapAttributeNpc(MapNum, i, BX, BY).X) And (MapAttributeNpc(MapNum, index, X, Y).Y = MapAttributeNpc(MapNum, i, BX, BY).Y) Then
                                    CanAttributeNPCMoveAttributeNPC = False
                                    Exit Function
                                End If

                            End Select
                        End If
                    End If
                Next i

            End If
        Next BY

    Next BX

End Function

Function CanAttributeNPCMoveNPC(ByVal MapNum As Long, ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal Dir As Long) As Boolean

  Dim i As Long

    CanAttributeNPCMoveNPC = True

    For i = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum, i).num > 0 Then

            Select Case Dir
             Case DIR_UP

                If (MapAttributeNpc(MapNum, index, X, Y).X = MapNpc(MapNum, i).X) And (MapAttributeNpc(MapNum, index, X, Y).Y - 1 = MapNpc(MapNum, i).Y) Then
                    CanAttributeNPCMoveNPC = False
                    Exit Function
                End If

             Case DIR_DOWN

                If (MapAttributeNpc(MapNum, index, X, Y).X = MapNpc(MapNum, i).X) And (MapAttributeNpc(MapNum, index, X, Y).Y + 1 = MapNpc(MapNum, i).Y) Then
                    CanAttributeNPCMoveNPC = False
                    Exit Function
                End If

             Case DIR_LEFT

                If (MapAttributeNpc(MapNum, index, X, Y).X - 1 = MapNpc(MapNum, i).X) And (MapAttributeNpc(MapNum, index, X, Y).Y = MapNpc(MapNum, i).Y) Then
                    CanAttributeNPCMoveNPC = False
                    Exit Function
                End If

             Case DIR_RIGHT

                If (MapAttributeNpc(MapNum, index, X, Y).X + 1 = MapNpc(MapNum, i).X) And (MapAttributeNpc(MapNum, index, X, Y).Y = MapNpc(MapNum, i).Y) Then
                    CanAttributeNPCMoveNPC = False
                    Exit Function
                End If

            End Select
        End If
    Next i

End Function

Function CanAttributeNPCMovePlayer(ByVal MapNum As Long, ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal Dir As Long) As Boolean

  Dim i As Long

    CanAttributeNPCMovePlayer = True

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then

                Select Case Dir
                 Case DIR_UP

                    If (MapAttributeNpc(MapNum, index, X, Y).X = GetPlayerX(i)) And (MapAttributeNpc(MapNum, index, X, Y).Y - 1 = GetPlayerY(i)) Then
                        CanAttributeNPCMovePlayer = False
                        Exit Function
                    End If

                 Case DIR_DOWN

                    If (MapAttributeNpc(MapNum, index, X, Y).X = GetPlayerX(i)) And (MapAttributeNpc(MapNum, index, X, Y).Y + 1 = GetPlayerY(i)) Then
                        CanAttributeNPCMovePlayer = False
                        Exit Function
                    End If

                 Case DIR_LEFT

                    If (MapAttributeNpc(MapNum, index, X, Y).X - 1 = GetPlayerX(i)) And (MapAttributeNpc(MapNum, index, X, Y).Y = GetPlayerY(i)) Then
                        CanAttributeNPCMovePlayer = False
                        Exit Function
                    End If

                 Case DIR_RIGHT

                    If (MapAttributeNpc(MapNum, index, X, Y).X + 1 = GetPlayerX(i)) And (MapAttributeNpc(MapNum, index, X, Y).Y = GetPlayerY(i)) Then
                        CanAttributeNPCMovePlayer = False
                        Exit Function
                    End If

                End Select
            End If
        End If
    Next i

End Function

Function CanNPCMoveAttributeNPC(ByVal MapNum As Long, ByVal index As Long, ByVal Dir As Long) As Boolean

  Dim i As Long
  Dim BX As Long
  Dim BY As Long

    CanNPCMoveAttributeNPC = True

    For BX = 0 To MAX_MAPX
        For BY = 0 To MAX_MAPY

            If Map(MapNum).tile(BX, BY).Type = TILE_TYPE_NPC_SPAWN Then

                For i = 1 To MAX_ATTRIBUTE_NPCS

                    If MapAttributeNpc(MapNum, i, BX, BY).num > 0 Then

                        Select Case Dir
                         Case DIR_UP

                            If (MapNpc(MapNum, index).X = MapAttributeNpc(MapNum, i, BX, BY).X) And (MapNpc(MapNum, index).Y - 1 = MapAttributeNpc(MapNum, i, BX, BY).Y) Then
                                CanNPCMoveAttributeNPC = False
                                Exit Function
                            End If

                         Case DIR_DOWN

                            If (MapNpc(MapNum, index).X = MapAttributeNpc(MapNum, i, BX, BY).X) And (MapNpc(MapNum, index).Y + 1 = MapAttributeNpc(MapNum, i, BX, BY).Y) Then
                                CanNPCMoveAttributeNPC = False
                                Exit Function
                            End If

                         Case DIR_LEFT

                            If (MapNpc(MapNum, index).X - 1 = MapAttributeNpc(MapNum, i, BX, BY).X) And (MapNpc(MapNum, index).Y = MapAttributeNpc(MapNum, i, BX, BY).Y) Then
                                CanNPCMoveAttributeNPC = False
                                Exit Function
                            End If

                         Case DIR_RIGHT

                            If (MapNpc(MapNum, index).X + 1 = MapAttributeNpc(MapNum, i, BX, BY).X) And (MapNpc(MapNum, index).Y = MapAttributeNpc(MapNum, i, BX, BY).Y) Then
                                CanNPCMoveAttributeNPC = False
                                Exit Function
                            End If

                        End Select
                    End If
                Next i

            End If
        Next BY

    Next BX

End Function

Function GetPlayerTargetNpc(ByVal index As Long)

    If index > 0 Then
        GetPlayerTargetNpc = Player(index).targetnpc
    End If

End Function

Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, ByVal spawn_x As Long, ByVal spawn_y As Long, ByVal NpcNum As Long)

    '                         NPC_index               map_number          X spawn          y spawn            NPC_number
  Dim packet As String
  Dim i As Long

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
        MapNpc(MapNum, MapNpcNum).X = 0
        MapNpc(MapNum, MapNpcNum).Y = 0

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

    MapNpc(MapNum, MapNpcNum).X = spawn_x
    MapNpc(MapNum, MapNpcNum).Y = spawn_y

    packet = PacketID.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, packet)

    Call SaveMap(MapNum)

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call SendDataTo(i, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & SEP_CHAR & END_CHAR)
        End If

    Next i

End Sub

Sub SpawnAttributeNPC(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal MapNum As Long)

  Dim packet As String
  Dim NpcNum As Long
  Dim i As Long
  Dim Spawned As Boolean
  Dim BX As Long
  Dim BY As Long
  Dim BX2 As Long
  Dim BY2 As Long
  Dim BX3 As Long
  Dim BY3 As Long

    If index > Map(MapNum).tile(X, Y).Data2 Then Exit Sub

    ' Check for subscript out of range

    If index <= 0 Or index > MAX_ATTRIBUTE_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    Spawned = False

    NpcNum = Map(MapNum).tile(X, Y).Data1
    'If NpcNum > 0 Then

    If GameTime = TIME_NIGHT Then
        If Npc(NpcNum).SpawnTime = 1 Then
            MapAttributeNpc(MapNum, index, X, Y).num = 0
            MapAttributeNpc(MapNum, index, X, Y).SpawnWait = GetTickCount
            MapAttributeNpc(MapNum, index, X, Y).HP = 0
            Call SendDataToMap(MapNum, PacketID.AttributeNPCDead & SEP_CHAR & index & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

     Else

        If Npc(NpcNum).SpawnTime = 2 Then
            MapAttributeNpc(MapNum, index, X, Y).num = 0
            MapAttributeNpc(MapNum, index, X, Y).SpawnWait = GetTickCount
            MapAttributeNpc(MapNum, index, X, Y).HP = 0
            Call SendDataToMap(MapNum, PacketID.AttributeNPCDead & SEP_CHAR & index & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

    End If

    MapAttributeNpc(MapNum, index, X, Y).num = NpcNum
    MapAttributeNpc(MapNum, index, X, Y).Target = 0

    MapAttributeNpc(MapNum, index, X, Y).HP = GetNpcMaxhp(NpcNum)
    MapAttributeNpc(MapNum, index, X, Y).MP = GetNpcMaxMP(NpcNum)
    MapAttributeNpc(MapNum, index, X, Y).SP = GetNpcMaxSP(NpcNum)

    MapAttributeNpc(MapNum, index, X, Y).Dir = Int(Rnd * 4)

    ' Well try 100 times to randomly place the sprite

    If Map(MapNum).tile(X, Y).Data3 > 0 Then
        BX3 = X + Map(MapNum).tile(X, Y).Data3
        BX2 = X - Map(MapNum).tile(X, Y).Data3
        BY3 = Y + Map(MapNum).tile(X, Y).Data3
        BY2 = Y - Map(MapNum).tile(X, Y).Data3

        If BX2 < 0 Then BX2 = 1
        If BX3 > MAX_MAPX Then BX3 = MAX_MAPX
        If BY2 < 0 Then BY2 = 1
        If BY3 > MAX_MAPY Then BY3 = MAX_MAPY

        For i = 1 To 100
            BX = Int(Rand(BX3, BX2))
            BY = Int(Rand(BY3, BY2))

            BX = BX - 1
            BY = BY - 1

            ' Check if the tile is walkable

            If Map(MapNum).tile(BX, BY).Type = TILE_TYPE_WALKABLE Or Map(MapNum).tile(BX, BY).Type = TILE_TYPE_NPC_SPAWN Then
                MapAttributeNpc(MapNum, index, X, Y).X = BX
                MapAttributeNpc(MapNum, index, X, Y).Y = BY
                Spawned = True
                Exit For
            End If

        Next i

        ' Didn't spawn, so now we'll just try to find a free tile

        If Not Spawned Then

            For BY = BY2 To BY3
                For BX = BX2 To BX3

                    If Map(MapNum).tile(BX, BY).Type = TILE_TYPE_WALKABLE Or Map(MapNum).tile(BX, BY).Type = TILE_TYPE_NPC_SPAWN Then
                        MapAttributeNpc(MapNum, index, X, Y).X = BX
                        MapAttributeNpc(MapNum, index, X, Y).Y = BY
                        Spawned = True
                    End If

                Next BX
            Next BY
        End If

     Else
        MapAttributeNpc(MapNum, index, X, Y).X = X
        MapAttributeNpc(MapNum, index, X, Y).Y = Y
        Spawned = True
    End If

    ' If we suceeded in spawning then send it to everyone

    If Spawned Then
        packet = PacketID.SpawnAttributeNPC & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).num & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).X & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Y & SEP_CHAR & MapAttributeNpc(MapNum, index, X, Y).Dir & SEP_CHAR & Npc(MapAttributeNpc(MapNum, index, X, Y).num).Big & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)
    End If

    'End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & index & SEP_CHAR & MapAttributeNpc(MapNum, index, x, y).HP & SEP_CHAR & GetNpcMaxHP(MapAttributeNpc(MapNum, index, x, y).num) & SEP_CHAR & END_CHAR)

End Sub

