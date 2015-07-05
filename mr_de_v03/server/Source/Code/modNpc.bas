Attribute VB_Name = "modNpc"
Option Explicit

'**********
' NPC Data
'**********

'***************************************
' Stats
'***************************************
Public Function Npc_Stat(ByVal NpcNum As Long, ByVal Stat As Stats) As Long
    Npc_Stat = Npc(NpcNum).Stat(Stat)
End Function

'***************************************
' Initial Max Vital
'***************************************
Public Function Npc_MaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    If NpcNum <= 0 Then Exit Function
    Select Case Vital
        Case Vitals.HP
            Npc_MaxVital = Npc(NpcNum).MaxHP
        Case Vitals.MP
            Npc_MaxVital = Npc_Stat(NpcNum, Stats.Wisdom) * 2
        Case Vitals.SP
            Npc_MaxVital = Npc_Stat(NpcNum, Stats.Dexterity) * 2
    End Select
End Function

'***************************************
' Returns exp for a npc (STR * DEF)
'***************************************
Public Function Npc_Exp(ByVal NpcNum As Long) As Long
    Npc_Exp = Npc(NpcNum).MaxEXP * (ExpMod * 0.01)
End Function





'******************
'** Map Npc Data **
'******************

'***************************************
' Check for mod stats - buffs/debuffs
'***************************************
Public Function MapNpc_Stat(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Stat As Stats) As Long
Dim i As Long
Dim SpellNum As Long

    MapNpc_Stat = Npc_Stat(MapData(MapNum).MapNpc(MapNpcNum).Num, Stat)
    For i = 1 To MAX_STATUS
        SpellNum = MapData(MapNum).MapNpc(MapNpcNum).Status(i).SpellNum
        If SpellNum > 0 Then
            If Spell(SpellNum).Type = SPELL_TYPE_BUFF Then
                MapNpc_Stat = MapNpc_Stat + Spell(SpellNum).ModStat(Stat)
            End If
        End If
    Next
End Function

'***************************************
' Calculates the mapnpcs vital regen
'***************************************
Public Function MapNpc_VitalRegen(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Vital As Vitals) As Long
    
    Select Case Vital
        Case Vitals.HP
            MapNpc_VitalRegen = MapNpc_Stat(MapNum, MapNpcNum, Stats.Vitality) \ 3
        Case Vitals.MP
            MapNpc_VitalRegen = MapNpc_Stat(MapNum, MapNpcNum, Stats.Wisdom) \ 3
        Case Vitals.SP
            MapNpc_VitalRegen = MapNpc_Stat(MapNum, MapNpcNum, Stats.Dexterity) \ 3
    End Select
    
    MapNpc_VitalRegen = Clamp(MapNpc_VitalRegen, 0, MAX_LONG)
    
End Function

'***********************************************
' MapNpc vital regen based on current mod stats
'***********************************************
Public Function MapNpc_MaxVital(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Vital As Vitals) As Long
Dim NpcNum As Long
    
    NpcNum = MapData(MapNum).MapNpc(MapNpcNum).Num
    MapNpc_MaxVital = Npc_MaxVital(NpcNum, Vital)
End Function

'***************************************
' MapNpc Current Vital
'***************************************
Public Function MapNpc_Current_Vital(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Vital As Vitals) As Long
    MapNpc_Current_Vital = MapData(MapNum).MapNpc(MapNpcNum).Vital(Vital)
End Function

'***************************************
' MapNpc Updates Vitals
'***************************************
Public Sub MapNpc_Update_Vital(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Vital As Vitals, ByVal Value As Long)
    MapData(MapNum).MapNpc(MapNpcNum).Vital(Vital) = Clamp(Value, 0, MapNpc_MaxVital(MapNum, MapNpcNum, Vital))
    ' If this value changed that means something attacked/healed/whatever and we reset the LastDamaged + 10seconds
    MapData(MapNum).MapNpc(MapNpcNum).LastDamageTaken = GetTickCount + 10000
End Sub

'***************************************
' MapNpc Damage
'***************************************
Public Function MapNpc_Damage(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Long
    MapNpc_Damage = Clamp(MapNpc_Stat(MapNum, MapNpcNum, Stats.Strength) + (MapNpc_Stat(MapNum, MapNpcNum, Stats.Dexterity) \ 2), 0, MAX_LONG)
End Function

'***************************************
' MapNpc Protection
'***************************************
Public Function MapNpc_Protection(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Long
    MapNpc_Protection = Clamp(MapNpc_Stat(MapNum, MapNpcNum, Stats.Vitality) \ 2, 0, MAX_LONG)
End Function

''***************************************
'' Calculate base magic damage
''***************************************
'Public Function MapNpc_MagicDamage(ByVal Index As Long) As Long
'    'Current_MagicDamage = Clamp(Current_Stat(Index, Stats.Intelligence) \ 4, 0, MAX_LONG)
'    MapNpc_MagicDamage = Clamp(((Current_Stat(Index, Stats.Intelligence) \ 2) + (Current_Stat(Index, Stats.Wisdom) \ 5.5)) \ 2, 0, MAX_LONG)
'End Function

'***************************************
' Calculate base magic defense
'***************************************
Public Function MapNpc_MagicProtection(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Long
    'Current_MagicProtection = Clamp(Current_Stat(Index, Stats.Wisdom) \ 6, 0, MAX_LONG)
    MapNpc_MagicProtection = Clamp((MapNpc_Stat(MapNum, MapNpcNum, Stats.Intelligence) + MapNpc_Stat(MapNum, MapNpcNum, Stats.Wisdom)) \ 2, 0, MAX_LONG)
End Function

Public Sub MapNpc_AttackPlayer(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Target As Long)
Dim Damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNum, MapNpcNum, Target) Then
        ' Send this packet so they can see the person attacking
        SendNpcAttack MapNum, MapNpcNum
        ' Check if the player can block the hit
        If Not CanPlayerBlockHit(Target) Then
            ' get the map npcs damage
            Damage = MapNpc_Damage(MapNum, MapNpcNum)
            Damage = Rand(Damage * 0.9, Damage * 1.1)
            ' Subtract the players defense
            Damage = Damage - Rand(Current_Protection(Target) * 0.9, Current_Protection(Target) * 1.1)
            If Damage > 0 Then
                ' Check if would kill the player
                If Damage >= Current_BaseVital(Target, Vitals.HP) Then
                    ' Set NPC target to 0
                    MapData(MapNum).MapNpc(MapNpcNum).Target = 0
                    
                    ' Kill the player
                    OnDeath Target
                Else
                    ' Player not dead, just do the damage
                    Update_BaseVital Target, Vitals.HP, Current_BaseVital(Target, Vitals.HP) - Damage
                    SendActionMsg MapNum, "-" & Damage & " HP", BrightRed, ACTIONMSG_SCROLL, Current_X(Target), Current_Y(Target)
                End If
            Else
                SendActionMsg Current_Map(Target), "Deflected!", BrightCyan, ACTIONMSG_SCROLL, Current_X(Target), Current_Y(Target)
            End If
        Else
            SendActionMsg Current_Map(Target), "Block!", BrightCyan, ACTIONMSG_SCROLL, Current_X(Target), Current_Y(Target)
        End If
    End If
End Sub

'***************************************
' MapNpc Update
'***************************************
Public Sub MapNpc_Update(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim SpellNum As Long
Dim i As Long
Dim n As Long
Dim NpcNum As Long
Dim Attacker As Long
Dim Damage As Long
Dim Modifier As Long
Dim Target As Long
    
    NpcNum = MapData(MapNum).MapNpc(MapNpcNum).Num
    
    '*************************
    '**  Checks for target  **
    '*************************
    Target = MapData(MapNum).MapNpc(MapNpcNum).Target
    If Not MapNpc_IsValidTarget(MapNum, MapNpcNum, Target) Then
        MapNpc_FindNextClosestTarget MapNum, MapNpcNum
    End If
    
    '*********************************
    '**  Checks for status effects  **
    '*********************************
    For i = 1 To MAX_STATUS
        SpellNum = MapData(MapNum).MapNpc(MapNpcNum).Status(i).SpellNum
        If SpellNum > 0 Then
            If GetTickCount >= MapData(MapNum).MapNpc(MapNpcNum).Status(i).TickUpdate Then
                ' if we have an overtime spell - let's do our damage/heal
                Select Case Spell(SpellNum).Type
                    Case SPELL_TYPE_OVERTIME
                        ' find the player that cast the spell
                        Attacker = FindPlayer(MapData(MapNum).MapNpc(MapNpcNum).Status(i).Caster)
                            
                        ' if the player is online then we calculate how much their magic damage affects
                        If Attacker > 0 Then Modifier = Current_MagicDamage(Attacker) \ Spell(SpellNum).TickCount
                            
                        For n = 1 To Vitals.Vital_Count
                            ' Positive - Healing
                            If Spell(SpellNum).ModVital(n) > 0 Then
                                Damage = (Spell(SpellNum).ModVital(n) + Modifier)
                                MapNpc_Update_Vital MapNum, MapNpcNum, n, MapNpc_Current_Vital(MapNum, MapNpcNum, n) + Damage
                                SendActionMsg MapNum, "+" & CStr(Damage) & " " & VitalName(n), BrightGreen, ACTIONMSG_SCROLL, MapData(MapNum).MapNpc(MapNpcNum).X, MapData(MapNum).MapNpc(MapNpcNum).Y
                                ' Check if it's above
                                If MapNpc_Current_Vital(MapNum, MapNpcNum, n) > MapNpc_MaxVital(MapNum, MapNpcNum, n) Then MapNpc_Update_Vital MapNum, MapNpcNum, n, MapNpc_MaxVital(MapNum, MapNpcNum, n)
                            ' Negative - Damage
                            ElseIf Spell(SpellNum).ModVital(n) < 0 Then
                                Damage = Abs(Spell(SpellNum).ModVital(n) - Modifier)
                                ' You can't do more damage than the npcs hp
                                If Damage > MapNpc_Current_Vital(MapNum, MapNpcNum, n) Then Damage = MapNpc_Current_Vital(MapNum, MapNpcNum, n)
                                MapNpc_Update_Vital MapNum, MapNpcNum, n, MapNpc_Current_Vital(MapNum, MapNpcNum, n) - Damage
                                SendActionMsg MapNum, CStr(-Damage) & " " & VitalName(n), Yellow, ACTIONMSG_SCROLL, MapData(MapNum).MapNpc(MapNpcNum).X, MapData(MapNum).MapNpc(MapNpcNum).Y
                                ' Adds the damage and checks it target
                                If Attacker > 0 Then MapNpc_AddDamage MapNum, MapNpcNum, Attacker, Damage
                            End If
                        Next
                                        
                        ' check if it killed them?
                        If MapNpc_Current_Vital(MapNum, MapNpcNum, Vitals.HP) <= 0 Then
                            ' Checks if the player that casted the spell is online, if not it kills the npc
                            If Attacker > 0 Then
                                ' check if they are on the same map
                                ' otherwise it would be easy for someone to cast a dot then run off screen
                                ' and get the exp
                                If Current_Map(Attacker) = MapNum Then
                                    MapNpc_OnDeath MapNum, MapNpcNum
                                Else
                                    MapNpc_Kill MapNum, MapNpcNum
                                End If
                            Else
                                MapNpc_Kill MapNum, MapNpcNum
                            End If
                            Exit Sub
                        End If
                End Select
                
                ' subtract one from the count
                MapData(MapNum).MapNpc(MapNpcNum).Status(i).TickCount = MapData(MapNum).MapNpc(MapNpcNum).Status(i).TickCount - 1
                
                ' check if it's over?
                If MapData(MapNum).MapNpc(MapNpcNum).Status(i).TickCount <= 0 Then
                    MapData(MapNum).MapNpc(MapNpcNum).Status(i).SpellNum = 0
                    MapData(MapNum).MapNpc(MapNpcNum).Status(i).TickCount = 0
                    MapData(MapNum).MapNpc(MapNpcNum).Status(i).TickUpdate = 0
                ' if not set the next tick
                Else
                    MapData(MapNum).MapNpc(MapNpcNum).Status(i).TickUpdate = GetTickCount + (Spell(SpellNum).TickUpdate * 1000)
                End If
            End If
        End If
    Next
    
    '************************************
    '**  Checks for last time damaged  **
    '************************************
    ' Make sure the LastDamageTaken has a value
    If MapData(MapNum).MapNpc(MapNpcNum).LastDamageTaken Then
        ' Check if time is up
        If GetTickCount >= MapData(MapNum).MapNpc(MapNpcNum).LastDamageTaken Then
            ' Restore back to full health and clear the Damage dictionary
            For i = 1 To Vitals.Vital_Count
                MapData(MapNum).MapNpc(MapNpcNum).Vital(i) = Npc_MaxVital(NpcNum, i)
            Next
            MapData(MapNum).MapNpc(MapNpcNum).Damage.RemoveAll
            MapData(MapNum).MapNpc(MapNpcNum).Target = 0
            MapData(MapNum).MapNpc(MapNpcNum).LastDamageTaken = 0
        End If
    End If
End Sub

Public Sub MapNpc_Move(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal NpcNum As Long, ByVal Target As Long)
Dim i As Long, n As Long
Dim Dir As Long
Dim DidWalk As Boolean
                        
    Dir = -1

    ' Check to see if we are following a player or not
    If Target Then
        ' Check if the target is above the npc
        If Current_Y(Target) < MapData(MapNum).MapNpc(MapNpcNum).Y Then
            ' Check if it's directly above - meaning they have the same X coord
            If Current_X(Target) = MapData(MapNum).MapNpc(MapNpcNum).X Then
                ' Move straight up
                Dir = DIR_UP
            Else
                ' Check if the player northeast - above and right
                If Current_X(Target) > MapData(MapNum).MapNpc(MapNpcNum).X Then
                    Select Case Rand(0, 1)
                        Case 0: Dir = DIR_UP
                        Case 1: Dir = DIR_RIGHT
                    End Select
                 ' Else the player is northwest - above and left
                ElseIf Current_X(Target) < MapData(MapNum).MapNpc(MapNpcNum).X Then
                    Select Case Rand(0, 1)
                        Case 0: Dir = DIR_UP
                        Case 1: Dir = DIR_LEFT
                    End Select
                End If
            End If
        ' Else if the player is below the npc
        ElseIf Current_Y(Target) > MapData(MapNum).MapNpc(MapNpcNum).Y Then
            ' Check if it's directly below - meaning they have the same X coord
            If Current_X(Target) = MapData(MapNum).MapNpc(MapNpcNum).X Then
                ' Move straight down
                Dir = DIR_DOWN
            Else
                ' Check if the player southeast - below and right
                If Current_X(Target) > MapData(MapNum).MapNpc(MapNpcNum).X Then
                    Select Case Rand(0, 1)
                        Case 0: Dir = DIR_DOWN
                        Case 1: Dir = DIR_RIGHT
                    End Select
                 ' Else the player is southwest - below and left
                ElseIf Current_X(Target) < MapData(MapNum).MapNpc(MapNpcNum).X Then
                    Select Case Rand(0, 1)
                        Case 0: Dir = DIR_DOWN
                        Case 1: Dir = DIR_LEFT
                    End Select
                End If
            End If
        ' Else we are on the same y
        ElseIf Current_Y(Target) = MapData(MapNum).MapNpc(MapNpcNum).Y Then
            ' Check if to the left or right
            If Current_X(Target) > MapData(MapNum).MapNpc(MapNpcNum).X Then
                ' Move to the right
                Dir = DIR_RIGHT
            ElseIf Current_X(Target) < MapData(MapNum).MapNpc(MapNpcNum).X Then
                ' Move to the left
                Dir = DIR_LEFT
            End If
        End If

        ' If the npc is not running away from the player
        If Not MapData(MapNum).MapNpc(MapNpcNum).IsRunning Then
            ' Move the npc depending on the facing
            If Dir > -1 Then DidWalk = MapNpc_MoveDir(MapNum, MapNpcNum, Dir)

            If Not DidWalk Then
                ' Check if the player is next to the npc
                ' Check if player is above or below
                If Current_X(Target) = MapData(MapNum).MapNpc(MapNpcNum).X Then
                    If Current_Y(Target) + 1 = MapData(MapNum).MapNpc(MapNpcNum).Y Then
                        NpcDir MapNum, MapNpcNum, DIR_UP
                        DidWalk = True
                    End If
                    If Current_Y(Target) - 1 = MapData(MapNum).MapNpc(MapNpcNum).Y Then
                        NpcDir MapNum, MapNpcNum, DIR_DOWN
                        DidWalk = True
                    End If
                End If

                ' Check if the player is left or right
                If Current_Y(Target) = MapData(MapNum).MapNpc(MapNpcNum).Y Then
                    If Current_X(Target) + 1 = MapData(MapNum).MapNpc(MapNpcNum).X Then
                        NpcDir MapNum, MapNpcNum, DIR_LEFT
                        DidWalk = True
                    End If
                    If Current_X(Target) - 1 = MapData(MapNum).MapNpc(MapNpcNum).X Then
                        NpcDir MapNum, MapNpcNum, DIR_RIGHT
                        DidWalk = True
                    End If
                End If
            End If

            ' If we couldn't move for whatever reason
            If Not DidWalk Then
                If Int(Rnd * 2) = 1 Then
                    DidWalk = MapNpc_MoveDir(MapNum, MapNpcNum, Rand(DIR_UP, DIR_RIGHT))
                End If

                ' Try to find the next closest target
                MapNpc_FindNextClosestTarget MapNum, MapNpcNum
                
                MapData(MapNum).MapNpc(MapNpcNum).StepsTaken = MapData(MapNum).MapNpc(MapNpcNum).StepsTaken + 1
'                If MapData(MapNum).MapNpc(MapNpcNum).StepsTaken > 5 Then
'                    MapData(MapNum).MapNpc(MapNpcNum).IsRunning = True
'                    MapData(MapNum).MapNpc(MapNpcNum).StepsTaken = 0
'                End If
            End If
        ' If the NPC is running away
        Else
            ' run away
            If Dir > -1 Then
                Select Case Dir
                    Case DIR_UP: Dir = DIR_DOWN
                    Case DIR_DOWN: Dir = DIR_UP
                    Case DIR_LEFT: Dir = DIR_RIGHT
                    Case DIR_RIGHT: Dir = DIR_LEFT
                End Select
                DidWalk = MapNpc_MoveDir(MapNum, MapNpcNum, Dir)
            End If

            MapData(MapNum).MapNpc(MapNpcNum).StepsTaken = MapData(MapNum).MapNpc(MapNpcNum).StepsTaken + 1
            If MapData(MapNum).MapNpc(MapNpcNum).StepsTaken > 5 Then
                MapData(MapNum).MapNpc(MapNpcNum).IsRunning = False
                MapData(MapNum).MapNpc(MapNpcNum).StepsTaken = 0
            End If
        End If

        ' Check if the npc didn't move
        If Not DidWalk Then MapNpc_FindNextClosestTarget MapNum, MapNpcNum
        
        ' For attacking players
        MapNpc_AttackPlayer MapNum, MapNpcNum, Target
    Else
    ' If the Npc doesn't have a target, then just move around randomally
        Select Case Npc(NpcNum).MovementFrequency
            ' Won't move
            Case 1
                n = 1
            ' Normal
            Case 2
                n = 4
            ' Fast
            Case 3
                n = 2
        End Select

        i = Int(Rnd * n)

        If i = 1 Then
            i = Int(Rnd * 4)
            If CanNpcMove(MapNum, MapNpcNum, i) Then
                NpcMove MapNum, MapNpcNum, i
            End If
        End If
    End If
End Sub

Public Function MapNpc_MoveDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
    If CanNpcMove(MapNum, MapNpcNum, Dir) Then
        NpcMove MapNum, MapNpcNum, Dir
        MapNpc_MoveDir = True
    End If
End Function

Public Sub MapNpc_AddDamage(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Index As Long, ByVal Damage As Long)
Dim Target As Long

    ' Add the players damage to the dictionary
    MapData(MapNum).MapNpc(MapNpcNum).Damage.Item(Current_Name(Index)) = MapData(MapNum).MapNpc(MapNpcNum).Damage.Item(Current_Name(Index)) + Damage

    ' Get the current target of the npc
    Target = MapData(MapNum).MapNpc(MapNpcNum).Target
    
    ' If they have a target
    ' Check it's threat againest the new threat of the index
    If Target > 0 Then
        ' Check to make sure the target is in the dictionary
        If MapData(MapNum).MapNpc(MapNpcNum).Damage.Exists(Current_Name(Target)) Then
            If DamageToThreat(Index, MapData(MapNum).MapNpc(MapNpcNum).Damage.Item(Current_Name(Index))) > DamageToThreat(Target, MapData(MapNum).MapNpc(MapNpcNum).Damage.Item(Current_Name(Target))) Then
                ' TODO: Figure out sort method to go through dictionary incase this target is out of range
                ' Make sure our new target is valid, if not try to find the next closest
                If MapNpc_IsValidTarget(MapNum, MapNpcNum, Index) Then
                    MapData(MapNum).MapNpc(MapNpcNum).Target = Index
                Else
                    MapNpc_FindNextClosestTarget MapNum, MapNpcNum
                End If
            End If
        End If
    Else
        ' No target, so set it to the index
        MapData(MapNum).MapNpc(MapNpcNum).Target = Index
    End If
End Sub

Public Sub MapNpc_OnDeath(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim keyArray As Variant, element As Variant
Dim i As Long, n As Long
Dim QuestNeedsNum As Long, QuestNum As Long
Dim NpcNum As Long
Dim Exp As Long
Dim PartyDict As Dictionary
Dim PartyIndex As Long
Dim PartyUpdateQuest As Boolean
Dim Damage As Long

    NpcNum = MapData(MapNum).MapNpc(MapNpcNum).Num

    Set PartyDict = New Dictionary
    
    ' Get the keyarray
    keyArray = MapData(MapNum).MapNpc(MapNpcNum).Damage.Keys
    For Each element In keyArray
        ' Get the index of the player
        i = FindPlayer(element)
        If i > 0 Then
            ' Get the damage once so we don't have to keep accessing the dictionary
            Damage = MapData(MapNum).MapNpc(MapNpcNum).Damage.Item(element)
            
            ' Make sure it's a valid damage
            If Damage Then
            
                ' Check if in party
                If Not Player(i).InParty Then
                    ' Check if on the same map
                    If Current_Map(i) = MapNum Then
                        ' QUEST STUFF
                        ' If you did at least the QUEST_PERCENT to the Npc, then we will count it towards quest progress
                        If (Damage / Npc_MaxVital(NpcNum, Vitals.HP)) * 100 >= QUEST_PERCENT Then
                            ' Update the players quest progress
                            OnUpdateQuestProgress i, NpcNum, 1, True, QuestTypes.KillNpc
                        End If
                        
                        ' EXP STUFF
                        Exp = Clamp(Npc_Exp(NpcNum) * (1 + 0.1 * (Npc(NpcNum).Level - Current_Level(i))), 0, MAX_LONG)
                        Exp = Exp * (Damage / Npc_MaxVital(NpcNum, Vitals.HP))
                    
                        SendActionMsg MapNum, "+" & Exp & " EXP!", Yellow, ACTIONMSG_SCROLL, Current_X(i), Current_Y(i), i
                        Update_Exp i, Current_Exp(i) + Exp
                    End If
                Else
                    ' If in party, add up all other users in party of who did damage and add it to the dictionary
                    PartyIndex = Player(i).PartyIndex
                    PartyDict.Item(PartyIndex) = PartyDict.Item(PartyIndex) + Damage
                End If
            End If
        End If
    Next
    
    If PartyDict.Count Then
        ' Party exp: Exp is based off the highest player
        keyArray = PartyDict.Keys
        For Each element In keyArray
            ' Get the damage once so we don't have to keep accessing the dictionary
            Damage = PartyDict.Item(element)
            
            ' Make sure it's a valid damage
            If Damage Then
            
                ' QUEST
                ' If you did at least the QUEST_PERCENT to the Npc, then we will count it towards quest progress
                If (Damage / Npc_MaxVital(NpcNum, Vitals.HP)) * 100 >= QUEST_PERCENT Then
                    PartyUpdateQuest = True
                End If
                
                ' EXP
                ' Calculate exp based off the highest level in the party
                ' Exp is then divided evenally among the amount of people in the party
                ' For each person in the party, there is a bonus of 10%, 2 people 20%, 3 people 30%, so on
                Exp = Clamp(Npc_Exp(NpcNum) * (1 + 0.1 * (Npc(NpcNum).Level - Party(element).HighLevel)), 0, MAX_LONG)
                Exp = Exp * (Damage / Npc_MaxVital(NpcNum, Vitals.HP))
                If Party(element).PartyCount > 1 Then
                    Exp = Exp + (Exp * (Party(element).PartyCount * 0.1))
                    Exp = Exp / Party(element).PartyCount
                End If
                
                ' Give exp to each party member
                For i = 1 To MAX_PLAYER_PARTY
                    If Party(element).PartyPlayers(i) <> vbNullString Then
                        n = FindPlayer(Party(element).PartyPlayers(i))
                        If n > 0 Then
                            ' Check if they are on the same map
                            If Current_Map(n) = MapNum Then
                                ' Check if they are in a 10 level range
                                If Current_Level(n) > Party(element).HighLevel - 10 Then
                                    ' Update the players quest progress
                                    OnUpdateQuestProgress n, NpcNum, 1, True, QuestTypes.KillNpc
                                    
                                    SendActionMsg MapNum, "+" & Exp & " EXP!", Yellow, ACTIONMSG_SCROLL, Current_X(n), Current_Y(n), i
                                    Update_Exp n, Current_Exp(n) + Exp
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End If
    
    ' Drop goods
    For i = 1 To 4
        If Npc(NpcNum).Drop(i).Chance > Rand(0, 100) Then
            SpawnItem Npc(NpcNum).Drop(i).Item, Npc(NpcNum).Drop(i).ItemValue, MapNum, MapData(MapNum).MapNpc(MapNpcNum).X, MapData(MapNum).MapNpc(MapNpcNum).Y, GetTickCount
        End If
    Next
    
    MapNpc_Kill MapNum, MapNpcNum
End Sub

Public Sub MapNpc_Kill(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim i As Long

    For i = 1 To MapData(MapNum).MapPlayersCount
        If Player(MapData(MapNum).MapPlayers(i)).TargetType = TARGET_TYPE_NPC Then
            If Player(MapData(MapNum).MapPlayers(i)).Target = MapNpcNum Then
                ChangeTarget MapData(MapNum).MapPlayers(i), 0, TARGET_TYPE_NONE
            End If
        End If
    Next
    
    ClearMapNpc MapNum, MapNpcNum
    
    MapData(MapNum).MapNpc(MapNpcNum).SpawnWait = GetTickCount + (Npc(MapData(MapNum).Npc(MapNpcNum).Npc).SpawnSecs * 1000)

    SendNpcDead MapNum, MapNpcNum
End Sub

Public Function MapNpc_DistanceToPlayer(MapNum As Long, MapNpcNum As Long, Index As Long)
    MapNpc_DistanceToPlayer = Abs(MapData(MapNum).MapNpc(MapNpcNum).X - Current_X(Index)) + Abs(MapData(MapNum).MapNpc(MapNpcNum).Y - Current_Y(Index))
End Function

' This sub will find the closest target that has done any damage to it
Public Sub MapNpc_FindNextClosestTarget(MapNum As Long, MapNpcNum As Long)
Dim keyArray As Variant, element As Variant
Dim i As Long
Dim CurrentTarget As Long
Dim TargetDistance As Long
    
    ' Check if there are any possible targets in the damage dictionary
    If MapData(MapNum).MapNpc(MapNpcNum).Damage.Count Then
        
        ' Get the MapNpcs current target
        CurrentTarget = MapData(MapNum).MapNpc(MapNpcNum).Target
    
        ' Make sure it's a valid target
        If CurrentTarget Then
        
            ' Set TargetDistance to the old targets' distance
            TargetDistance = MapNpc_DistanceToPlayer(MapNum, MapNpcNum, CurrentTarget)
            
            ' Check if there are any other possible targets
            ' Loop through and check the range and if it's not the current target
            keyArray = MapData(MapNum).MapNpc(MapNpcNum).Damage.Keys
            For Each element In keyArray
                ' Get the index of the player
                i = FindPlayer(element)
                If i > 0 Then
                    ' Check it's distance
                    If MapNpc_DistanceToPlayer(MapNum, MapNpcNum, i) < TargetDistance Then
                        ' Now check if it's a valid target
                        If MapNpc_IsValidTarget(MapNum, MapNpcNum, i) Then
                            CurrentTarget = i
                            TargetDistance = MapNpc_DistanceToPlayer(MapNum, MapNpcNum, i)
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    ' Will set our new target
    MapData(MapNum).MapNpc(MapNpcNum).Target = CurrentTarget
End Sub

' This sub will determine if the Index is a valid target
Public Function MapNpc_IsValidTarget(MapNum As Long, MapNpcNum As Long, Index As Long)
    ' Check if they are playing
    If IsPlaying(Index) Then
        ' Make sure they are on the same map
        If Current_Map(Index) = MapNum Then
            ' Check if they are dead
            If Not Current_IsDead(Index) Then
'                ' Check if they are inrange + 2
'                If Not PlayerInRange(Index, MapData(MapNum).MapNpc(MapNpcNum).X, MapData(MapNum).MapNpc(MapNpcNum).Y, Npc(NpcNum).Range + 2) Then
'                    Exit Function
'                End If
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    MapNpc_IsValidTarget = True
End Function

' This sub will get the next highest on the threat
' Need to finish
Public Sub MapNpc_FindNextTarget(ByVal MapNum As Long, ByVal MapNpcNum As Long)
'Dim keyArray As Variant, element As Variant
'Dim i As Long
'Dim TargetsCount As Long
'Dim Targets() As Long
'Dim TargetsThreat() As Long
'Dim CurrentTarget As Long
'
'Dim CurrentTargetThreat As Long
'Dim NewTargets() As Long
'
'    CurrentTarget = MapData(MapNum).MapNpc(MapNpcNum).Target
'
'    keyArray = MapData(MapNum).MapNpc(MapNpcNum).Damage.Keys
'    For Each element In keyArray
'        ' Get the index of the player
'        i = FindPlayer(element)
'        If i > 0 Then
'            ReDim Targets(TargetsCount)
'            Targets(TargetsCount) = i
'            TargetsThreat(TargetsCount) = DamageToThreat(i, MapData(MapNum).MapNpc(MapNpcNum).Damage.Item(element))
'            TargetsCount = TargetsCount + 1
'        End If
'
'        ' Now we sort the array
'        ReDim NewTargets(TargetsCount - 1)
'        For i = 0 To TargetsCount - 1
'            ' This is which place the current target is in the threat list
'            If Targets(TargetsCount) = CurrentTarget Then
'                CurrentTargetThreat = i
'            End If
'        Next
'    Next
End Sub

Public Sub MapNpc_SortDamageTable(ByVal MapNum As Long, ByVal MapNpcNum As Long)

End Sub
