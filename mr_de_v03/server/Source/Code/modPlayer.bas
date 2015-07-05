Attribute VB_Name = "modPlayer"
Option Explicit

Public Sub UpdateOnlinePlayers()
Dim i As Long
Dim ii As Long

    OnlinePlayersCount = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            OnlinePlayersCount = OnlinePlayersCount + 1
        End If
    Next
          
    If OnlinePlayersCount = 0 Then Exit Sub
    
    ReDim OnlinePlayers(1 To OnlinePlayersCount)

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ii = ii + 1
            OnlinePlayers(ii) = i
            
            ' Early finish if all players are found (THANKS MS4)
            If ii >= OnlinePlayersCount Then Exit For
        End If
    Next
End Sub

Public Sub UpdateMapPlayers(ByVal MapNum As Long)
Dim i As Long
Dim ii As Long
    
    ' Adding a check
    If MapNum <= 0 Then Exit Sub
    
    ' Clear the map player count so we can recalculate it
    MapData(MapNum).MapPlayersCount = 0
    
    For i = 1 To OnlinePlayersCount
        If Current_Map(OnlinePlayers(i)) = MapNum Then
            MapData(MapNum).MapPlayersCount = MapData(MapNum).MapPlayersCount + 1
        End If
    Next
    
    If MapData(MapNum).MapPlayersCount = 0 Then Exit Sub
    
    ' Clear the map players array
    ReDim MapData(MapNum).MapPlayers(1 To MapData(MapNum).MapPlayersCount)
    
    ' Loop the OnlinePlayersCount checking for players on this map
    For i = 1 To OnlinePlayersCount
        If Current_Map(OnlinePlayers(i)) = MapNum Then
            ii = ii + 1
            MapData(MapNum).MapPlayers(ii) = OnlinePlayers(i)
        End If
    Next
End Sub

Sub JoinGame(ByVal Index As Long)
Dim i As Long

    ' Set the flag so we know the person is in the game
    Player(Index).InGame = True
        
    ' Update online players
    UpdateOnlinePlayers
    
    ' Check guild
    If Current_Guild(Index) <> 0 Then
        If Current_GuildName(Index) <> GetGuildName(Current_Guild(Index)) Then
            Update_Guild Index, 0
            Update_GuildRank Index, 0
            Update_GuildName Index, vbNullString
        End If
    End If
            
    ' Send a global message that he/she joined
    If Current_Access(Index) <= ADMIN_DEVELOPER Then
        SendGlobalMsg "[Realm Event] " & Current_Name(Index) & " has entered the realm.", JoinLeftColor
    Else
        SendGlobalMsg "[Realm Event] Realm Master " & Current_Name(Index) & " has arrived.", JoinLeftColor
    End If
        
    SendLoginOk Index

    ' Send Game Data
    SendClassesData Index
    SendItems Index
    SendNpcs Index
    SendEmoticons Index
    SendShops Index
    SendSpells Index
    SendAnimations Index
    SendQuests Index
    
    CheckEquippedItems Index
    CheckPlayerInventoryItems Index
    CheckPlayerQuests Index
    
    ' Send the quests
    SendPlayerQuests Index
    
    SendPlayerInv Index
    For i = 1 To Slots.Slot_Count
        SendPlayerWornEq Index, i
    Next
    SendPlayerSpells Index
    
    ' Update mods
    Update_ModStats Index
    Update_ModVitals Index

    SendPlayerGuild Index
    SendPlayerExp Index

    ' Send welcome messages
    SendWelcome Index

    ' Check death
    If Current_IsDead(Index) Then
        If GetTickCount > Current_IsDeadTimer(Index) Then
            SendPlayerMsg Index, "You have been auto-released.", BrightRed
            OnRelease Index
        End If
    End If
    
    ' Warp the player to his saved location
    PlayerWarp Index, Current_Position(Index)

    SendInGame Index

    Player(Index).LastUpdateVitals = GetTickCount + 10000
    Player(Index).LastUpdateSave = GetTickCount + 600000

End Sub

Sub LeftGame(ByVal Index As Long)
Dim i As Long
Dim MapNum As Long
Dim NewPosition As PositionRec

    If Player(Index).InGame Then
        
        MapNum = Current_Map(Index)
        
        For i = 1 To MapData(MapNum).MapPlayersCount
            If Player(MapData(MapNum).MapPlayers(i)).TargetType = TARGET_TYPE_PLAYER Then
                If Player(MapData(MapNum).MapPlayers(i)).Target = Index Then
                    ChangeTarget MapData(MapNum).MapPlayers(i), 0, TARGET_TYPE_NONE
                End If
            End If
        Next
                        
        Player(Index).InGame = False
        
        ' Check for boot map
        If Map(MapNum).BootMap > 0 Then
            NewPosition.Map = Map(MapNum).BootMap
            NewPosition.X = Map(MapNum).BootX
            NewPosition.Y = Map(MapNum).BootY
            Update_Position Index, NewPosition
        End If
                
        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If Player(Index).InParty Then Party_Quit Index
            
        SavePlayer Index
    
        ' Send a global message that he/she left
        If Current_Access(Index) <= ADMIN_MONITER Then
            SendGlobalMsg "[Realm Event] " & Current_Name(Index) & " has left the realm.", JoinLeftColor
        Else
            SendGlobalMsg "[Realm Event] Realm Master " & Current_Name(Index) & " has left.", JoinLeftColor
        End If
        ' Log it
        AddLog Current_Name(Index) & " has disconnected from " & GAME_NAME & ".", PLAYER_LOG
        AddText frmServer.txtText, Current_Name(Index) & " has disconnected from " & GAME_NAME & "."
        
        SendLeftGame Index
    End If
    
    ClearPlayer Index
    
    UpdateOnlinePlayers
    
    ' Check if player was the only player on the map and stop npc processing if so
    UpdateMapPlayers MapNum
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
Dim X As Long, Y As Long

    CanAttackPlayer = False
    
    ' Check for subscript out of range
    If Not IsPlaying(Attacker) Then Exit Function
    If Not IsPlaying(Victim) Then Exit Function
    
    ' Check attack timer
    If GetTickCount < Player(Attacker).AttackTimer + 1000 Then Exit Function
    
    ' Make sure they are on the same map
    If Current_Map(Attacker) <> Current_Map(Victim) Then Exit Function
        
    ' Make sure we dont attack the player if they are switching maps
    If Player(Victim).GettingMap = 1 Then Exit Function
            
    Select Case Current_Dir(Attacker)
        Case DIR_UP
            X = Current_X(Attacker)
            Y = Current_Y(Attacker) - 1
        Case DIR_DOWN
            X = Current_X(Attacker)
            Y = Current_Y(Attacker) + 1
        Case DIR_LEFT
            X = Current_X(Attacker) - 1
            Y = Current_Y(Attacker)
        Case DIR_RIGHT
            X = Current_X(Attacker) + 1
            Y = Current_Y(Attacker)
    End Select
    
    If Y <> Current_Y(Victim) Then Exit Function
    If X <> Current_X(Victim) Then Exit Function
    
    ' Doesn't matter if they are dead
    If Current_IsDead(Victim) Then Exit Function
    
    ' Make sure they have more then 0 hp
    If Current_BaseVital(Victim, Vitals.HP) <= 0 Then Exit Function
    
    If CheckAttackPlayer(Attacker, Victim) Then CanAttackPlayer = True
    
End Function

Public Function CheckAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean

    CheckAttackPlayer = False
    
    If Attacker = Victim Then
        SendActionMsg Current_Map(Attacker), "As much as you would like to, you can't attack yourself...", AlertColor, ACTIONMSG_SCREEN, 0, 0, Attacker
        Exit Function
    End If
    
    ' Check if map is attackable
    If Map(Current_Map(Attacker)).Moral = MAP_MORAL_SAFE Then
        If Current_PK(Victim) = 0 Then
            SendActionMsg Current_Map(Attacker), "This is a haven.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Attacker
            Exit Function
        End If
    End If
        
    ' Check if they are dead
    If Current_IsDead(Victim) Then
        SendActionMsg Current_Map(Attacker), "That player is currently dead.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Attacker
        Exit Function
    End If
    
    ' Check to make sure that they dont have access
    If Current_Access(Attacker) > ADMIN_MONITER Then
        SendActionMsg Current_Map(Attacker), "Realm Masters may not murder.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Attacker
        Exit Function
    End If
        
    ' Check to make sure the victim isn't an admin
    If Current_Access(Victim) > ADMIN_MONITER Then
        SendActionMsg Current_Map(Attacker), "You cannot attack a Realm Master.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Attacker
        Exit Function
    End If
          
    ' Make sure they are high enough level
    If Current_Level(Attacker) < 10 Then
        SendActionMsg Current_Map(Attacker), "You must be level 10+ to fight.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Attacker
        Exit Function
    End If
          
    ' Make sure they are high enough level
    If Current_Level(Victim) < 10 Then
        SendActionMsg Current_Map(Attacker), "They are under level 10.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Attacker
        Exit Function
    End If
    
    If Current_Guild(Attacker) > 0 Then
        If Current_Guild(Attacker) = Current_Guild(Victim) Then
            SendActionMsg Current_Map(Attacker), "Cannot attack guild members.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Attacker
            Exit Function
        End If
    End If
    
    If Player(Attacker).InParty Then
        If Player(Attacker).PartyIndex = Player(Victim).PartyIndex Then
            SendActionMsg Current_Map(Attacker), "Cannot attack party members.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Attacker
            Exit Function
        End If
    End If
    
    ' if you get through the checks, you're golden
    CheckAttackPlayer = True
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long, NpcNum As Long
Dim X As Long, Y As Long

    CanAttackNpc = False
    
    ' Check for subscript out of range
    If Not IsPlaying(Attacker) Then Exit Function
    If MapNpcNum <= 0 Then Exit Function
    If MapNpcNum > MapData(Current_Map(Attacker)).NpcCount Then Exit Function
            
    MapNum = Current_Map(Attacker)
    NpcNum = MapData(MapNum).MapNpc(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapData(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Function
    
    If NpcNum <= 0 Then Exit Function
    If GetTickCount < Player(Attacker).AttackTimer + 1000 Then Exit Function
            
    ' Check if at same coordinates
    Select Case Current_Dir(Attacker)
        Case DIR_UP
            X = MapData(MapNum).MapNpc(MapNpcNum).X
            Y = MapData(MapNum).MapNpc(MapNpcNum).Y + 1
        Case DIR_DOWN
            X = MapData(MapNum).MapNpc(MapNpcNum).X
            Y = MapData(MapNum).MapNpc(MapNpcNum).Y - 1
        Case DIR_LEFT
            X = MapData(MapNum).MapNpc(MapNpcNum).X + 1
            Y = MapData(MapNum).MapNpc(MapNpcNum).Y
        Case DIR_RIGHT
            X = MapData(MapNum).MapNpc(MapNpcNum).X - 1
            Y = MapData(MapNum).MapNpc(MapNpcNum).Y
    End Select
            
    If Y = Current_Y(Attacker) Then
        If X = Current_X(Attacker) Then
            If Npc(NpcNum).Behavior = NPC_BEHAVIOR_FRIENDLY Then Exit Function
            If Npc(NpcNum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then Exit Function
            If Npc(NpcNum).Behavior = NPC_BEHAVIOR_QUEST Then Exit Function
            
            CanAttackNpc = True
        End If
    End If
    
End Function

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long)
Dim MapNum As Long, NpcNum As Long, Damage As Long, ItemNum As Long
Dim Msg As String

    MapNum = Current_Map(Attacker)
    NpcNum = MapData(MapNum).MapNpc(MapNpcNum).Num
    
    ' Set the target to this npc
    ChangeTarget Attacker, MapNpcNum, TARGET_TYPE_NPC
    
    Damage = Current_Damage(Attacker)
    If CanPlayerCriticalHit(Attacker) Then
        Damage = Damage + Int(Rnd * (Damage \ 2)) + 1
        Msg = "Crit! "
    End If
    Damage = Rand(Damage * 0.9, Damage * 1.1) - (Npc(NpcNum).Stat(Stats.Vitality) \ 2)
    
    If Damage <= 0 Then
        SendActionMsg MapNum, "Deflected", BrightRed, ACTIONMSG_SCROLL, MapData(MapNum).MapNpc(MapNpcNum).X, MapData(MapNum).MapNpc(MapNpcNum).Y
        Exit Sub
    End If
    
    ' You can't do more damage than the npcs hp
    If Damage > MapNpc_Current_Vital(MapNum, MapNpcNum, Vitals.HP) Then Damage = MapNpc_Current_Vital(MapNum, MapNpcNum, Vitals.HP)
    
    ' Show the damage
    SendActionMsg Current_Map(Attacker), Msg & "-" & Damage & " HP", BrightRed, ACTIONMSG_SCROLL, MapData(Current_Map(Attacker)).MapNpc(MapNpcNum).X, MapData(Current_Map(Attacker)).MapNpc(MapNpcNum).Y
    
    ' Sends the attack animation for the weapon if there is one
    ItemNum = Current_EquipmentSlot(Attacker, Slots.Weapon)
    If ItemNum > 0 Then
        SendAnimation MapNum, Item(ItemNum).Data2, MapData(MapNum).MapNpc(MapNpcNum).X, MapData(MapNum).MapNpc(MapNpcNum).Y
    End If
    
    ' damage the npc
    MapNpc_Update_Vital MapNum, MapNpcNum, Vitals.HP, MapNpc_Current_Vital(MapNum, MapNpcNum, Vitals.HP) - Damage
    
    ' Adds the damage and checks it target
    MapNpc_AddDamage MapNum, MapNpcNum, Attacker, Damage
    
    ' check if the npc is killed
    If MapNpc_Current_Vital(MapNum, MapNpcNum, Vitals.HP) <= 0 Then MapNpc_OnDeath MapNum, MapNpcNum
    
    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long)
Dim Damage As Long, ItemNum As Long, Exp As Long
Dim Msg As String

    If Not CanPlayerBlockHit(Victim) Then
        Damage = Current_Damage(Attacker)
        If CanPlayerCriticalHit(Attacker) Then
            Damage = Damage + Int(Rnd * (Damage \ 2)) + 1
            Msg = "Crit! "
        End If
        Damage = Rand(Damage * 0.9, Damage * 1.1) - Current_Protection(Victim)
    End If
    
    If Damage < 0 Then
        SendActionMsg Current_Map(Victim), "Deflected", Yellow, ACTIONMSG_SCROLL, Current_X(Victim), Current_Y(Victim)
        Exit Sub
    End If
    
    SendActionMsg Current_Map(Victim), Msg & "-" & Damage & " HP", BrightRed, ACTIONMSG_SCROLL, Current_X(Victim), Current_Y(Victim)
    
    ' Sends the attack animation for the weapon if there is one
    ItemNum = Current_EquipmentSlot(Attacker, Slots.Weapon)
    If ItemNum > 0 Then
        SendAnimation Current_Map(Attacker), Item(ItemNum).Data2, Current_X(Victim), Current_Y(Victim)
    End If
        
    If Damage < Current_BaseVital(Victim, Vitals.HP) Then
        Update_BaseVital Victim, Vitals.HP, Current_BaseVital(Victim, Vitals.HP) - Damage
    Else
        SendActionMsg Current_Map(Attacker), "You have slain " & Current_Name(Victim) & ".", BrightRed, ACTIONMSG_SCREEN, 0, 0, Attacker
        
        Exp = (Current_Exp(Victim) \ 10) * (ExpMod * 0.01)
        If Exp < 0 Then Exp = 0
        If Exp > 0 Then
            SendActionMsg Current_Map(Attacker), "+" & Exp & " EXP!", Yellow, ACTIONMSG_SCROLL, Current_X(Attacker), Current_Y(Attacker), Attacker
            SendActionMsg Current_Map(Victim), "-" & Exp & " EXP!", Yellow, ACTIONMSG_SCROLL, Current_X(Victim), Current_Y(Victim), Victim
            Update_Exp Victim, Current_Exp(Victim) - Exp
            Update_Exp Attacker, Current_Exp(Attacker) + Exp
        End If
        
        If Current_PK(Victim) = 0 Then
            If Current_PK(Attacker) = 0 Then
                Update_PK Attacker, 1
                SendPlayerData (Attacker)
            End If
        Else
            Update_PK Victim, 0
            SendPlayerData (Victim)
        End If
        
        OnDeath Victim
    End If
    
    ' Reset timer for attacking
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub PlayerWarp(ByVal Index As Long, ByRef NewPosition As PositionRec)
Dim OldMap As Long
Dim i As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Sub
    If NewPosition.Map <= 0 Then Exit Sub
    If NewPosition.Map > MAX_MAPS Then Exit Sub
    If NewPosition.X < 0 Then Exit Sub
    If NewPosition.Y < 0 Then Exit Sub
    
    ' Check to see if you are out of bounds
    If NewPosition.X > Map(NewPosition.Map).MaxX Then NewPosition.X = Map(NewPosition.Map).MaxX
    If NewPosition.Y > Map(NewPosition.Map).MaxY Then NewPosition.Y = Map(NewPosition.Map).MaxY
    
    OldMap = Current_Map(Index)
    
    ' Check to see if other players had you targeted
    For i = 1 To MapData(OldMap).MapPlayersCount
        If Player(MapData(OldMap).MapPlayers(i)).TargetType = TARGET_TYPE_PLAYER Then
            If Player(MapData(OldMap).MapPlayers(i)).Target = Index Then
                ChangeTarget MapData(OldMap).MapPlayers(i), 0, TARGET_TYPE_NONE
            End If
        End If
    Next
    
    ' Change your target
    ChangeTarget Index, 0, TARGET_TYPE_NONE
    
    ' Check if they were casting
    CheckCasting Index
    
    ' Save old map to send erase player data to
    SendLeaveMap Index, OldMap
    
    ' This will set the player to the new position
    Update_Position Index, NewPosition
            
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    UpdateMapPlayers OldMap
    
    ' Sets it so we know to process npcs on the map
    UpdateMapPlayers NewPosition.Map
    
    ' Sets the flag that we are now getting a map
    Player(Index).GettingMap = 1
    SendCheckForMap Index, NewPosition.Map
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Moved As Boolean
Dim NewPosition As PositionRec
Dim MapNum As Long, X As Long, Y As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Sub
    If Dir < DIR_UP Then Exit Sub
    If Dir > DIR_RIGHT Then Exit Sub
    If Movement < 1 Then Exit Sub
    If Movement > 2 Then Exit Sub

    Moved = False

    MapNum = Current_Map(Index)
    X = Current_X(Index)
    Y = Current_Y(Index)

    Update_Dir Index, Dir
    
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If Y > 0 Then
                If Not CheckDirection(Index, MapNum, Dir) Then
                    Update_Y Index, Y - 1
                    SendPlayerMove Index
                    Moved = True
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(MapNum).Up > 0 Then
                    NewPosition.Map = Map(MapNum).Up
                    NewPosition.X = X
                    NewPosition.Y = Map(NewPosition.Map).MaxY
                    PlayerWarp Index, NewPosition
                    Exit Sub
                End If
            End If

        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY Then
                If Not CheckDirection(Index, MapNum, Dir) Then
                    Update_Y Index, Y + 1
                    SendPlayerMove Index
                    Moved = True
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(MapNum).Down > 0 Then
                    NewPosition.Map = Map(MapNum).Down
                    NewPosition.X = X
                    NewPosition.Y = 0
                    PlayerWarp Index, NewPosition
                    Exit Sub
                End If
            End If

        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If X > 0 Then
                If Not CheckDirection(Index, MapNum, Dir) Then
                    Update_X Index, X - 1
                    SendPlayerMove Index
                    Moved = True
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(MapNum).Left > 0 Then
                    NewPosition.Map = Map(MapNum).Left
                    NewPosition.X = Map(NewPosition.Map).MaxX
                    NewPosition.Y = Y
                    PlayerWarp Index, NewPosition
                    Exit Sub
                End If
            End If

        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MaxX Then
                If Not CheckDirection(Index, MapNum, Dir) Then
                    Update_X Index, X + 1
                    SendPlayerMove Index
                    Moved = True
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(MapNum).Right > 0 Then
                    NewPosition.Map = Map(MapNum).Right
                    NewPosition.X = 0
                    NewPosition.Y = Y
                    PlayerWarp Index, NewPosition
                    Exit Sub
                End If
            End If
    End Select
    
    ' need to get these values again for map keys
    X = Current_X(Index)
    Y = Current_Y(Index)

    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WARP Then
        NewPosition.Map = Map(MapNum).Tile(X, Y).Data1
        NewPosition.X = Map(MapNum).Tile(X, Y).Data2
        NewPosition.Y = Map(MapNum).Tile(X, Y).Data3
        PlayerWarp Index, NewPosition
        Exit Sub
    End If

    ' Check for key trigger open
    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_KEYOPEN Then
        X = Map(MapNum).Tile(Current_X(Index), Current_Y(Index)).Data1
        Y = Map(MapNum).Tile(Current_X(Index), Current_Y(Index)).Data2

        If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_KEY And Not MapData(MapNum).TempTile.DoorOpen(X, Y) Then
            MapData(MapNum).TempTile.DoorOpen(X, Y) = True
            MapData(MapNum).TempTile.DoorTimer(X, Y) = GetTickCount + 5000

            SendMapKey MapNum, X, Y, 1
            SendActionMsg MapNum, "A door has opened.", AlertColor, ACTIONMSG_SCREEN, 0, 0
        End If
    End If

    ' if they can't move to the new spot warp them back
    If Not Moved Then
        PlayerWarp Index, Current_Position(Index)
    End If
End Sub

Function CheckDirection(ByVal Index As Long, ByVal MapNum As Long, ByVal Direction As Byte) As Boolean
Dim X As Long
Dim Y As Long

    CheckDirection = False
    
    Select Case Direction
        Case DIR_UP
            X = Current_X(Index)
            Y = Current_Y(Index) - 1
        Case DIR_DOWN
            X = Current_X(Index)
            Y = Current_Y(Index) + 1
        Case DIR_LEFT
            X = Current_X(Index) - 1
            Y = Current_Y(Index)
        Case DIR_RIGHT
            X = Current_X(Index) + 1
            Y = Current_Y(Index)
    End Select
                
    ' Check to see if the map tile is blocked or not
    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check for item block
    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_ITEM Then
        If Map(MapNum).Tile(X, Y).Data3 Then
            CheckDirection = True
            Exit Function
        End If
    End If
                                
    ' Check to see if the key door is open or not
    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_KEY Then
        ' This actually checks if its open or not
        If Not MapData(MapNum).TempTile.DoorOpen(X, Y) Then
            CheckDirection = True
            Exit Function
        End If
    End If
End Function

Sub CheckEquippedItems(ByVal Index As Long)
Dim Slot As Long, i As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Sub

    For i = 1 To Slots.Slot_Count
        Slot = Current_EquipmentSlot(Index, i)
        If Slot > 0 Then
            If Len(Trim$(Item(Slot).Name)) <= 0 Then OnUnequipSlot Index, i
            If Item(Slot).Data1 <> i Then OnUnequipSlot Index, i
            If Not CanUseItem(Index, Slot) Then OnUnequipSlot Index, i
        End If
    Next
End Sub

Sub CheckPlayerInventoryItems(ByVal Index As Long)
Dim i As Long
    
    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Sub

    For i = 1 To MAX_INV
        If (Current_InvItemNum(Index, i) > 0) And (Current_InvItemNum(Index, i) <= MAX_ITEMS) Then
        
            ' TODO: Come up with a better way to see if the item is still there...
            If Len(Trim$(Item(Current_InvItemNum(Index, i)).Name)) > 0 Then
                ' Check to see if stack and stack size has changed
                If Item(Current_InvItemNum(Index, i)).Stack = 1 Then
                    ' Check to see if the value is greater than it's max stack
                    If Current_InvItemValue(Index, i) > Item(Current_InvItemNum(Index, i)).StackMax Then
                        ' Set the item to the Stack Max
                        Update_InvItemValue Index, i, Item(Current_InvItemNum(Index, i)).StackMax
                        SendPlayerInvUpdate Index, i
                    End If
                Else
                    ' If the item isn't stackable and has more than 1 set the item value to 1
                    If Current_InvItemValue(Index, i) > 1 Then
                        ' Set the item value to 1
                        Update_InvItemValue Index, i, 1
                        SendPlayerInvUpdate Index, i
                    End If
                End If
            Else
                ' The item name was deleted so delete it from your inv
                TakeInventoryItem Index, i, 1
                SendPlayerInvUpdate Index, i
            End If
        End If
    Next
End Sub

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If Current_Spell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next
End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If Current_Spell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next
End Function

Function FindNextOpenStack(ByVal Index As Long, ByVal ItemNum As Long, ByVal Value As Long) As Long
Dim n As Long
Dim i As Long

    For i = 1 To Value
        n = FindOpenInvSlot(Index, ItemNum)
       
        If n <> 0 Then
            ' Set item slot
            Update_InvItemNum Index, n, ItemNum
            
            ' Check if it's bind on pickup, if so bind it
            If Item(ItemNum).Bound = ItemBind.BindOnPickup Then
                ' Bind it
                Update_InvItemBound Index, n, True
            End If
           
            If Item(Current_InvItemNum(Index, n)).Stack = 1 Then
                If Current_InvItemValue(Index, n) + 1 > Item(Current_InvItemNum(Index, n)).StackMax Then
                    ' self to set to yet another inv slot
                    FindNextOpenStack Index, ItemNum, Value
                Else
                    Update_InvItemValue Index, n, Current_InvItemValue(Index, n) + 1
                End If
            End If
           
            ' Subtract 1 from value just in case we exced max stack and we don't have room for it
            Value = Value - 1
        Else
            FindNextOpenStack = Value
        End If
    Next
End Function

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim OldBound As Boolean
Dim NewNum As Long
Dim NewValue As Long
Dim NewBound As Boolean
Dim OverFlow As Long

        If OldSlot <= 0 Then Exit Sub
    If OldSlot > MAX_INV Then Exit Sub
    If NewSlot <= 0 Then Exit Sub
    If NewSlot > MAX_INV Then Exit Sub
    
    OldNum = Current_InvItemNum(Index, OldSlot)
    OldValue = Current_InvItemValue(Index, OldSlot)
    OldBound = Current_InvItemBound(Index, OldSlot)
    
    NewNum = Current_InvItemNum(Index, NewSlot)
    NewValue = Current_InvItemValue(Index, NewSlot)
    NewBound = Current_InvItemBound(Index, NewSlot)
    
    ' Combine item values if same
    If OldNum > 0 Then
        If NewNum > 0 Then
            If OldNum = NewNum Then
                ' Check to see if stackable
                If Item(OldNum).Stack And Item(NewNum).Stack Then
                    ' If the newvalue is at max value it wasn't switching inv slots
                    ' Added below check to try to fix it
                    If NewValue < Item(NewNum).StackMax Then
                        ' Check if the item values will overflow
                        If OldValue + NewValue > Item(NewNum).StackMax Then
                            OverFlow = Item(NewNum).StackMax - OldValue
                            OldValue = OldValue + OverFlow
                            NewValue = NewValue - OverFlow
                            If OldValue <= 0 Then
                                NewNum = 0
                                NewValue = 0
                                NewBound = False
                            End If
                        Else
                            OldValue = OldValue + NewValue
                            NewNum = 0
                            NewValue = 0
                            NewBound = False
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Update_InvItemNum Index, NewSlot, OldNum
    Update_InvItemValue Index, NewSlot, OldValue
    Update_InvItemBound Index, NewSlot, OldBound
    
    Update_InvItemNum Index, OldSlot, NewNum
    Update_InvItemValue Index, OldSlot, NewValue
    Update_InvItemBound Index, OldSlot, NewBound
        
    SendPlayerInvUpdate Index, NewSlot
    SendPlayerInvUpdate Index, OldSlot
End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    
    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Function
    If ItemNum <= 0 Then Exit Function
    If ItemNum > MAX_ITEMS Then Exit Function
    
    ' Check if the item is stackable
    If Item(ItemNum).Stack Then
        ' If stackable then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If Current_InvItemNum(Index, i) = ItemNum Then
                ' Check to see if we're at max stack for this item
                If Current_InvItemValue(Index, i) < Item(ItemNum).StackMax Then
                    FindOpenInvSlot = i
                    Exit Function
                End If
            End If
        Next
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If Current_InvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next
End Function

Sub TakeInventoryItem(ByVal Index As Long, ByVal InvNum As Long, ByVal ItemVal As Long)

    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Sub
    If InvNum <= 0 Then Exit Sub
    If InvNum > MAX_INV Then Exit Sub
    
    If Item(Current_InvItemNum(Index, InvNum)).Stack = 1 Then
        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= Current_InvItemValue(Index, InvNum) Then
            Update_InvItem Index, InvNum, 0, 0, False
            Exit Sub
        Else
            Update_InvItemValue Index, InvNum, Current_InvItemValue(Index, InvNum) - ItemVal
            Exit Sub
        End If
    Else
        Update_InvItem Index, InvNum, 0, 0, False
        Exit Sub
    End If

End Sub

Function CanTakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean

    CanTakeItem = False
    
    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Function
    If ItemNum <= 0 Then Exit Function
    If ItemNum > MAX_ITEMS Then Exit Function
    
    If Item(ItemNum).Stack Then
        If TakeStackedItem(Index, ItemNum, ItemVal) Then
            CanTakeItem = True
            SendPlayerInv (Index)
            Exit Function
        End If
    Else
        If TakeItem(Index, ItemNum, ItemVal) Then
            CanTakeItem = True
            SendPlayerInv (Index)
            Exit Function
        End If
    End If
End Function

Function TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
Dim i As Long, n As Long

    TakeItem = False
    
    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Function
    If ItemNum <= 0 Then Exit Function
    If ItemNum > MAX_ITEMS Then Exit Function
   
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If Current_InvItemNum(Index, i) = ItemNum Then
            ' Add up what we have to see if we have enough
            n = Current_InvItemValue(Index, i) + n
        End If
    Next
    
    ' we don't have enough , exit
    If n < ItemVal Then Exit Function
    
    ' Loop through inv
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If Current_InvItemNum(Index, i) = ItemNum Then
            ' Take away 1 from the slot
            TakeInventoryItem Index, i, 1
            ' Subtract 1 from our itemval so we can check if we used the right amount
            ItemVal = ItemVal - 1
            ' If we have 0 val left then we are done and we took what we needed
            If ItemVal <= 0 Then
                TakeItem = True
                Exit Function
            End If
        End If
    Next
End Function

Function TakeStackedItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
Dim i As Long, ii As Long, n As Long

    TakeStackedItem = False
    
    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Function
    If ItemNum <= 0 Then Exit Function
    If ItemNum > MAX_ITEMS Then Exit Function
   
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If Current_InvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Stack Then
                ' Add up what we have to see if we have enough
                n = Current_InvItemValue(Index, i) + n
            End If
        End If
    Next
    
    ' we don't have enough , exit
    If n < ItemVal Then Exit Function
    
    ' Loop through inv
    For i = 1 To MAX_INV
        ' Check to see if current inv is the item we want
        If Current_InvItemNum(Index, i) = ItemNum Then
            If Item(ItemNum).Stack Then
                ' Loop to the max value for the invslot and take away 1 till it's no more
                For ii = 1 To Current_InvItemValue(Index, i)
                    ' Take away 1 from the slot
                    TakeInventoryItem Index, i, 1
                    ' Subtract 1 from our itemval so we can check if we used the right amount
                    ItemVal = ItemVal - 1
                    ' If we have 0 val left then we are done and we took what we needed
                    If ItemVal <= 0 Then
                        TakeStackedItem = True
                        Exit Function
                    End If
                Next
            End If
        End If
    Next
End Function

Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim i As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Sub
    If ItemNum <= 0 Then Exit Sub
    If ItemNum > MAX_ITEMS Then Exit Sub
    
    i = FindOpenInvSlot(Index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        If Item(ItemNum).Stack Then
            FindNextOpenStack Index, ItemNum, ItemVal
        Else
            Update_InvItemNum Index, i, ItemNum
            Update_InvItemValue Index, i, Current_InvItemValue(Index, i) + ItemVal
            ' Check if it's bind on pickup, if so bind it
            If Item(ItemNum).Bound = ItemBind.BindOnPickup Then
                ' Bind it
                Update_InvItemBound Index, i, True
            End If
        End If
        
        SendPlayerInv Index
    Else
        SendActionMsg Current_Map(Index), "You are fully burdened.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
    End If
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
    
    ZeroMemory ByVal VarPtr(Player(Index)), LenB(Player(Index))
    
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString

    Player(Index).Char.Name = vbNullString
    Player(Index).Char.GuildName = vbNullString

    Set Player(Index).Buffer = New clsBuffer
    Player(Index).InGame = False
End Sub

Sub ClearChar(ByVal Index As Long)

    ZeroMemory ByVal VarPtr(Player(Index).Char), LenB(Player(Index).Char)
   
    Player(Index).Char.Name = vbNullString
    Player(Index).Char.GuildName = vbNullString
    
    Update_IsDead Index, False
End Sub

Sub ClearStatusEffects(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Player(Index).Char.Status(1)), LenB(Player(Index).Char.Status(1))
    Update_ModStats Index
    Update_ModVitals Index
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Public Function Current_IP(ByVal Index As Long) As String
    Current_IP = frmServer.Socket(Index).RemoteHostIP
End Function

'///////////////
'// User Data //
'///////////////

'***************************************
' Login name
'***************************************
Function Current_Login(ByVal Index As Long) As String
    Current_Login = Trim$(Player(Index).Login)
End Function
Sub Update_Login(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

'***************************************
' Password
'***************************************
Function Current_Password(ByVal Index As Long) As String
    Current_Password = Trim$(Player(Index).Password)
End Function
Sub Update_Password(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

'///////////////
'// Char Data //
'///////////////

'***************************************
' CharNum
'***************************************
Public Function Current_CharNum(ByVal Index As Long) As Long
    Current_CharNum = Player(Index).CharNum
End Function
Sub Update_CharNum(ByVal Index As Long, ByVal CharNum As Long)
    Player(Index).CharNum = CharNum
End Sub

'***************************************
' CharName
'***************************************
Public Function Current_Name(ByVal Index As Long) As String
    Current_Name = Trim$(Player(Index).Char.Name)
End Function
Sub Update_Name(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char.Name = Name
End Sub

'***************************************
' Class
'***************************************
Function Current_Class(ByVal Index As Long) As Long
    Current_Class = Player(Index).Char.Class
End Function
Sub Update_Class(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char.Class = ClassNum
End Sub

'***************************************
' Sprite
'***************************************
Public Function Current_Sprite(ByVal Index As Long) As Long
    Current_Sprite = Player(Index).Char.Sprite
End Function
Sub Update_Sprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char.Sprite = Sprite
End Sub

'***************************************
' Level
'***************************************
Public Function Current_Level(ByVal Index As Long) As Long
    Current_Level = Player(Index).Char.Level
End Function
Sub Update_Level(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Char.Level = Level
End Sub

'***************************************
' Experience
'***************************************
Public Function Current_Exp(ByVal Index As Long) As Long
    Current_Exp = Player(Index).Char.Exp
End Function
Sub Update_Exp(ByVal Index As Long, ByVal Exp As Long)
    ' Can't gain exp while dead
    If Current_IsDead(Index) Then Exit Sub
    If Current_Level(Index) < MAX_LEVEL Then
        Player(Index).Char.Exp = Exp
        ' If this has changed that means we need to send it
        SendPlayerExp Index
        ' Check for a level up
        If Current_Exp(Index) >= Current_NextLevel(Index) Then OnLevelUp Index
    End If
End Sub

'***************************************
' Access Level
'***************************************
Public Function Current_Access(ByVal Index As Long) As Long
    Current_Access = Player(Index).Char.Access
End Function
Sub Update_Access(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char.Access = Access
End Sub

'***************************************
' PK Status
'***************************************
Public Function Current_PK(ByVal Index As Long) As Long
    Current_PK = Player(Index).Char.PK
End Function
Sub Update_PK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char.PK = PK
End Sub

'***************************************
' Vitals
'***************************************
Public Function Current_BaseVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Current_BaseVital = Player(Index).Char.Vital(Vital)
End Function
Sub Update_BaseVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Char.Vital(Vital) = Clamp(Value, 0, Current_MaxVital(Index, Vital))
    ' Since the vital has changed - let's send it
    SendVital Index, Vital
End Sub

'***************************************
' Stats
'***************************************
Public Function Current_BaseStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Current_BaseStat = Player(Index).Char.Stat(Stat)
End Function
Public Sub Update_BaseStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).Char.Stat(Stat) = Value
End Sub

'***************************************
' Points
'***************************************
Public Function Current_Points(ByVal Index As Long) As Long
    Current_Points = Player(Index).Char.Points
End Function
Sub Update_Points(ByVal Index As Long, ByVal Points As Long)
    Player(Index).Char.Points = Points
End Sub

'***************************************
' Map Position
'***************************************
Public Function Current_Position(ByVal Index As Long) As PositionRec
    Current_Position = Player(Index).Char.Position
End Function
Sub Update_Position(ByVal Index As Long, ByRef NewPosition As PositionRec)
    Player(Index).Char.Position = NewPosition
End Sub
'***************************************
' Map
'***************************************
Public Function Current_Map(ByVal Index As Long) As Long
    Current_Map = Player(Index).Char.Position.Map
End Function
Sub Update_Map(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char.Position.Map = MapNum
    End If
End Sub
'***************************************
' X
'***************************************
Public Function Current_X(ByVal Index As Long) As Long
    Current_X = Player(Index).Char.Position.X
End Function
Sub Update_X(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char.Position.X = X
End Sub
'***************************************
' Y
'***************************************
Public Function Current_Y(ByVal Index As Long) As Long
    Current_Y = Player(Index).Char.Position.Y
End Function
Sub Update_Y(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char.Position.Y = Y
End Sub

'***************************************
' Bound Position
'***************************************
Public Function Current_Bound(ByVal Index As Long) As PositionRec
    Current_Bound = Player(Index).Char.Bound
End Function
Sub Update_Bound(ByVal Index As Long, ByRef NewBound As PositionRec)
    Player(Index).Char.Bound = NewBound
End Sub
'***************************************
' Bound Map
'***************************************
Public Function Current_BoundMap(ByVal Index As Long) As Long
    Current_BoundMap = Player(Index).Char.Bound.Map
End Function
Sub Update_BoundMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char.Bound.Map = MapNum
    End If
End Sub
'***************************************
' Bound X
'***************************************
Public Function Current_BoundX(ByVal Index As Long) As Long
    Current_BoundX = Player(Index).Char.Bound.X
End Function
Sub Update_BoundX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char.Bound.X = X
End Sub
'***************************************
' Bound Y
'***************************************
Public Function Current_BoundY(ByVal Index As Long) As Long
    Current_BoundY = Player(Index).Char.Bound.Y
End Function
Sub Update_BoundY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char.Bound.Y = Y
End Sub

'***************************************
' Dir
'***************************************
Function Current_Dir(ByVal Index As Long) As Long
    Current_Dir = Player(Index).Char.Dir
End Function
Sub Update_Dir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char.Dir = Dir
End Sub

'***************************************
' Gold
'***************************************
Function Current_Gold(ByVal Index As Long) As Long
    Current_Gold = Player(Index).Char.Gold
End Function
Sub Update_Gold(ByVal Index As Long, ByVal Value As Long)
    Player(Index).Char.Gold = Value
    ' Now send it
    SendPlayerGold Index
End Sub

'***************************************
' InvItemNum
'***************************************
Public Function Current_InvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    Current_InvItemNum = Player(Index).Char.Inv(InvSlot).Num
End Function
Sub Update_InvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char.Inv(InvSlot).Num = ItemNum
    ' Update the players item collection quests
    OnUpdateQuestProgress Index, ItemNum, Current_InvItemCount(Index, ItemNum), False, QuestTypes.ItemCollection
End Sub

'***************************************
' InvItemValue
'***************************************
Public Function Current_InvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    Current_InvItemValue = Player(Index).Char.Inv(InvSlot).Value
End Function
Sub Update_InvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char.Inv(InvSlot).Value = ItemValue
    ' Update the players item collection quests
    OnUpdateQuestProgress Index, Current_InvItemNum(Index, InvSlot), Current_InvItemCount(Index, Current_InvItemNum(Index, InvSlot)), False, QuestTypes.ItemCollection
End Sub

'***************************************
' Inv Item Bound
'***************************************
Public Function Current_InvItemBound(ByVal Index As Long, ByVal InvSlot As Long) As Boolean
    Current_InvItemBound = Player(Index).Char.Inv(InvSlot).Bound
End Function
Sub Update_InvItemBound(ByVal Index As Long, ByVal InvSlot As Long, ByVal Bound As Boolean)
    Player(Index).Char.Inv(InvSlot).Bound = Bound
End Sub

'***************************************
' Inv Item
'***************************************
Public Sub Update_InvItem(ByVal Index As Long, ByVal InvNum As Long, ByVal ItemNum As Long, ByVal Value As Long, ByVal Bound As Boolean)
    Update_InvItemNum Index, InvNum, ItemNum
    Update_InvItemValue Index, InvNum, Value
    Update_InvItemBound Index, InvNum, Bound
    SendPlayerInvUpdate Index, InvNum
End Sub

'***************************************
' Spell
'***************************************
Public Function Current_Spell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    Current_Spell = Player(Index).Char.Spell(SpellSlot).SpellNum
End Function
Sub Update_Spell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char.Spell(SpellSlot).SpellNum = SpellNum
    ' send the spells
    SendPlayerSpells Index
End Sub

'***************************************
' Spell Cooldown
'***************************************
Public Function Current_SpellCooldown(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    Current_SpellCooldown = Player(Index).Char.Spell(SpellSlot).Cooldown
End Function
Sub Update_SpellCooldown(ByVal Index As Long, ByVal SpellSlot As Long, ByVal Cooldown As Long)
    Player(Index).Char.Spell(SpellSlot).Cooldown = Cooldown
End Sub

'***************************************
' Equipment
'***************************************
Public Function Current_EquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Slots) As Long
    Current_EquipmentSlot = Player(Index).Char.Equipment(EquipmentSlot)
End Function
Sub Update_EquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Slots, ByVal InvNum As Long)
    Player(Index).Char.Equipment(EquipmentSlot) = InvNum
    ' If this has changed, we will need to send the update to the client
    SendPlayerWornEq Index, EquipmentSlot
    ' Update mods and then send vitals
    Update_ModStats Index
    Update_ModVitals Index
End Sub

'***************************************
' Guild
'***************************************
Function Current_Guild(ByVal Index As Long) As Long
    Current_Guild = Player(Index).Char.Guild
End Function
Sub Update_Guild(ByVal Index As Long, ByVal GuildNum As Long)
    Player(Index).Char.Guild = GuildNum
End Sub

'***************************************
' Guild rank
'***************************************
Function Current_GuildRank(ByVal Index As Long) As Long
    Current_GuildRank = Player(Index).Char.GuildRank
End Function
Sub Update_GuildRank(ByVal Index As Long, ByVal GuildRank As Long)
    Player(Index).Char.GuildRank = GuildRank
End Sub

'***************************************
' Guild Name
'***************************************
Function Current_GuildName(ByVal Index As Long) As String
    Current_GuildName = Trim$(Player(Index).Char.GuildName)
End Function
Sub Update_GuildName(ByVal Index As Long, ByVal GuildName As String)
    Player(Index).Char.GuildName = GuildName
End Sub

'***************************************
' Guild Abbreviation
'***************************************
Function Current_GuildAbbreviation(ByVal Index As Long) As String
    Current_GuildAbbreviation = GetGuildAbbreviation(Current_Guild(Index))
End Function

'***************************************
' IsDead
'***************************************
Function Current_IsDead(ByVal Index As Long) As Boolean
    Current_IsDead = Player(Index).Char.IsDead
End Function
Sub Update_IsDead(ByVal Index As Long, ByVal Dead As Boolean)
    Player(Index).Char.IsDead = Dead
End Sub

'***************************************
' IsDeadTimer
'***************************************
Function Current_IsDeadTimer(ByVal Index As Long) As Long
    Current_IsDeadTimer = Player(Index).Char.IsDeadTimer
End Function
Sub Update_IsDeadTimer(ByVal Index As Long, ByVal Time As Long)
    Player(Index).Char.IsDeadTimer = Time
End Sub

'///////////
'// Other //
'///////////

'***************************************
' Check to see if you can use an item
'***************************************
Public Function CanUseItem(ByVal Index As Long, ByVal ItemNum As Long) As Boolean
Dim i As Long

    CanUseItem = False
    
    ' Check for class requirement
    ' Will check your current class to the item
    ' Checks the binary flag is set for your class
    If Not Item(ItemNum).ClassReq And (2 ^ Current_Class(Index)) Then
        SendActionMsg Current_Map(Index), "[Your class can not use this item.]", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
        Exit Function
    End If
    
    ' Check for level requirement
    If Item(ItemNum).LevelReq > 0 Then
        ' If there's a level requirement then check if you can use it
        ' Checks if your level is below the req and if so - will exit
        If Current_Level(Index) < Item(ItemNum).LevelReq Then
            SendActionMsg Current_Map(Index), "[Level Req: " & Item(ItemNum).LevelReq & "]", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
            Exit Function
        End If
    End If
    
    ' Check for stat requirement
    ' Will check all stats - if one isn't high enough will exit
    For i = 1 To Stats.Stat_Count
        ' Make sure the stat has a requirement
        If Item(ItemNum).StatReq(i) > 0 Then
            If Current_BaseStat(Index, i) < Item(ItemNum).StatReq(i) Then
                SendActionMsg Current_Map(Index), "[" & StatName(i) & " req: " & Item(ItemNum).StatReq(i) & "]", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
                Exit Function
            End If
        End If
    Next
    
    ' If we get through all the checks it means we can use the item
    CanUseItem = True
End Function


'//////////////////
'// Calculations //
'//////////////////

'***************************************
' Calculate Exp to Next Level
'***************************************
Public Function Current_NextLevel(ByVal Index As Long) As Long
    Current_NextLevel = (Current_Level(Index) + 1) * (((Current_Level(Index) + 1) * 3) + 25) * 15
End Function

'**********************************************
' Calculates all the players modstats
'**********************************************
Public Sub Update_ModStats(ByVal Index As Long)
Dim i As Long
    For i = 1 To Stats.Stat_Count
        Update_ModStat Index, i
    Next
    SendStats Index
End Sub

'**************************************************
' Calculates your mod stat for a specific stat
'**************************************************
Public Sub Update_ModStat(ByVal Index As Long, ByVal Stat As Stats)
Dim i As Long
Dim ItemNum As Long, SpellNum As Long

    Player(Index).ModStat(Stat) = 0
    
    For i = 1 To Slots.Slot_Count
        ItemNum = Current_EquipmentSlot(Index, i)
        If ItemNum > 0 Then
            Player(Index).ModStat(Stat) = Player(Index).ModStat(Stat) + Item(ItemNum).ModStat(Stat)
        End If
    Next
    
    For i = 1 To MAX_STATUS
        SpellNum = Player(Index).Char.Status(i).SpellNum
        If SpellNum > 0 Then
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_BUFF
                    Player(Index).ModStat(Stat) = Player(Index).ModStat(Stat) + Spell(SpellNum).ModStat(Stat)
            End Select
        End If
    Next
End Sub

'**********************************************
' Gets your mod stat for a specific stat
'**********************************************
Public Function Current_ModStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Current_ModStat = Player(Index).ModStat(Stat)
End Function

'*******************************************
' Calculates your current stat (base + mod)
'*******************************************
Public Function Current_Stat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Current_Stat = Clamp(Current_BaseStat(Index, Stat) + Current_ModStat(Index, Stat), 0, MAX_LONG)
End Function

'************************************************
' Calculates your mod vitals
'************************************************
Public Sub Update_ModVitals(ByVal Index As Long)
Dim i As Long
    For i = 1 To Vitals.Vital_Count
        Update_ModVital Index, i
    Next
    ' Update our vitals now that our stats have changed
    Update_Vitals Index
End Sub

'************************************************
' Calculates your mod vital for a specific vital
'************************************************
Public Sub Update_ModVital(ByVal Index As Long, ByVal Vital As Vitals)
Dim i As Long
Dim ItemNum As Long, SpellNum As Long

    Player(Index).ModVital(Vital) = 0
            
    For i = 1 To Slots.Slot_Count
        ItemNum = Current_EquipmentSlot(Index, i)
        If ItemNum > 0 Then
            Player(Index).ModVital(Vital) = Player(Index).ModVital(Vital) + Item(ItemNum).ModVital(Vital)
        End If
    Next
    
    For i = 1 To MAX_STATUS
        SpellNum = Player(Index).Char.Status(i).SpellNum
        If SpellNum > 0 Then
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_BUFF
                    Player(Index).ModVital(Vital) = Player(Index).ModVital(Vital) + Spell(SpellNum).ModVital(Vital)
            End Select
        End If
    Next
End Sub

'************************************************
' Gets a specific modVital
'************************************************
Public Function Current_ModVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Current_ModVital = Player(Index).ModVital(Vital)
End Function

'***************************************
' Calculate Max Vital
'***************************************
Public Function Current_MaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

    Current_MaxVital = 0
    
    Select Case Vital
        Case HP
            Current_MaxVital = (((Current_Stat(Index, Stats.Vitality) \ 2) + (Current_Stat(Index, Stats.Strength) \ 5)) * Class(Current_Class(Index)).Vital(Vitals.HP)) + Current_Level(Index)
        Case MP
            Current_MaxVital = (((Current_Stat(Index, Stats.Wisdom) \ 2) + (Current_Stat(Index, Stats.Intelligence) \ 5)) * Class(Current_Class(Index)).Vital(Vitals.MP)) + Current_Level(Index)
        Case SP
            Current_MaxVital = (Current_Level(Index) + (Current_Stat(Index, Stats.Dexterity) \ 2) + Class(Current_Class(Index)).Stat(Stats.Dexterity)) * 2
    End Select
    
    Current_MaxVital = Clamp(Current_MaxVital + Current_ModVital(Index, Vital), 1, MAX_LONG)
End Function

'***************************************
' Calculate Vital Regen
' Based on Current Stats (Base + Mod)
'***************************************
Public Function Current_VitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
Dim i As Long

    Current_VitalRegen = 0
    
    Select Case Vital
        Case HP
            i = Current_Stat(Index, Stats.Vitality) \ 2
        Case MP
            i = Current_Stat(Index, Stats.Wisdom) \ 2
        Case SP
            i = Current_Stat(Index, Stats.Dexterity) \ 2
    End Select
    Current_VitalRegen = Clamp(i, 1, MAX_LONG)
End Function

'***************************************
' Calculate base damage
'***************************************
Public Function Current_Damage(ByVal Index As Long) As Long
    Current_Damage = Clamp((Current_Stat(Index, Stats.Strength) \ 2) + (Current_Stat(Index, Stats.Dexterity) \ 5.5), 0, MAX_LONG)
End Function

'***************************************
' Calculate base defense
'***************************************
Public Function Current_Protection(ByVal Index As Long) As Long
    Current_Protection = Clamp(Current_Stat(Index, Stats.Vitality) \ 2.5, 0, MAX_LONG)
End Function

'***************************************
' Calculate base magic damage
'***************************************
Public Function Current_MagicDamage(ByVal Index As Long) As Long
    Current_MagicDamage = Clamp(((Current_Stat(Index, Stats.Intelligence) \ 2) + (Current_Stat(Index, Stats.Wisdom) \ 5.5)) \ 2, 0, MAX_LONG)
End Function

'***************************************
' Calculate base magic defense
'***************************************
Public Function Current_MagicProtection(ByVal Index As Long) As Long
    Current_MagicProtection = Clamp(((Current_Stat(Index, Stats.Intelligence) \ 5.5) + (Current_Stat(Index, Stats.Wisdom) \ 5.5)) \ 2, 0, MAX_LONG)
End Function

'***************************************
' Determines if you can crit
'***************************************
Public Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    CanPlayerCriticalHit = False
    If Current_CritChance(Index) > Rand(0, 100) Then CanPlayerCriticalHit = True
End Function

'***************************************
' Players current crit chance
'***************************************
Public Function Current_CritChance(ByVal Index As Long) As Long
    Current_CritChance = (Current_Stat(Index, Stats.Dexterity) \ ((Current_Level(Index) * 0.2) + 2.5)) + Class(Current_Class(Index)).BaseCrit
End Function

'***************************************
' Determines if you will block a hit
'***************************************
Public Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    CanPlayerBlockHit = False
    If Current_EquipmentSlot(Index, Shield) > 0 Then
        If Current_BlockChance(Index) > Rand(0, 100) Then CanPlayerBlockHit = True
    End If
End Function

'***************************************
' Players current block chance
'***************************************
Public Function Current_BlockChance(ByVal Index As Long) As Long
    Current_BlockChance = (Current_Stat(Index, Stats.Vitality) \ ((Current_Level(Index) * 0.15) + 2.5)) + Class(Current_Class(Index)).BaseBlock
End Function

'***************************************
' Players current dodge chance
'***************************************
Public Function Current_DodgeChance(ByVal Index As Long) As Long
    Current_DodgeChance = (Current_Stat(Index, Stats.Dexterity) \ ((Current_Level(Index) * 0.15) + 2.5)) + Class(Current_Class(Index)).BaseDodge
End Function

'***************************************
' Players threat for the damage
'***************************************
Public Function DamageToThreat(ByVal Index As Long, ByVal Damage As Long) As Long
    DamageToThreat = Class(Current_Class(Index)).Threat * Abs(Damage)
End Function

'////////////
'// Events //
'////////////

'***************************************
' Events for updating
'***************************************
Sub OnUpdate(ByVal Index As Long)
Dim SpellInfo As Long       ' Used for spellnum and spellslot
Dim i As Long, n As Long
    
    '*****************************
    '**  Checks to save player  **
    '*****************************
    If GetTickCount > Player(Index).LastUpdateSave Then
        SavePlayer Index
        SendActionMsg Current_Map(Index), "Progress Auto-Saved.", White, ACTIONMSG_SCREEN, 0, 0, Index
        Player(Index).LastUpdateSave = GetTickCount + 600000    ' 10 minutes
    End If
        
    '*************************************
    '**  Checks cooldown on all spells  **
    '*************************************
    For i = 1 To MAX_PLAYER_SPELLS
        If Current_SpellCooldown(Index, i) > 0 Then
            If GetTickCount >= Current_SpellCooldown(Index, i) Then
                Update_SpellCooldown Index, i, 0
                SendSpellReady Index, i
            End If
        End If
    Next
    
    '**************
    '**  Living  **
    '**************
    If Not Current_IsDead(Index) Then
        '**************************************
        '**  Checks to update player vitals  **
        '**************************************
        If GetTickCount > Player(Index).LastUpdateVitals Then
            For n = 1 To Vitals.Vital_Count
                If Current_BaseVital(Index, n) < Current_MaxVital(Index, n) Then
                    Update_BaseVital Index, n, Current_BaseVital(Index, n) + Current_VitalRegen(Index, n)
                End If
            Next
            Player(Index).LastUpdateVitals = GetTickCount + 10000   ' 10 seconds
        End If
        
        '********************************************
        '**  Checks casting time on casting spell  **
        '********************************************
        SpellInfo = Player(Index).CastingSpell
        If SpellInfo > 0 Then
            If GetTickCount >= Player(Index).CastTime Then
                OnCastSpell Index, SpellInfo
            End If
        End If
        
        '*********************************
        '**  Checks for status effects  **
        '*********************************
        For i = 1 To MAX_STATUS
            SpellInfo = Player(Index).Char.Status(i).SpellNum
            If SpellInfo > 0 Then
                ' Check if it's time to do something for the spell
                If GetTickCount >= Player(Index).Char.Status(i).TickUpdate Then
                    ' if we have an overtime spell - let's do our damage/heal
                    Select Case Spell(SpellInfo).Type
                        Case SPELL_TYPE_OVERTIME
                            For n = 1 To Vitals.Vital_Count
                                Update_BaseVital Index, n, Current_BaseVital(Index, n) + Spell(SpellInfo).ModVital(n)
                                If Spell(SpellInfo).ModVital(n) > 0 Then
                                    SendActionMsg Current_Map(Index), "+" & CStr(Spell(SpellInfo).ModVital(n)) & " " & VitalName(n), BrightGreen, ACTIONMSG_SCROLL, Current_X(Index), Current_Y(Index)
                                ElseIf Spell(SpellInfo).ModVital(n) < 0 Then
                                    SendActionMsg Current_Map(Index), CStr(Spell(SpellInfo).ModVital(n)) & " " & VitalName(n), Yellow, ACTIONMSG_SCROLL, Current_X(Index), Current_Y(Index)
                                End If
                            Next
                            
                            ' check if it killed them?
                            If Current_BaseVital(Index, Vitals.HP) <= 0 Then
                                OnDeath Index
                                Exit Sub
                            End If
                    End Select
                    
                    ' subtract one from the count
                    Player(Index).Char.Status(i).TickCount = Player(Index).Char.Status(i).TickCount - 1
                    
                    ' check if it's over?
                    If Player(Index).Char.Status(i).TickCount <= 0 Then
                        Player(Index).Char.Status(i).SpellNum = 0
                        Player(Index).Char.Status(i).TickCount = 0
                        Player(Index).Char.Status(i).TickUpdate = 0
                        
                        ' Update mods
                        Update_ModStats Index
                        Update_ModVitals Index
                        
                    ' if not set the next tick
                    Else
                        Player(Index).Char.Status(i).TickUpdate = GetTickCount + (Spell(SpellInfo).TickUpdate * 1000)
                    End If
                End If
            End If
        Next
    
    '************
    '**  Dead  **
    '************
    Else
        '********************************
        '**  Checks to release player  **
        '********************************
        If GetTickCount > Current_IsDeadTimer(Index) Then
            SendPlayerMsg Index, "You have been auto-released.", BrightRed
            OnRelease Index
        End If
    End If
End Sub

'***************************************
' Events for leveling up
'***************************************
Sub OnLevelUp(ByVal Index As Long)
Dim expRollover As Long
    
    expRollover = (Current_Exp(Index) - Current_NextLevel(Index))
    
    ' Actually update the level
    Update_Level Index, Current_Level(Index) + 1
        
    ' Set the ammount of skill Points to add
    Update_Points Index, Current_Points(Index) + STATS_PER_LEVEL
    SendStats Index
    
    ' Send a message
    SendActionMsg Current_Map(Index), "Level Up!", Yellow, ACTIONMSG_SCROLL, Current_X(Index), Current_Y(Index)
    SendActionMsg Current_Map(Index), "You are now level " & Current_Level(Index) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
    SendGlobalMsg "[Advance] " & Current_Name(Index) & " has gained a level!", ActionColor
    SendPlayerMsg Index, "[Advance] You have gained a level! You now have " & Current_Points(Index) & " skill Points!", AlertColor
    
    ' Check if we now are at max level so we can set the exp rollover to 1
    If Current_Level(Index) = MAX_LEVEL Then expRollover = 0

    ' Update with the rollover exp
    Update_Exp Index, expRollover
End Sub

'***************************************
' Events for death
'***************************************
Sub OnDeath(ByVal Index As Long, Optional ByVal Msg As String = "You have been slain.")
Dim i As Long

    ' Set hp to 0
    Update_BaseVital Index, Vitals.HP, 0
    
    ' Set revivable to 0
    Player(Index).Revivable = 0
    
    ' Set the isdead flag
    Update_IsDead Index, True
    
    ' Set the isdead timer
    Update_IsDeadTimer Index, GetTickCount + 1800000  ' 30 minutes
    
    ' Target check
    ChangeTarget Index, 0, TARGET_TYPE_NONE
    
    ' Check if they were casting
    CheckCasting Index
    
    ' Clear out all status effects
    ClearStatusEffects Index
    
    ' If the player was a pk then take it away
    If Current_PK(Index) Then
        Update_PK Index, 0
        SendPlayerData Index
    End If
    
    SendPlayerDead Index
    
    ' Player is dead
    SendActionMsg Current_Map(Index), "You have been slain.", BrightRed, ACTIONMSG_SCREEN, 0, 0, Index
End Sub

'***************************************
' Events for being revived
' Revives in the current position
' Only lose 5% exp
'***************************************
Sub OnRevive(ByVal Index As Long)
Dim i As Long
Dim expLoss As Long

    ' Revive in current position
    PlayerWarp Index, Current_Position(Index)
    
    ' Update mods
    Update_ModStats Index
    Update_ModVitals Index
    
    ' Restore vitals based on the revive spell cast upon the player
    For i = 1 To Vitals.Vital_Count
        Update_BaseVital Index, i, Int(Current_MaxVital(Index, i) * (Spell(Player(Index).Revivable).ModVital(i) * 0.01))
    Next
    
    ' Lose 5% of exp
    expLoss = Int(Current_Exp(Index) * 0.05)
    Update_Exp Index, Current_Exp(Index) - expLoss
    SendPlayerMsg Index, "You have lost " & expLoss & " experience points.", BrightRed
    
    ' Update isdead
    Update_IsDead Index, False
    
     ' Set revivable to 0
    Player(Index).Revivable = 0
End Sub

'***************************************
' Events for releasing from your body
' Revives at boot map or bound map
' Loses 10% exp
'***************************************
Sub OnRelease(ByVal Index As Long)
Dim i As Long
Dim expLoss As Long
Dim NewPosition As PositionRec
        
    ' If there is a boot map - warp there - otherwise go to regular spot
    If Map(Current_Map(Index)).BootMap > 0 Then
        NewPosition.Map = Map(Current_Map(Index)).BootMap
        NewPosition.X = Map(Current_Map(Index)).BootX
        NewPosition.Y = Map(Current_Map(Index)).BootY
        PlayerWarp Index, NewPosition
    Else
        PlayerWarp Index, Current_Bound(Index)
    End If
    
    ' Update mods
    Update_ModStats Index
    Update_ModVitals Index
    
    ' Restore vitals
    For i = 1 To Vitals.Vital_Count
        ' 25% of max
        Update_BaseVital Index, i, Int(Current_MaxVital(Index, i) * 0.25)
    Next
    
    ' Lose 10% of exp
    expLoss = Int(Current_Exp(Index) * 0.1)
    Update_Exp Index, Current_Exp(Index) - expLoss
    SendPlayerMsg Index, "You have lost " & expLoss & " experience points.", BrightRed
    
     ' Update isdead
    Update_IsDead Index, False
    
    ' Set revivable to 0
    Player(Index).Revivable = 0
End Sub

'***************************************
' Events for using an item
'***************************************
Sub OnUseItem(ByVal Index As Long, ByVal InvNum As Long)
Dim ItemNum As Long
Dim n As Long, X As Long, Y As Long, i As Long

    ' Can't do while dead
    If Current_IsDead(Index) Then Exit Sub
    
    ItemNum = Current_InvItemNum(Index, InvNum)
    
    ' Do a check to see if it's a valid itemnum
    If ItemNum <= 0 Then Exit Sub
    If ItemNum > MAX_ITEMS Then Exit Sub
        
    ' If we can't use the item, exit
    If Not CanUseItem(Index, ItemNum) Then Exit Sub
            
    ' Find out what kind of item it is
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_EQUIPMENT
            OnEquipSlot Index, InvNum, ItemNum, Item(ItemNum).Data1
            
        Case ITEM_TYPE_POTION
            For i = 1 To Vitals.Vital_Count
                If Item(ItemNum).ModVital(i) <> 0 Then
                    Update_BaseVital Index, i, Current_BaseVital(Index, i) + Item(ItemNum).ModVital(i)
                    If Item(ItemNum).ModVital(i) > 0 Then
                        SendActionMsg Current_Map(Index), "+" & CStr(Item(ItemNum).ModVital(i)) & " " & VitalName(i), Yellow, ACTIONMSG_SCROLL, Current_X(Index), Current_Y(Index)
                    Else
                        SendActionMsg Current_Map(Index), CStr(Item(ItemNum).ModVital(i)) & " " & VitalName(i), Yellow, ACTIONMSG_SCROLL, Current_X(Index), Current_Y(Index)
                    End If
                End If
            Next
            TakeInventoryItem Index, InvNum, 1
            SendPlayerInvUpdate Index, InvNum
            
        Case ITEM_TYPE_KEY
            Select Case Current_Dir(Index)
                Case DIR_UP
                    If Current_Y(Index) > 0 Then
                        X = Current_X(Index)
                        Y = Current_Y(Index) - 1
                    Else
                        Exit Sub
                    End If
                    
                Case DIR_DOWN
                    If Current_Y(Index) < Map(Current_Map(Index)).MaxY Then
                        X = Current_X(Index)
                        Y = Current_Y(Index) + 1
                    Else
                        Exit Sub
                    End If
                        
                Case DIR_LEFT
                    If Current_X(Index) > 0 Then
                        X = Current_X(Index) - 1
                        Y = Current_Y(Index)
                    Else
                        Exit Sub
                    End If
                        
                Case DIR_RIGHT
                    If Current_X(Index) < Map(Current_Map(Index)).MaxX Then
                        X = Current_X(Index) + 1
                        Y = Current_Y(Index)
                    Else
                        Exit Sub
                    End If
            End Select
            
            ' Check if a key exists
            If Map(Current_Map(Index)).Tile(X, Y).Type = TILE_TYPE_KEY Then
                ' Check if the key they are using matches the map key
                If ItemNum = Map(Current_Map(Index)).Tile(X, Y).Data1 Then
                    ' Check if the door is open
                    If Not MapData(Current_Map(Index)).TempTile.DoorOpen(X, Y) Then
                        MapData(Current_Map(Index)).TempTile.DoorOpen(X, Y) = True
                        MapData(Current_Map(Index)).TempTile.DoorTimer(X, Y) = GetTickCount + 5000
                        
                        SendMapKey Current_Map(Index), X, Y, 1
                        SendMapMsg Current_Map(Index), "A door has been unlocked.", ActionColor
                        
                        ' Check if we are supposed to take away the item
                        If Map(Current_Map(Index)).Tile(X, Y).Data2 = 1 Then
                            'TakeItem(Index, ItemNum, 1)
                            TakeInventoryItem Index, InvNum, 1
                            SendPlayerInvUpdate Index, InvNum
                            SendActionMsg Current_Map(Index), "The key vanishes.", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
                        End If
                    End If
                End If
            End If
            
        Case ITEM_TYPE_SPELL
            ' Get the spell num
            n = Item(ItemNum).Data1
            If n > 0 Then
                ' Check for class requirement
                ' Will check your current class to the item
                ' Checks the binary flag is set for your class
                If Not Spell(n).ClassReq And (2 ^ Current_Class(Index)) Then
                    SendActionMsg Current_Map(Index), "[Your class can not use this spell.]", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
                    Exit Sub
                End If
                
                ' Check for level requirement
                If Spell(n).LevelReq > 0 Then
                    ' If there's a level requirement then check if you can use it
                    ' Checks if your level is below the req and if so - will exit
                    If Current_Level(Index) < Spell(n).LevelReq Then
                        SendActionMsg Current_Map(Index), "[Level Req: " & Spell(n).LevelReq & "]", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
                        Exit Sub
                    End If
                End If
                    
                i = FindOpenSpellSlot(Index)
                ' Make sure they have an open spell slot
                If i > 0 Then
                    ' Make sure they dont already have the spell
                    If Not HasSpell(Index, n) Then
                        Update_Spell Index, i, n
                        TakeInventoryItem Index, InvNum, 1
                        SendPlayerInvUpdate Index, InvNum
                        SendActionMsg Current_Map(Index), "You have learnt a new spell.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
                    Else
                        SendActionMsg Current_Map(Index), "You already know this spell.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
                    End If
                Else
                    SendActionMsg Current_Map(Index), "You cannot learn more spells.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
                End If
            Else
                SendActionMsg Current_Map(Index), "This item is bugged, contact a Realm Master.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
            End If
    End Select
End Sub

'***************************************
' Events for casting spells
'***************************************
Sub OnCastSpell(ByVal Index As Long, ByVal SpellSlot As Long)
Dim SpellNum As Long
Dim i As Long
Dim StatusSlot As Long
Dim MapNum As Long
Dim Target As Long
Dim Damage As Long

    SpellNum = Current_Spell(Index, SpellSlot)
    MapNum = Current_Map(Index)
    Target = Player(Index).CastTarget
    
    ' Check to make sure we haven't died since the begining of the cast
    If Current_IsDead(Index) Then
        CancelCastSpell Index
    End If
    
    ' Check if we still have enough vitals to cast
    For i = 1 To Vitals.Vital_Count
        If Spell(SpellNum).VitalReq(i) > Current_BaseVital(Index, i) Then
            SendPlayerMsg Index, "Spell fizzled out. " & CStr(Spell(SpellNum).VitalReq(i)) & " " & VitalName(i) & " required.", BrightRed
            CancelCastSpell Index
            Exit Sub
        End If
    Next
    
    ' Check if timer is ok
    If GetTickCount < Player(Index).AttackTimer + 1000 Then Exit Sub

    ' Player to Player spells
    If Player(Index).CastTargetType = TARGET_TYPE_PLAYER Then
        ' Check if player is still on the same map
        If Current_Map(Target) <> MapNum Then
            SendPlayerMsg Index, "Spell fizzled out. Target no longer valid.", BrightRed
            CancelCastSpell Index
            Exit Sub
        End If
        
        ' Check if they are still in range
        If Not PlayerInRange(Index, Current_X(Target), Current_Y(Target), Spell(SpellNum).Range) Then
            SendPlayerMsg Index, "Spell fizzled out. Target no longer in range.", BrightRed
            CancelCastSpell Index
            Exit Sub
        End If
    
        ' Our different spell types
        Select Case Spell(SpellNum).Type
            Case SPELL_TYPE_VITAL
                ' Check again if they have died since the prespell check
                If Current_IsDead(Target) Then
                    SendPlayerMsg Index, "Spell fizzled out. Target is dead.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
        
                For i = 1 To Vitals.Vital_Count
                    If Spell(SpellNum).ModVital(i) <> 0 Then
                        If Spell(SpellNum).ModVital(i) > 0 Then
                            Update_BaseVital Target, i, Current_BaseVital(Target, i) + (Spell(SpellNum).ModVital(i) + Current_MagicDamage(Index))
                            SendActionMsg MapNum, "+" & CStr(Spell(SpellNum).ModVital(i) + Current_MagicDamage(Index)) & " " & VitalName(i), BrightGreen, ACTIONMSG_SCROLL, Current_X(Target), Current_Y(Target)
                        ElseIf Spell(SpellNum).ModVital(i) < 0 Then
                            Update_BaseVital Target, i, Current_BaseVital(Target, i) + (Spell(SpellNum).ModVital(i) - Current_MagicDamage(Index))
                            SendActionMsg MapNum, CStr(Spell(SpellNum).ModVital(i) - Current_MagicDamage(Index)) & " " & VitalName(i), Yellow, ACTIONMSG_SCROLL, Current_X(Target), Current_Y(Target)
                        End If
                    End If
                Next
                
                ' Send animation
                If Spell(SpellNum).Animation Then SendAnimation MapNum, Spell(SpellNum).Animation, Current_X(Target), Current_Y(Target)
                
                ' Check if it kills players
                If Current_BaseVital(Target, Vitals.HP) <= 0 Then
                    SendActionMsg Current_Map(Index), "You have slain " & Current_Name(Target) & ".", BrightRed, ACTIONMSG_SCREEN, 0, 0, Index
                    
                    i = Clamp((Current_Exp(Target) \ 10) * (ExpMod * 0.01), 0, MAX_LONG)
                    If i > 0 Then
                        SendActionMsg Current_Map(Index), "+" & i & " EXP!", Yellow, ACTIONMSG_SCROLL, Current_X(Index), Current_Y(Index), Index
                        SendActionMsg Current_Map(Target), "-" & i & " EXP!", Yellow, ACTIONMSG_SCROLL, Current_X(Target), Current_Y(Index), Target
                        Update_Exp Index, Current_Exp(Index) + i
                        Update_Exp Target, Current_Exp(Target) - 1
                    End If
                    
                    If Current_PK(Target) = 0 Then
                        If Current_PK(Index) = 0 Then
                            Update_PK Index, 1
                            SendPlayerData Index
                        End If
                    Else
                        Update_PK Target, 0
                        SendPlayerData Target
                    End If
                    
                    OnDeath Target, "You have been slain by " & Current_Name(Index) & "."
                End If
                
            Case SPELL_TYPE_OVERTIME
                ' Check again if they have died since the prespell check
                If Current_IsDead(Target) Then
                    SendPlayerMsg Index, "Spell fizzled out. Target is dead.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                ' Loop our current status effects, check if you have the spell already
                ' if so reapply, if not add it
                For i = 1 To MAX_STATUS
                    ' check if we have a status slot
                    ' this will set our statusslot to the first open slot
                    If StatusSlot = 0 Then
                        If Player(Target).Char.Status(i).SpellNum = 0 Then
                            StatusSlot = i
                        End If
                    End If
                    
                    ' checks if we have the spell already
                    If Player(Target).Char.Status(i).SpellNum = SpellNum Then
                        StatusSlot = i
                        Exit For
                    End If
                Next
                
                ' If no available slot, we'll just fizzle for now
                ' TODO: Think about what to do?
                If StatusSlot = 0 Then
                    SendPlayerMsg Index, "Spell fizzled out.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                ' Apply the overtime spell
                Player(Target).Char.Status(StatusSlot).SpellNum = SpellNum
                Player(Target).Char.Status(StatusSlot).TickCount = Spell(SpellNum).TickCount 'Set the base number so we can subtract from this
                Player(Target).Char.Status(StatusSlot).TickUpdate = GetTickCount + (Spell(SpellNum).TickUpdate * 1000)   ' Set the first tick

                ' Send animation
                If Spell(SpellNum).Animation Then SendAnimation MapNum, Spell(SpellNum).Animation, Current_X(Target), Current_Y(Target)
                
                ' Update mods of target
                Update_ModStats Target
                Update_ModVitals Target
                
            Case SPELL_TYPE_BUFF
                ' Check again if they have died since the prespell check
                If Current_IsDead(Target) Then
                    SendPlayerMsg Index, "Spell fizzled out. Target is dead.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                ' Loop our current status check if you have the spell already - if so reapply, if not
                For i = 1 To MAX_STATUS
                    ' check if we have a status slot
                    ' this will set our statusslot to the first open slot
                    If StatusSlot = 0 Then
                        If Player(Target).Char.Status(i).SpellNum = 0 Then
                            StatusSlot = i
                        End If
                    End If
                    
                    ' checks if we have the spell already
                    If Player(Target).Char.Status(i).SpellNum = SpellNum Then
                        StatusSlot = i
                        Exit For
                    End If
                Next
                
                ' If no available slot we'll just fizzle for now
                ' TODO: Think about what to do?
                If StatusSlot = 0 Then
                    SendPlayerMsg Index, "Spell fizzled out.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                ' Apply the buff
                Player(Target).Char.Status(StatusSlot).SpellNum = SpellNum
                Player(Target).Char.Status(StatusSlot).TickCount = Spell(SpellNum).TickCount 'Set the base number so we can subtract from this
                Player(Target).Char.Status(StatusSlot).TickUpdate = GetTickCount + (Spell(SpellNum).TickUpdate * 1000)   ' Set the first tick
                
                ' Send animation
                If Spell(SpellNum).Animation Then SendAnimation MapNum, Spell(SpellNum).Animation, Current_X(Target), Current_Y(Target)
                
                ' Update mods of target
                Update_ModStats Target
                Update_ModVitals Target
            
            Case SPELL_TYPE_REVIVE
                ' Check again if they have revived since the prespell check
                If Not Current_IsDead(Target) Then
                    SendPlayerMsg Index, "Spell fizzled out. Target not dead.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                'Player(Target).Revivable = 1
                Player(Target).Revivable = SpellNum
                
                ' Send accept revival
                SendPlayerRevival Target, Current_Name(Index)

        End Select
    ' Player to NPC spells
    ElseIf Player(Index).CastTargetType = TARGET_TYPE_NPC Then
        ' Make sure your target is still good
        If MapData(MapNum).MapNpc(Target).Num <= 0 Then
            SendPlayerMsg Index, "Spell fizzled out. Target no longer valid.", BrightRed
            CancelCastSpell Index
            Exit Sub
        End If
        
        ' Check if they are still in range
        If Not NpcInRange(Index, Target, Spell(SpellNum).Range) Then
            SendPlayerMsg Index, "Spell fizzled out. Target no longer in range.", BrightRed
            CancelCastSpell Index
            Exit Sub
        End If
        
        Select Case Spell(SpellNum).Type
            Case SPELL_TYPE_VITAL
                For i = 1 To Vitals.Vital_Count
                    ' Positive = Healing
                    If Spell(SpellNum).ModVital(i) > 0 Then
                        Damage = (Spell(SpellNum).ModVital(i) + Current_MagicDamage(Index))
                        Damage = Rand(Damage * 0.9, Damage * 1.1)
                        MapNpc_Update_Vital MapNum, Target, i, MapNpc_Current_Vital(MapNum, Target, i) + Damage
                        SendActionMsg MapNum, "+" & CStr(Damage) & " " & VitalName(i), BrightGreen, ACTIONMSG_SCROLL, MapData(MapNum).MapNpc(Target).X, MapData(MapNum).MapNpc(Target).Y
                        ' Check if it's above
                        If MapNpc_Current_Vital(MapNum, Target, i) > MapNpc_MaxVital(MapNum, Target, i) Then MapNpc_Update_Vital MapNum, Target, i, MapNpc_MaxVital(MapNum, Target, i)
                    ' Negative = Damage
                    ElseIf Spell(SpellNum).ModVital(i) < 0 Then
                        Damage = Abs((Spell(SpellNum).ModVital(i) - Current_MagicDamage(Index)))
                        Damage = Rand(Damage * 0.9, Damage * 1.1)
                        ' You can't do more damage than the npcs max vital
                        If Damage > MapNpc_Current_Vital(MapNum, Target, i) Then Damage = MapNpc_Current_Vital(MapNum, Target, i)
                        MapNpc_Update_Vital MapNum, Target, i, MapNpc_Current_Vital(MapNum, Target, i) - Damage
                        SendActionMsg MapNum, CStr(-Damage) & " " & VitalName(i), Yellow, ACTIONMSG_SCROLL, MapData(MapNum).MapNpc(Target).X, MapData(MapNum).MapNpc(Target).Y
                        ' Adds the damage and checks it target
                        MapNpc_AddDamage MapNum, Target, Index, Damage
                    End If
                Next
                
                ' Send animation
                If Spell(SpellNum).Animation Then SendAnimation MapNum, Spell(SpellNum).Animation, MapData(MapNum).MapNpc(Target).X, MapData(MapNum).MapNpc(Target).Y
                
                ' check if it kills npc
                If MapNpc_Current_Vital(MapNum, Target, Vitals.HP) <= 0 Then MapNpc_OnDeath MapNum, Target
                
            Case SPELL_TYPE_OVERTIME
                ' loop our current status check if you have the spell already - if so reapply, if not
                For i = 1 To MAX_STATUS
                    ' check if we have a status slot
                    ' this will set our statusslot to the first open slot
                    If StatusSlot = 0 Then
                        If MapData(MapNum).MapNpc(Target).Status(i).SpellNum = 0 Then
                            StatusSlot = i
                        End If
                    End If
                    
                    ' checks if we have the spell already
                    If MapData(MapNum).MapNpc(Target).Status(i).SpellNum = SpellNum Then
                        StatusSlot = i
                        Exit For
                    End If
                Next
                
                ' If no available slot we'll just fizzle for now
                ' TODO: Think about what to do?
                If StatusSlot = 0 Then
                    SendPlayerMsg Index, "Spell fizzled out.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
    
                ' Apply the overtime spell
                MapData(MapNum).MapNpc(Target).Status(StatusSlot).SpellNum = SpellNum
                MapData(MapNum).MapNpc(Target).Status(StatusSlot).TickCount = Spell(SpellNum).TickCount 'Set the base number so we can subtract from this
                MapData(MapNum).MapNpc(Target).Status(StatusSlot).TickUpdate = GetTickCount + (Spell(SpellNum).TickUpdate * 1000)   ' Set the first tick
                MapData(MapNum).MapNpc(Target).Status(StatusSlot).Caster = Current_Name(Index)
                
                ' Send animation
                If Spell(SpellNum).Animation Then SendAnimation MapNum, Spell(SpellNum).Animation, MapData(MapNum).MapNpc(Target).X, MapData(MapNum).MapNpc(Target).Y
                
            Case SPELL_TYPE_BUFF
                ' loop our current status check if you have the spell already - if so reapply, if not
                For i = 1 To MAX_STATUS
                    ' check if we have a status slot
                    ' this will set our statusslot to the first open slot
                    If StatusSlot = 0 Then
                        If MapData(MapNum).MapNpc(Target).Status(i).SpellNum = 0 Then
                            StatusSlot = i
                        End If
                    End If
                    
                    ' checks if we have the spell already
                    If MapData(MapNum).MapNpc(Target).Status(i).SpellNum = SpellNum Then
                        StatusSlot = i
                        Exit For
                    End If
                Next
                
                ' If no available slot we'll just fizzle for now
                ' TODO: Think about what to do?
                If StatusSlot = 0 Then
                    SendPlayerMsg Index, "Spell fizzled out.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
    
                ' Apply the buff
                MapData(MapNum).MapNpc(Target).Status(StatusSlot).SpellNum = SpellNum
                MapData(MapNum).MapNpc(Target).Status(StatusSlot).TickCount = Spell(SpellNum).TickCount 'Set the base number so we can subtract from this
                MapData(MapNum).MapNpc(Target).Status(StatusSlot).TickUpdate = GetTickCount + (Spell(SpellNum).TickUpdate * 1000)   ' Set the first tick
                MapData(MapNum).MapNpc(Target).Status(StatusSlot).Caster = Current_Name(Index)
                
                ' Send animation
                If Spell(SpellNum).Animation Then SendAnimation MapNum, Spell(SpellNum).Animation, MapData(MapNum).MapNpc(Target).X, MapData(MapNum).MapNpc(Target).Y
        End Select
    End If

    ' Now actually take the vitals needed since we casted
    For i = 1 To Vitals.Vital_Count
        Update_BaseVital Index, i, Current_BaseVital(Index, i) - Spell(SpellNum).VitalReq(i)
    Next
    
    ' Clears it out?
    ' TODO:
    CancelCastSpell Index
    
    ' set the cooldown
    Update_SpellCooldown Index, SpellSlot, GetTickCount + (Spell(SpellNum).Cooldown * 1000)
    
    ' send the cooldown
    SendSpellCooldown Index, SpellSlot
End Sub

'***************************************
' Events for equiping equipment
'***************************************
Sub OnEquipSlot(ByVal Index As Long, ByVal InvNum As Long, ByVal ItemNum As Long, ByVal EquipmentSlot As Slots)
Dim n As Long

    ' Can't do while dead
    If Current_IsDead(Index) Then Exit Sub
    
    If EquipmentSlot <= 0 Then Exit Sub
    If EquipmentSlot > Slots.Slot_Count Then Exit Sub
    
    SendActionMsg Current_Map(Index), "You equipped " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
    
    n = Current_EquipmentSlot(Index, EquipmentSlot)
    If n > 0 Then
        ' Set the equipment slot
        Update_EquipmentSlot Index, EquipmentSlot, ItemNum
        ' Set the item to invnum
        Update_InvItemNum Index, InvNum, n
        Update_InvItemValue Index, InvNum, 1
        
        ' Check if needs to be bound
        If Item(ItemNum).Bound Then
            Update_InvItemBound Index, InvNum, True
        End If
    Else
        ' Set itemnum to the equipmentslot
        Update_EquipmentSlot Index, EquipmentSlot, ItemNum
        ' Delete the invitem
        TakeInventoryItem Index, InvNum, 1
    End If
    
    SendPlayerInvUpdate Index, InvNum
End Sub

'***************************************
' Events for unequiping equipment
' Will either set in inv or ground
'***************************************
Sub OnUnequipSlot(ByVal Index As Long, ByVal EquipmentSlot As Slots)
Dim n As Long, ItemNum As Long
    
    ' Can't do while dead
    If Current_IsDead(Index) Then Exit Sub
    
    If EquipmentSlot <= 0 Then Exit Sub
    If EquipmentSlot > Slots.Slot_Count Then Exit Sub
    
    ItemNum = Current_EquipmentSlot(Index, EquipmentSlot)
    If ItemNum > 0 Then
        
        n = FindOpenInvSlot(Index, ItemNum)
        If n <> 0 Then
            ' Set the armor slot to the inv num
            Update_InvItemNum Index, n, ItemNum
            Update_InvItemValue Index, n, 1
            
            ' Check if needs to be bound
            If Item(ItemNum).Bound Then
                Update_InvItemBound Index, n, True
            End If
            
            ' Clear the equipment
            Update_EquipmentSlot Index, EquipmentSlot, 0
            ' Send updated inventory
            SendPlayerInvUpdate Index, n
        Else
            ' Since we don't have room in the inventory, tell the player
            SendActionMsg Current_Map(Index), "You can not unequip at this time.", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
        End If
        
        SendActionMsg Current_Map(Index), "You unequiped " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
    End If
End Sub

'***************************************
' Events for when vitals can change
'***************************************
Sub Update_Vitals(ByVal Index As Long)
Dim i As Long

    For i = 1 To Vitals.Vital_Count
        ' If you have more vital than the max, lower it to the max
        ' If it's not, just send the vital
        If Current_BaseVital(Index, i) > Current_MaxVital(Index, i) Then
            Update_BaseVital Index, i, Current_MaxVital(Index, i)
        Else
            SendVital Index, i
        End If
    Next
End Sub

'*************************************************
' Used for things that will cancel spell casting
'*************************************************
Public Sub CheckCasting(ByVal Index As Long)
    If Player(Index).CastingSpell Then CancelCastSpell Index
End Sub

'***************************************
' Cancels spell casting
'***************************************
Public Sub CancelCastSpell(ByVal Index As Long)
    Player(Index).CastingSpell = 0
    Player(Index).CastTime = 0
    Player(Index).CastTarget = 0
    Player(Index).CastTargetType = TARGET_TYPE_NONE
    SendCancelSpell Index
End Sub

'***************************************
' Changes the players current target
'***************************************
Public Sub ChangeTarget(ByVal Index As Long, ByVal Target As Long, ByVal TargetType As Long)
    Player(Index).Target = Target
    Player(Index).TargetType = TargetType
    SendTarget Index
End Sub

'***************************************
' Checks specified range
'***************************************
Function PlayerInRange(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Distance As Byte) As Boolean
Dim DistanceX As Long, DistanceY As Long

    PlayerInRange = False
    
    If Not IsPlaying(Index) Then Exit Function
    
    DistanceX = X - Current_X(Index)
    DistanceY = Y - Current_Y(Index)
    
    ' Make sure we get a positive value
    If DistanceX < 0 Then DistanceX = -DistanceX
    If DistanceY < 0 Then DistanceY = -DistanceY
    
    ' Are they in range?
    If DistanceX <= Distance Then
        If DistanceY <= Distance Then
            PlayerInRange = True
        End If
    End If
End Function

Function PlayerData(ByVal Index As Long) As Byte()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString Current_Name(Index)
    Buffer.WriteByte Current_Class(Index)
    Buffer.WriteLong Current_Sprite(Index)
    Buffer.WriteLong Current_Map(Index)
    Buffer.WriteLong Current_X(Index)
    Buffer.WriteLong Current_Y(Index)
    Buffer.WriteLong Current_Dir(Index)
    Buffer.WriteLong Current_Access(Index)
    Buffer.WriteLong Current_PK(Index)
    Buffer.WriteString Current_GuildName(Index)
    Buffer.WriteString Current_GuildAbbreviation(Index)
    Buffer.WriteByte Current_IsDead(Index)
    Buffer.WriteLong Current_IsDeadTimer(Index) - GetTickCount
    
    PlayerData = Buffer.ToArray()
End Function
