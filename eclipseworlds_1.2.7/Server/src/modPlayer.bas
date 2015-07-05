Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal index As Long)
    Call JoinGame(index)
    Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Name & ".", "Player")
    Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Name & ".")
    Call UpdateCaption
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim Color As Long

    ' Set the flag so we know the person is in the game
    tempplayer(index).InGame = True

    ' Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
    
    ' Send an ok to client to start receiving in game data
    Call SendLogin(index)

    TotalPlayersOnline = TotalPlayersOnline + 1
        
    ' Send data
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNPCs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call CheckEquippedItems(index)
    Call SendHotbar(index)
    Call SendTitles(index)
    Call SendMorals(index)
    Call SendEmoticons(index)
    Call SendQuests(index)
    
    For i = 1 To Vital_Count - 1
        Call SendVital(index, i)
    Next
    
    Call SendPlayerExp(index)
    
    ' Spell Cooldowns
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) > 0 Then
            ' Check if the CD has expired
            If GetPlayerSpellCD(index, i) - timeGetTime < 1 Then Call SetPlayerSpellCD(index, i, 0)
            If GetPlayerSpellCD(index, i) - timeGetTime >= Spell(GetPlayerSpell(index, i)).CDTime * 1000 Then Call SetPlayerSpellCD(index, i, 0)
            If GetPlayerSpellCD(index, i) <= timeGetTime Then Call SetPlayerSpellCD(index, i, 0)
            
            ' Send it
            Call SendSpellCooldown(index, i)
        End If
    Next
    
    ' Check for glitches in the inventory
    Call UpdatePlayerItems(index)
    
    ' Check for glitches in equipment
    Call UpdatePlayerEquipmentItems(index)
    
    ' Check for new title
    Call CheckPlayerNewTitle(index)
    
    ' Send the player's data
    Call SetPlayerQuestData(index)

    ' Warp the player to their saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), True)
    
    ' Send welcome messages
    Call SendWelcome(index)
    
    Call UpdateClassData(index)
    
    ' Send a global message that they joined
    If GetPlayerAccess(index) <= STAFF_MODERATOR Then
        If Class(GetPlayerClass(index)).Color = Orange Then
            Color = RGB(255, 165, 0)
        Else
            Color = Class(GetPlayerClass(index)).Color
        End If
        
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Name & "!", Color)
    Else
         ' Color for access
        Select Case GetPlayerAccess(index)
            Case 0
                Color = 15
            Case 1
                Color = 3
            Case 2
                Color = 2
            Case 3
                Color = BrightBlue
            Case 4
                Color = Yellow
            Case 5
                Color = RGB(255, 165, 0)
        End Select
            
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Name & "!", Color)
    End If

    ' Send the flag so they know they can start doing stuff
    Call SendInGame(index)
    
    ' Refresh the friends list to all players online
    For i = 1 To Player_HighIndex
        Call UpdateFriendsList(i)
    Next
    
    ' Refresh the foes list to all players online
    For i = 1 To Player_HighIndex
        Call UpdateFoesList(i)
    Next
    
    ' Update guild list
    If GetPlayerGuild(index) > 0 Then
        Call SendPlayerGuildMembers(index)
    End If
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, i As Long
    Dim TradeTarget As Long

    If index < 1 Or index > Player_HighIndex Then Exit Sub
 
    If Not tempplayer(index).HasLogged And GetPlayerLogin(index) <> vbNullString Then
        Call CloseSocket(index)
    End If
    
    If tempplayer(index).InGame Or tempplayer(index).HasLogged Then
        If tempplayer(index).PVPTimer > timeGetTime Then
            Exit Sub
        End If
        
        tempplayer(index).InGame = False
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' Clear any invites out
        If tempplayer(index).TradeRequest > 0 Or tempplayer(index).PartyInvite > 0 Or tempplayer(index).GuildInvite > 0 Then
            If tempplayer(index).TradeRequest > 0 Then
                Call DeclineTradeRequest(index)
            End If
            
            If tempplayer(index).PartyInvite > 0 Then
                Call Party_InviteDecline(tempplayer(index).PartyInvite, index)
            End If
            
            If tempplayer(index).GuildInvite > 0 Then
                Call DeclineGuildInvite(index)
            End If
        End If
        
        ' Cancel any trade they're in
        If tempplayer(index).InTrade > 0 Then
            TradeTarget = tempplayer(index).InTrade
            PlayerMsg TradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade!", BrightRed
            
            ' Clear out trade
            For i = 1 To MAX_INV
                tempplayer(TradeTarget).TradeOffer(i).Num = 0
                tempplayer(TradeTarget).TradeOffer(i).Value = 0
            Next
            
            tempplayer(TradeTarget).InTrade = 0
            SendCloseTrade TradeTarget
        End If
        
        ' Leave party
        Party_PlayerLeave index

        ' Loop through entire map and purge npc targets from player
        For i = 1 To Map(GetPlayerMap(index)).NPC_HighIndex
            If MapNPC(GetPlayerMap(index)).NPC(i).Num > 0 Then
                If MapNPC(GetPlayerMap(index)).NPC(i).targetType = TARGET_TYPE_PLAYER Then
                    If MapNPC(GetPlayerMap(index)).NPC(i).target = index Then
                        MapNPC(GetPlayerMap(index)).NPC(i).target = 0
                        MapNPC(GetPlayerMap(index)).NPC(i).targetType = TARGET_TYPE_NONE
                        Call SendMapNPCTarget(GetPlayerMap(index), i, 0, 0)
                    End If
                End If
            End If
        Next
        
        ' Refresh guild members
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Not i = index Then
                    If GetPlayerGuild(i) = GetPlayerGuild(index) Then
                        SendPlayerGuildMembers i, index
                    End If
                End If
            End If
        Next
        
        ' Send a global message that they left
        If GetPlayerName(index) <> vbNullString Then
            If GetPlayerAccess(index) <= STAFF_MODERATOR Then
                Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Name & "!", Grey)
            Else
                Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Name & "!", DarkGrey)
            End If
        End If
        
        Call TextAdd(GetPlayerLogin(index) & " has logged out from " & Options.Name & ".")
        
        TotalPlayersOnline = TotalPlayersOnline - 1
        
        ' Leave game and clear data
        Call SendLeftGame(index)

        ' Refresh the friends list of all players online
        For i = 1 To Player_HighIndex
            Call UpdateFriendsList(i)
        Next
        
        ' Refresh the foes list of all players online
        For i = 1 To Player_HighIndex
            Call UpdateFoesList(i)
        Next
        
        Call SaveAccount(index)
        Call ClearAccount(index)
        
        ' re-set the high index
        Player_HighIndex = 0

        For i = MAX_PLAYERS To 1 Step -1

            If IsConnected(i) Then
                Player_HighIndex = i

                Exit For

            End If

        Next
        
        ' send the new highindex to all logged in players
        SendPlayer_HighIndex
        
        Call UpdateCaption
    End If
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Integer, ByVal x As Long, ByVal Y As Long, Optional ByVal NeedMap As Boolean = False, Optional ByVal Dir As Integer = -1)
    Dim ShopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub

    ' Check if you are out of bounds
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    If x < 0 Then x = 0
    If Y < 0 Then Y = 0
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    
    If OldMap <> MapNum Then
        UpdateMapBlock OldMap, GetPlayerX(index), GetPlayerY(index), False
    End If
    
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, Y)
    UpdateMapBlock MapNum, x, Y, True
    
    ' Set direction
    If Dir > -1 Then
        Call SetPlayerDir(index, Dir)
    End If
    
    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(index) And Not NeedMap Then
        Call SendPlayerXY(index)
        
        ' Clear spell casting
        ClearAccountSpellBuffer index
        Exit Sub
    End If
    
    ' Clear events
    tempplayer(index).EventProcessingCount = 0
    tempplayer(index).EventMap.CurrentEvents = 0
    
    ' Clear target
    tempplayer(index).target = 0
    tempplayer(index).targetType = TARGET_TYPE_NONE
    SendPlayerTarget index

    ' Loop through entire map and purge npc targets from player
    For i = 1 To Map(GetPlayerMap(index)).NPC_HighIndex
        If MapNPC(GetPlayerMap(index)).NPC(i).Num > 0 Then
            If MapNPC(GetPlayerMap(index)).NPC(i).targetType = TARGET_TYPE_PLAYER Then
                If MapNPC(GetPlayerMap(index)).NPC(i).target = index Then
                    MapNPC(GetPlayerMap(index)).NPC(i).target = 0
                    MapNPC(GetPlayerMap(index)).NPC(i).targetType = TARGET_TYPE_NONE
                    Call SendMapNPCTarget(OldMap, i, 0, 0)
                End If
            End If
        End If
    Next
    
    ' Leave the old map
    If Not OldMap = MapNum Then
        Call SendLeaveMap(index, OldMap)
        
        ' Set the new map
        Call SetPlayerMap(index, MapNum)
    End If
    
    ' Send player's equipment to new map
    SendMapEquipment index
    
    ' Send equipment of all people on new map
    If GetTotalMapPlayers(MapNum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    SendMapEquipmentTo i, index
                End If
            End If
        Next
    End If
    
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
        
        ' Get all NPCs' vitals
        For i = 1 To Map(OldMap).NPC_HighIndex
            If MapNPC(OldMap).NPC(i).Num > 0 Then
                MapNPC(OldMap).NPC(i).Vital(Vitals.HP) = GetNPCMaxVital(MapNPC(OldMap).NPC(i).Num, Vitals.HP)
            End If
        Next
    End If
    
    ' Clear spell casting
    ClearAccountSpellBuffer index
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    tempplayer(index).GettingMap = YES
    Set buffer = New clsBuffer
    Call SendCheckForMap(index, MapNum)
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal SendToSelf As Boolean = False)
    Dim buffer As clsBuffer, MapNum As Integer
    Dim x As Long, Y As Long, i As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim TileType As Long, VitalType As Long, Color As Long, Amount As Long
    Dim NewMapY As Long, NewMapX As Long
    Dim NewMapNum As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_DOWNRIGHT Or movement < 1 Or movement > 2 Then Exit Sub
    
    ' Don't allow them to move if they are transfering to a new map
    If tempplayer(index).GettingMap = YES Then Exit Sub
    
    ' Don't let them move if an event is waiting for their response
    If tempplayer(index).EventProcessingCount > 0 Then
        For i = 1 To tempplayer(index).EventProcessingCount
            If tempplayer(index).EventProcessing(i).WaitingForResponse > 0 Then
                Call SendPlayerXY(index)
                Exit Sub
            End If
        Next
    End If
    
    ' Prevent player from moving if they are casting a spell
    If tempplayer(index).SpellBuffer.Spell > 0 Then Exit Sub
    
    ' If stunned, stop them moving
    If tempplayer(index).StunDuration > 0 Then Exit Sub

    Call SetPlayerDir(index, Dir)
    
    Moved = NO
    MapNum = GetPlayerMap(index)
    
    Select Case Dir
        Case DIR_UPLEFT
            ' Check to make sure not outside of boundries
            If Map(GetPlayerMap(index)).Up > 0 And GetPlayerY(index) = 0 Then
                NewMapNum = Map(GetPlayerMap(index)).Up
                NewMapX = GetPlayerX(index)
                NewMapY = Map(NewMapNum).MaxY
            ElseIf Map(GetPlayerMap(index)).Left > 0 And GetPlayerX(index) = 0 Then
                NewMapNum = Map(GetPlayerMap(index)).Left
                NewMapX = Map(Map(GetPlayerMap(index)).Left).MaxX
                NewMapY = GetPlayerY(index)
            ElseIf GetPlayerX(index) - 1 > 0 And GetPlayerY(index) - 1 > 0 Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) And Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
                            Call SetPlayerY(index, GetPlayerY(index) - 1)
                            Call SetPlayerX(index, GetPlayerX(index) - 1)
                            SendPlayerMove index, movement, SendToSelf
                            Moved = YES
                        End If
                    End If
                End If
            End If
            
            If Moved = NO Then
                ' Check to see if we can move them to the another map
                If NewMapNum > 0 Then
                    If GetPlayerMap(index) <> NewMapNum Then
                        Call PlayerWarp(index, NewMapNum, NewMapX, NewMapY)
                        Moved = YES
                        ' clear their target
                        tempplayer(index).target = 0
                        tempplayer(index).targetType = TARGET_TYPE_NONE
                        SendPlayerTarget index
                    End If
                End If
            End If
            
        Case DIR_UPRIGHT
            ' Check to make sure not outside of boundries
            If Map(GetPlayerMap(index)).Up > 0 And GetPlayerY(index) = 0 Then
                NewMapNum = Map(GetPlayerMap(index)).Up
                NewMapX = GetPlayerX(index)
                NewMapY = Map(NewMapNum).MaxY
            ElseIf Map(GetPlayerMap(index)).Right > 0 And GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then
                NewMapNum = Map(GetPlayerMap(index)).Right
                NewMapX = 0
                NewMapY = GetPlayerY(index)
            ElseIf GetPlayerX(index) + 1 <= Map(GetPlayerMap(index)).MaxX And GetPlayerY(index) - 1 > 0 Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) And Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
                            Call SetPlayerY(index, GetPlayerY(index) - 1)
                            Call SetPlayerX(index, GetPlayerX(index) + 1)
                            SendPlayerMove index, movement, SendToSelf
                            Moved = YES
                        End If
                    End If
                End If
            End If
            
            ' Check to see if we can move them to the another map
            If Moved = NO Then
                If NewMapNum > 0 Then
                    Call PlayerWarp(index, NewMapNum, NewMapX, NewMapY)
                    Moved = YES
                    ' clear their target
                    tempplayer(index).target = 0
                    tempplayer(index).targetType = TARGET_TYPE_NONE
                    SendPlayerTarget index
                End If
            End If
            
        Case DIR_DOWNLEFT
            ' Check to make sure not outside of boundries
            If Map(GetPlayerMap(index)).Down > 0 And GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then
                NewMapNum = Map(GetPlayerMap(index)).Down
                NewMapX = GetPlayerX(index)
                NewMapY = 0
            ElseIf Map(GetPlayerMap(index)).Left > 0 And GetPlayerX(index) = 0 Then
                NewMapNum = Map(GetPlayerMap(index)).Left
                NewMapX = Map(NewMapNum).MaxX
                NewMapY = GetPlayerY(index)
            ElseIf GetPlayerX(index) - 1 > 0 And GetPlayerY(index) + 1 <= Map(GetPlayerMap(index)).MaxY Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) And Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
                            Call SetPlayerY(index, GetPlayerY(index) + 1)
                            Call SetPlayerX(index, GetPlayerX(index) - 1)
                            SendPlayerMove index, movement, SendToSelf
                            Moved = YES
                        End If
                    End If
                End If
            End If
            
            ' Check to see if we can move them to the another map
            If Moved = NO Then
                If NewMapNum > 0 Then
                    If GetPlayerMap(index) <> NewMapNum Then
                        Call PlayerWarp(index, NewMapNum, NewMapX, NewMapY)
                        Moved = YES
                        ' clear their target
                        tempplayer(index).target = 0
                        tempplayer(index).targetType = TARGET_TYPE_NONE
                        SendPlayerTarget index
                    End If
                End If
            End If
            
        Case DIR_DOWNRIGHT
            ' Check to make sure not outside of boundries
            If Map(GetPlayerMap(index)).Down > 0 And GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then
                NewMapNum = Map(GetPlayerMap(index)).Down
                NewMapX = GetPlayerX(index)
                NewMapY = 0
            ElseIf Map(GetPlayerMap(index)).Right > 0 And GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then
                NewMapNum = Map(GetPlayerMap(index)).Right
                NewMapX = 0
                NewMapY = GetPlayerY(index)
            ElseIf GetPlayerX(index) + 1 <= Map(GetPlayerMap(index)).MaxX And GetPlayerY(index) + 1 <= Map(GetPlayerMap(index)).MaxY Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) And Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
                            Call SetPlayerY(index, GetPlayerY(index) + 1)
                            Call SetPlayerX(index, GetPlayerX(index) + 1)
                            SendPlayerMove index, movement, SendToSelf
                            Moved = YES
                        End If
                    End If
                End If
            End If
            
            ' Check to see if we can move them to the another map
            If Moved = NO Then
                If NewMapNum > 0 Then
                    If GetPlayerMap(index) <> NewMapNum Then
                        Call PlayerWarp(index, NewMapNum, NewMapX, NewMapY)
                        Moved = YES
                        ' clear their target
                        tempplayer(index).target = 0
                        tempplayer(index).targetType = TARGET_TYPE_NONE
                        SendPlayerTarget index
                    End If
                End If
            End If
            
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Then
                    If Not IsPlayerBlocked(index, 0, -1) Then
                        If Not IsEventBlocked(index, 0, -1) Then
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
                                    Call SetPlayerY(index, GetPlayerY(index) - 1)
                                    SendPlayerMove index, movement, SendToSelf
                                    Moved = YES
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    If GetPlayerMap(index) <> Map(GetPlayerMap(index)).Up Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), Map(Map(GetPlayerMap(index)).Up).MaxY)
                        Moved = YES
                        
                        ' Clear their target
                        tempplayer(index).target = 0
                        tempplayer(index).targetType = TARGET_TYPE_NONE
                        SendPlayerTarget index
                    End If
                End If
            End If

        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(MapNum).MaxY Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Not IsPlayerBlocked(index, 0, 1) Then
                        If Not IsEventBlocked(index, 0, 1) Then
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
                                    Call SetPlayerY(index, GetPlayerY(index) + 1)
                                    SendPlayerMove index, movement, SendToSelf
                                    Moved = YES
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    If GetPlayerMap(index) <> Map(GetPlayerMap(index)).Down Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                        Moved = YES
                        
                        ' Clear their target
                        tempplayer(index).target = 0
                        tempplayer(index).targetType = TARGET_TYPE_NONE
                        SendPlayerTarget index
                    End If
                End If
            End If

        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Not IsPlayerBlocked(index, -1, 0) Then
                        If Not IsEventBlocked(index, -1, 0) Then
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                                    Call SetPlayerX(index, GetPlayerX(index) - 1)
                                    SendPlayerMove index, movement, SendToSelf
                                    Moved = YES
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    If GetPlayerMap(index) <> Map(GetPlayerMap(index)).Left Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, Map(Map(GetPlayerMap(index)).Left).MaxX, GetPlayerY(index))
                    Moved = YES
                    
                    ' Clear their target
                    tempplayer(index).target = 0
                    tempplayer(index).targetType = TARGET_TYPE_NONE
                    SendPlayerTarget index
                    End If
                End If
            End If

        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(MapNum).MaxX Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Not IsPlayerBlocked(index, 1, 0) Then
                        If Not IsEventBlocked(index, 1, 0) Then
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                                    Call SetPlayerX(index, GetPlayerX(index) + 1)
                                    SendPlayerMove index, movement, SendToSelf
                                    Moved = YES
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    If GetPlayerMap(index) <> Map(GetPlayerMap(index)).Right Then
                        Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                        Moved = YES
                        
                        ' Clear their target
                        tempplayer(index).target = 0
                        tempplayer(index).targetType = TARGET_TYPE_NONE
                        SendPlayerTarget index
                    End If
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            MapNum = .Data1
            x = .Data2
            Y = .Data3
            Call PlayerWarp(index, MapNum, x, Y)
            Moved = YES
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            
            If x > 0 Then ' Shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' Name exists?
                    SendOpenShop index, x
                    tempplayer(index).InShop = x ' Stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            tempplayer(index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            Amount = .Data2
            
            If VitalType = Int(Vitals.HP) Then
                Color = BrightGreen
            ElseIf VitalType = Int(Vitals.MP) Then
                Color = BrightBlue
            End If
            
            If Not GetPlayerVital(index, VitalType) = GetPlayerMaxVital(index, VitalType) Then
                If GetPlayerVital(index, VitalType) + Amount > GetPlayerMaxVital(index, VitalType) Then
                    Amount = GetPlayerMaxVital(index, VitalType) - GetPlayerVital(index, VitalType)
                End If
                SendActionMsg GetPlayerMap(index), "+" & Amount, Color, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + Amount
                Call SendVital(index, VitalType)
            Else
                SendActionMsg GetPlayerMap(index), "+0", Color, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                If tempplayer(index).InParty > 0 Then SendPartyVitals tempplayer(index).InParty, index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            VitalType = .Data1
            Amount = .Data2
            
            If VitalType = Int(Vitals.HP) Then
                Color = BrightRed
            ElseIf VitalType = Int(Vitals.MP) Then
                Color = Magenta
            End If
            
            If Not GetPlayerVital(index, VitalType) < 1 Then
                If GetPlayerVital(index, VitalType) - Amount < 1 Then
                    Amount = GetPlayerVital(index, VitalType)
                End If
                SendActionMsg GetPlayerMap(index), "-" & Amount, Color, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                If GetPlayerVital(index, HP) - Amount < 1 And VitalType = 1 Then
                    Call OnDeath(index)
                    Call GlobalMsg(GetPlayerName(index) & " has been killed by a trap!", BrightRed)
                Else
                    SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) - Amount
                    Call SendVital(index, VitalType)
                End If
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - Amount
                PlayerMsg index, "You're injured by a trap.", BrightRed
                Call SendVital(index, HP)
                ' Send vitals to party if in one
                If tempplayer(index).InParty > 0 Then SendPartyVitals tempplayer(index).InParty, index
            End If
            Moved = YES
        End If
            
        ' Checkpoint
        If .Type = TILE_TYPE_CHECKPOINT Then
            SetCheckpoint index, .Data1, .Data2, .Data3
            Moved = YES
        End If
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove index, MOVING_RUNNING, GetPlayerDir(index)
            Moved = YES
        End If
    End With
    
    ' They tried to hack
    If Moved = NO Then
        Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    Else
        x = GetPlayerX(index)
        Y = GetPlayerY(index)
    
        If Trim$(Account(index).Chars(GetPlayerChar(index)).Status) = "AFK" Then
            Account(index).Chars(GetPlayerChar(index)).Status = vbNullString
            Call SendPlayerStatus(index)
        End If
        
        ' Check to see if events are touched
        EventTouch index, x, Y
    End If
End Sub

Sub EventTouch(ByVal index As Long, ByVal x As Long, ByVal Y As Long)
    Dim EventTouched As Boolean, i As Long
    
    If tempplayer(index).EventMap.CurrentEvents > 0 Then
            For i = 1 To tempplayer(index).EventMap.CurrentEvents
                If Map(GetPlayerMap(index)).Events(tempplayer(index).EventMap.EventPages(i).eventID).Global = 1 Then
                    If Map(GetPlayerMap(index)).Events(tempplayer(index).EventMap.EventPages(i).eventID).x = x And Map(GetPlayerMap(index)).Events(tempplayer(index).EventMap.EventPages(i).eventID).Y = Y And Map(GetPlayerMap(index)).Events(tempplayer(index).EventMap.EventPages(i).eventID).Pages(tempplayer(index).EventMap.EventPages(i).PageID).Trigger = 1 And tempplayer(index).EventMap.EventPages(i).Visible = 1 Then EventTouched = True
                Else
                    If tempplayer(index).EventMap.EventPages(i).x = x And tempplayer(index).EventMap.EventPages(i).Y = Y And Map(GetPlayerMap(index)).Events(tempplayer(index).EventMap.EventPages(i).eventID).Pages(tempplayer(index).EventMap.EventPages(i).PageID).Trigger = 1 And tempplayer(index).EventMap.EventPages(i).Visible = 1 Then EventTouched = True
                End If
                
                If EventTouched Then
                    ' Process this event, it is on-touch and everything checks out.
                    If Map(GetPlayerMap(index)).Events(tempplayer(index).EventMap.EventPages(i).eventID).Pages(tempplayer(index).EventMap.EventPages(i).PageID).CommandListCount > 0 Then
                        tempplayer(index).EventProcessingCount = tempplayer(index).EventProcessingCount + 1
                        ReDim Preserve tempplayer(index).EventProcessing(tempplayer(index).EventProcessingCount)
                        tempplayer(index).EventProcessing(tempplayer(index).EventProcessingCount).ActionTimer = timeGetTime
                        tempplayer(index).EventProcessing(tempplayer(index).EventProcessingCount).CurList = 1
                        tempplayer(index).EventProcessing(tempplayer(index).EventProcessingCount).CurSlot = 1
                        tempplayer(index).EventProcessing(tempplayer(index).EventProcessingCount).eventID = tempplayer(index).EventMap.EventPages(i).eventID
                        tempplayer(index).EventProcessing(tempplayer(index).EventProcessingCount).PageID = tempplayer(index).EventMap.EventPages(i).PageID
                        tempplayer(index).EventProcessing(tempplayer(index).EventProcessingCount).WaitingForResponse = 0
                        ReDim tempplayer(index).EventProcessing(tempplayer(index).EventProcessingCount).ListLeftOff(0 To Map(GetPlayerMap(index)).Events(tempplayer(index).EventMap.EventPages(i).eventID).Pages(tempplayer(index).EventMap.EventPages(i).PageID).CommandListCount)
                    End If
                    
                    EventTouched = False
                End If
            Next
        End If
End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_DOWNRIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub

    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
        Case DIR_UPLEFT
            If GetPlayerY(index) = 0 And GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_UPRIGHT
            If GetPlayerY(index) = 0 And GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
        Case DIR_DOWNLEFT
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY And GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWNRIGHT
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY And GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select

    PlayerMove index, Direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim ItemNum As Integer
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(index, i)

        If ItemNum > 0 Then
            If Not Item(ItemNum).Type = ITEM_TYPE_EQUIPMENT Or Not Item(ItemNum).EquipSlot = i Then SetPlayerEquipment index, 0, i
        Else
            SetPlayerEquipment index, 0, i
        End If
    Next
End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

    If Item(ItemNum).stackable = 1 Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next

    End If

    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next
End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Integer) As Byte
    Dim i As Long

    ' Check for subscript out of range
    If Not IsPlaying(index) Or ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function

    If Not Item(ItemNum).Type = ITEM_TYPE_EQUIPMENT Then
        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next
    End If

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next
End Function

Function CheckBankSlots(ByVal index As Long, ByVal ItemNum As Integer) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function

    For i = 1 To MAX_BANK
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(index, i) = ItemNum Then
            CheckBankSlots = CheckBankSlots + 1
        End If
    Next
End Function

Function CheckInventorySlots(ByVal index As Long, ByVal ItemNum As Integer) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            CheckInventorySlots = CheckInventorySlots + 1
        End If
    Next
End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Integer) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).stackable = 1 Then
                HasItem = GetPlayerInvItemValue(index, i)
                Exit Function
            ElseIf Item(ItemNum).stackable = 0 Then
                HasItem = 1
                Exit Function
            End If
        End If
    Next
End Function

Function TakeInvItem(ByVal index As Long, ByVal ItemNum As Integer, ByVal ItemVal As Long, Optional Update As Boolean = True) As Boolean
    Dim i As Long, ii As Long, NPCNum As Long
    Dim n As Long
    Dim Parse() As String

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).stackable = 1 Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    
                    If Update Then Call SendInventoryUpdate(index, i)
                    
                    'check quests
                    For ii = 1 To MAX_QUESTS
                        Parse() = Split(HasQuestItems(index, ii, True), "|")
                        If UBound(Parse()) > 0 Then
                            NPCNum = Parse(0)
                            If NPCNum > 0 Then
                                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
                            End If
                        End If
                    Next ii
                    
                    Exit Function
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Call SetPlayerInvItemDur(index, i, 0)
                Call SetPlayerInvItemBind(index, i, 0)
                Exit For
            End If
        End If
    Next
    
    'check quests
    For ii = 1 To MAX_QUESTS
        Parse() = Split(HasQuestItems(index, ii, True), "|")
        If UBound(Parse()) > 0 Then
            NPCNum = Parse(0)
            If NPCNum > 0 Then
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
            End If
        End If
    Next ii
    
    ' Send the inventory update
    If Update Then Call SendInventory(index)
End Function

Function TakeInvSlot(ByVal index As Long, ByVal InvSlot As Byte, ByVal ItemVal As Long, Optional ByVal Update As Boolean = True) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ItemNum As Integer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvSlot < 1 Or InvSlot > MAX_ITEMS Then Exit Function
    
    ItemNum = GetPlayerInvItemNum(index, InvSlot)

    ' Prevent subscript out of range
    If ItemNum < 1 Then Exit Function
    
    If Item(ItemNum).stackable = 1 Then
        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, InvSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, InvSlot, GetPlayerInvItemValue(index, InvSlot) - ItemVal)
            
            ' Send the inventory update
            If Update Then
                Call SendInventoryUpdate(index, InvSlot)
            End If
            Exit Function
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, InvSlot, 0)
        Call SetPlayerInvItemValue(index, InvSlot, 0)
        Call SetPlayerInvItemDur(index, InvSlot, 0)
        Call SetPlayerInvItemBind(index, InvSlot, 0)
        
        ' Send the inventory update
        If Update Then
            Call SendInventoryUpdate(index, InvSlot)
        End If
    End If
End Function

Function GiveInvItem(ByVal index As Long, ByVal ItemNum As Integer, ByVal ItemVal As Long, Optional ByVal ItemDur As Integer = -1, Optional ByVal ItemBind As Integer = 0, Optional ByVal SendUpdate As Boolean = True) As Byte
    Dim i As Long, ii As Long, NPCNum As Long, x As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

    i = FindOpenInvSlot(index, ItemNum)

    ' Check to see if inventory is full
    If i > 0 And i <= MAX_INV Then
        If CDec(GetPlayerInvItemValue(index, i)) + CDec(ItemVal) > 2147483468 Then
            Call PlayerMsg(index, "Cannot give it to you, it exceeds the maximum limit!", BrightRed)
            Exit Function
        Else
            Call SetPlayerInvItemNum(index, i, ItemNum)
            
            If Item(ItemNum).stackable = 1 Then
                Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
            Else
                If Item(ItemNum).Type <> ITEM_TYPE_EQUIPMENT Then
                    Call SetPlayerInvItemValue(index, i, 1)
                ElseIf ItemVal = 0 Then
                    ItemVal = 1
                End If
                
                For x = 1 To ItemVal - 1
                    ii = FindOpenInvSlot(index, ItemNum)
                    
                    If ii > 0 And ii <= MAX_INV Then
                        Call SetPlayerInvItemNum(index, ii, ItemNum)
                        If Item(ItemNum).Type <> ITEM_TYPE_EQUIPMENT Then Call SetPlayerInvItemValue(index, ii, 1)
                        
                        If Item(GetPlayerInvItemNum(index, ii)).Type = ITEM_TYPE_EQUIPMENT Then
                            If ItemDur = -1 Then
                                Call SetPlayerInvItemDur(index, ii, Item(ItemNum).Data1)
                            Else
                                Call SetPlayerInvItemDur(index, ii, ItemDur)
                            End If
                        End If
                        
                        If ItemBind = BIND_ON_PICKUP Or Item(GetPlayerInvItemNum(index, ii)).BindType = BIND_ON_PICKUP Then
                            Call SetPlayerInvItemBind(index, ii, BIND_ON_PICKUP)
                        ElseIf ItemBind = BIND_ON_EQUIP Or Item(GetPlayerInvItemNum(index, ii)).BindType = BIND_ON_EQUIP Then
                            Call SetPlayerInvItemBind(index, ii, BIND_ON_EQUIP)
                        Else
                            Call SetPlayerInvItemBind(index, ii, 0)
                        End If
                        
                        If SendUpdate Then Call SendInventoryUpdate(index, ii)
                    Else
                        For ii = x To ItemVal - 1
                            If Item(ItemNum).Type <> ITEM_TYPE_EQUIPMENT Then
                                If ItemDur = -1 Then
                                    Call SpawnItem(ItemNum, 1, Item(ItemNum).Data1, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                                Else
                                    Call SpawnItem(ItemNum, 1, ItemDur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                                End If
                            Else
                                If ItemDur = -1 Then
                                    Call SpawnItem(ItemNum, 0, Item(ItemNum).Data1, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                                Else
                                    Call SpawnItem(ItemNum, 0, ItemDur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                                End If
                            End If
                        Next
                        
                        Call PlayerMsg(index, "Your inventory is full!", BrightRed)
                        Exit For
                    End If
                Next
            End If
        End If
        
        If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_EQUIPMENT Then
            If ItemDur = -1 Then
                Call SetPlayerInvItemDur(index, i, Item(ItemNum).Data1)
            Else
                Call SetPlayerInvItemDur(index, i, ItemDur)
            End If
        End If
        
        If ItemBind = BIND_ON_PICKUP Or Item(GetPlayerInvItemNum(index, i)).BindType = BIND_ON_PICKUP Then
            Call SetPlayerInvItemBind(index, i, BIND_ON_PICKUP)
        ElseIf ItemBind = BIND_ON_EQUIP Or Item(GetPlayerInvItemNum(index, i)).BindType = BIND_ON_EQUIP Then
            Call SetPlayerInvItemBind(index, i, BIND_ON_EQUIP)
        Else
            Call SetPlayerInvItemBind(index, i, 0)
        End If
        
        If SendUpdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
        
        'check quests
        For ii = 1 To MAX_QUESTS
            NPCNum = HasQuestItems(index, ii)
            If NPCNum > 0 Then
                Call SendShowTaskCompleteOnNPC(index, NPCNum, True)
            End If
        Next ii
    Else
        Call PlayerMsg(index, "Your inventory is full!", BrightRed)
    End If
    
    GiveInvItem = i
End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next
End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next
End Function

Sub PlayerMapGetItem(ByVal index As Long, ByVal i As Long)
    Dim n As Long
    Dim MapNum As Integer
    Dim Msg As String
    Dim tempVal As Variant
    Dim ItemNum As Long, Value As Long, Dur As Long, Bind As Long
    
    ' Check for subscript out of range
    If Not IsPlaying(index) Then Exit Sub
    
    MapNum = GetPlayerMap(index)

    ' See if there's even an item here
    If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
        ' Can we pick the item up?
        If CanPlayerPickupItem(index, i) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).x = GetPlayerX(index)) Then
                If (MapItem(MapNum, i).Y = GetPlayerY(index)) Then
                    ItemNum = MapItem(MapNum, i).Num
                    Value = MapItem(MapNum, i).Value
                    Dur = MapItem(MapNum, i).Durability
                    Bind = Item(ItemNum).BindType
                    If Value > 0 Then
                        Msg = Value & " " & Trim$(Item(ItemNum).Name)
                    Else
                        Msg = Trim$(Item(ItemNum).Name)
                    End If
                    
                    'sure made this a lot simpler than it was, removing roughly 30 lines of code in exchange for 5.  It could be done in 1 line
                    'but I chose to make it pretty and easy to debug had something went wrong.
                    Call GiveInvItem(index, ItemNum, Value, Dur, Bind, True)
                    
                    ' Erase the item from the map
                    MapItem(MapNum, i).Num = 0
                    MapItem(MapNum, i).Value = 0
                    MapItem(MapNum, i).Durability = 0
                    MapItem(MapNum, i).x = 0
                    MapItem(MapNum, i).Y = 0
                    
                    Call SendInventoryUpdate(index, n)
                    Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), 0, 0)
                    SendActionMsg GetPlayerMap(index), Msg, Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                End If
            End If
        End If
    End If
End Sub

Public Function CanPlayerPickupItem(ByVal index As Long, ByVal MapItemNum As Integer, Optional ByVal ItemVal As Long = 1)
    Dim MapNum As Integer

    MapNum = GetPlayerMap(index)
    
    ' Check for subscript out of range
    If MapNum < 1 Or MapNum > MAX_MAPS Then Exit Function
    
    If Moral(Map(MapNum).Moral).CanPickupItem = 1 Then
        ' No lock or locked to player?
        If Trim$(MapItem(MapNum, MapItemNum).playerName) = vbNullString Or Trim$(MapItem(MapNum, MapItemNum).playerName) = GetPlayerName(index) Then
            CanPlayerPickupItem = True
            Exit Function
        End If
    End If
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Byte, ByVal Amount As Long)
    Dim i As Long
    Dim Msg As String
    
    If (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(index))

        If Not i = 0 Then
            MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
            MapItem(GetPlayerMap(index), i).Y = GetPlayerY(index)
            MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
            MapItem(GetPlayerMap(index), i).PlayerTimer = timeGetTime + ITEM_SPAWN_TIME
            MapItem(GetPlayerMap(index), i).CanDespawn = True
            MapItem(GetPlayerMap(index), i).DespawnTimer = timeGetTime + ITEM_DESPAWN_TIME

            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_EQUIPMENT Then
                MapItem(GetPlayerMap(index), i).Durability = GetPlayerInvItemDur(index, InvNum)
            Else
                MapItem(GetPlayerMap(index), i).Durability = 0
            End If
            
            If Item(GetPlayerInvItemNum(index, InvNum)).stackable = 1 Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, InvNum)
                    Msg = GetPlayerInvItemValue(index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name)
                    
                    Call TakeInvSlot(index, InvNum, GetPlayerInvItemValue(index, InvNum), True)
                Else
                    MapItem(GetPlayerMap(index), i).Value = Amount
                    Msg = Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name)
                    Call TakeInvSlot(index, InvNum, Amount, True)
                End If
            Else
                ' It's not a currency object so this is easy
                Msg = Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name)
                MapItem(GetPlayerMap(index), i).Value = 0
                Call TakeInvSlot(index, InvNum, Amount, True)
            End If
            
            ' Send message
            SendActionMsg GetPlayerMap(index), Msg, BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, Amount, MapItem(GetPlayerMap(index), i).Durability, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsg(index, "There are too many items on the ground to drop anything else.", BrightRed)
        End If
    End If
End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim ExpRollOver As Long
    Dim Level_Count As Long

    If GetPlayerLevel(index) > 0 And GetPlayerLevel(index) < Options.MaxLevel Then
        Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
            ExpRollOver = GetPlayerExp(index) - GetPlayerNextLevel(index)
            Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
            Call SetPlayerPoints(index, GetPlayerPoints(index) + Options.StatsLevel)
            Call SetPlayerExp(index, ExpRollOver)
            Level_Count = Level_Count + 1
        Loop
        
        If Level_Count > 0 Then
            Call SendAnimation(GetPlayerMap(index), Options.LevelUpAnimation, 0, 0, TARGET_TYPE_PLAYER, index)
            Call PlayerMsg(index, "You are now level " & GetPlayerLevel(index) & ".", BrightGreen)
            
            ' Check for new title
            Call CheckPlayerNewTitle(index)
            
            ' Check if any of the player's spells can rank up
            For i = 1 To MAX_PLAYER_SPELLS
                If GetPlayerSpell(index, i) > 0 Then
                    If Spell(GetPlayerSpell(index, i)).NewSpell > 0 Then
                        If Spell(Spell(GetPlayerSpell(index, i)).NewSpell).CastRequired > 0 Then
                            Call CheckSpellRankUp(index, GetPlayerSpell(index, i), i)
                        End If
                    End If
                End If
            Next
            
            ' Restore and send vitals
            For i = 1 To Vitals.Vital_Count - 1
                Call SetPlayerVital(index, i, GetPlayerMaxVital(index, i))
                Call SendVital(index, i)
            Next
            
            ' Send updated data
            Call SendPlayerExp(index)
            Call SendPlayerStats(index)
            Call SendPlayerPoints(index)
            Call SendPlayerLevel(index)
        End If
    End If
End Sub

Sub CheckPlayerSkillLevelUp(ByVal index As Long, ByVal SkillNum As Byte)
    Dim ExpRollOver As Long
    Dim Level_Count As Long
    
    Level_Count = 0

    If GetPlayerSkill(index, SkillNum) > 0 And GetPlayerSkill(index, SkillNum) < Options.MaxLevel Then
        Do While GetPlayerSkillExp(index, SkillNum) >= GetPlayerNextSkillLevel(index, SkillNum)
            ExpRollOver = GetPlayerSkillExp(index, SkillNum) - GetPlayerNextSkillLevel(index, SkillNum)
            Call SetPlayerSkill(index, GetPlayerSkill(index, SkillNum) + 1, SkillNum)
            Call SetPlayerSkillExp(index, ExpRollOver, SkillNum)
            Level_Count = Level_Count + 1
        Loop
        
        If Level_Count > 0 Then
            Call SendAnimation(GetPlayerMap(index), Options.LevelUpAnimation, 0, 0, TARGET_TYPE_PLAYER, index)
            Call PlayerMsg(index, "Your " & CheckGrammar(GetSkillName(SkillNum)) & " level is now " & GetPlayerSkill(index, SkillNum) & ".", BrightGreen)
            Call SendPlayerSkills(index)
        End If
    End If
End Sub

Private Function AutoLife(ByVal index As Long) As Boolean
    Dim i As Byte
    
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) > 0 Then
            If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_AUTOLIFE Then
                If CanPlayerUseItem(index, GetPlayerInvItemNum(index, i), False) Then
                    ' HP
                    If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddHP > 0 Then
                        If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddHP > GetPlayerMaxVital(index, HP) Then
                            SendActionMsg GetPlayerMap(index), "+" & GetPlayerMaxVital(index, HP), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        Else
                            SendActionMsg GetPlayerMap(index), "+" & Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        End If
                        Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddHP)
                        Call SendVital(index, Vitals.HP)
                    End If
                    
                    ' MP
                    If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddMP > 0 Then
                        If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddMP > GetPlayerMaxVital(index, MP) Then
                            SendActionMsg GetPlayerMap(index), "+" & GetPlayerMaxVital(index, MP), BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        Else
                            SendActionMsg GetPlayerMap(index), "+" & Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        End If
                        Call SendVital(index, Vitals.MP)
                        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) + Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddMP)
                    End If
                    
                    ' If it is not reusable then take the item away
                    If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).IsReusable = False Then
                        Call TakeInvItem(index, Account(index).Chars(GetPlayerChar(index)).Inv(i).Num, 0)
                    End If
                    
                    Call SendAnimation(GetPlayerMap(index), Item(GetPlayerInvItemNum(index, i)).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                    
                    ' Warp player away
                    If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).Data1 = 1 Then
                        Call WarpPlayer(index)
                    End If
                    
                    Call PlayerMsg(index, "You have been given another life!", Yellow)
                    
                    AutoLife = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Sub OnDeath(ByVal index As Long, Optional ByVal Attacker As Long = 0)
    Dim i As Long, RemoveItem As Boolean
   
    ' Set HP to 0
    Call SetPlayerVital(index, Vitals.HP, 0)
    
    ' Exit out if they were saved
    If AutoLife(index) Then Exit Sub
    
    ' If map moral can drop items or not
    If Moral(Map(GetPlayerMap(index)).Moral).DropItems = 1 Or GetPlayerPK(index) = PLAYER_KILLER Or (GetPlayerPK(index) = PLAYER_DEFENDER And GetPlayerPK(Attacker) = PLAYER_KILLER) Then
        If GetPlayerPK(index) <> NO Then
            Call SetPlayerPK(index, NO)
            Call SendPlayerPK(index)
        End If

        ' Drop all worn items
        For i = 1 To Equipment.Equipment_Count - 1
            RemoveItem = False
            
            If GetPlayerEquipment(index, i) > 0 Then
                If tempplayer(index).InParty > 0 Then
                    Call Party_GetLoot(tempplayer(Attacker).InParty, GetPlayerEquipment(index, i), 1, GetPlayerX(index), GetPlayerY(index))
                    RemoveItem = True
                Else
                    If Moral(Map(GetPlayerMap(index)).Moral).CanDropItem = 1 Then
                        If Attacker > 0 Then
                            Call SpawnItem(GetPlayerEquipment(index, i), 1, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(Attacker))
                            RemoveItem = True
                        Else
                            Call SpawnItem(GetPlayerEquipment(index, i), 1, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                            RemoveItem = True
                        End If
                    Else
                        If Attacker > 0 Then
                            Call GiveInvItem(Attacker, GetPlayerEquipment(index, i), 1)
                            RemoveItem = True
                        End If
                    End If
                End If
                            
                ' Remove equipment item
                If RemoveItem Then
                    ' Send a message to the world indicating that they dropped an item
                    Call GlobalMsg(GetPlayerName(index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerEquipment(index, i)).Name)) & "!", Yellow)
                    
                    SetPlayerEquipment index, 0, i
                    SetPlayerEquipmentDur index, 0, i
                    SetPlayerEquipmentBind index, 0, i
                End If
            End If
        Next
        
        ' Drop 10% of their Gold
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = 1 Then
                If Round(GetPlayerInvItemValue(index, i) / 10) > 0 Then
                    Call TakeInvItem(index, GetPlayerInvItemNum(index, i), Round(GetPlayerInvItemValue(index, i) / 10))
                    Call SpawnItem(1, Round(GetPlayerInvItemValue(index, i) / 10), 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(Attacker))
                    Exit For
                End If
            End If
        Next
    
        ' Add the player kill
        If Attacker > 0 Then Account(FindPlayer(GetPlayerName(Attacker))).Chars(GetPlayerChar(i)).PlayerKills = Account(FindPlayer(GetPlayerName(Attacker))).Chars(GetPlayerChar(i)).PlayerKills + 1
        
        ' Check for new title
        Call CheckPlayerNewTitle(index)
    End If
    
    ' Loop through entire map and purge npc targets from player
    For i = 1 To Map(GetPlayerMap(index)).NPC_HighIndex
        If MapNPC(GetPlayerMap(index)).NPC(i).Num > 0 Then
            If MapNPC(GetPlayerMap(index)).NPC(i).targetType = TARGET_TYPE_PLAYER Then
                If MapNPC(GetPlayerMap(index)).NPC(i).target = index Then
                    MapNPC(GetPlayerMap(index)).NPC(i).target = 0
                    MapNPC(GetPlayerMap(index)).NPC(i).targetType = TARGET_TYPE_NONE
                    Call SendMapNPCTarget(GetPlayerMap(index), i, 0, 0)
                End If
            End If
        End If
    Next

    ' Set player direction
    Call SetPlayerDir(index, DIR_DOWN)
    
    ' Warp away player
    Call WarpPlayer(index)
    
    ' Clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With tempplayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With tempplayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    Call ClearAccountSpellBuffer(index)
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))

    ' Send vitals to party if in one
    If tempplayer(index).InParty > 0 Then SendPartyVitals tempplayer(index).InParty, index
    
    ' Send vitals
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
End Sub

Private Sub WarpPlayer(ByVal index As Long)
     With Map(GetPlayerMap(index))
        If .BootMap = 0 Then
            ' Warp to the checkpoint
            Call WarpToCheckPoint(index)
        Else
            ' Warp to the boot map
            If .BootMap > 0 And .BootMap <= MAX_MAPS Then
                PlayerWarp index, .BootMap, .BootX, .BootY
            Else
                ' Warp to the start map
                Call PlayerWarp(index, Class(GetPlayerClass(index)).Map, Class(GetPlayerClass(index)).x, Class(GetPlayerClass(index)).Y, False, Class(GetPlayerClass(index)).Dir)
            End If
        End If
     End With
End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal Y As Long)
    Dim Resource_Num As Long
    Dim Resource_Index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    Dim RndNum As Long
    
    If Map(GetPlayerMap(index)).Tile(x, Y).Type = TILE_TYPE_RESOURCE Then
        Resource_Num = 0
        Resource_Index = Map(GetPlayerMap(index)).Tile(x, Y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).Y = Y Then
                    Resource_Num = i
                End If
            End If
        Next

        If Resource_Num > 0 Then
            ' Check if they meet the level required
            If Resource(Resource_Index).LevelReq > 0 Then
                If GetPlayerSkill(index, Resource(Resource_Index).Skill) < Resource(Resource_Index).LevelReq Then
                    Call PlayerMsg(index, "Your " & CheckGrammar(GetSkillName(Resource(Resource_Index).LevelReq)) & " skill level does not meet the requirement to use this resource!", BrightRed)
                    Exit Sub
                End If
            End If
            
            ' Check if they have the right tool
            If Resource(Resource_Index).ToolRequired > 0 Then
                If GetPlayerEquipment(index, Weapon) < 1 Then
                    PlayerMsg index, "You need a tool to interact with this resource!", BrightRed
                    Exit Sub
                End If
                
                If Item(GetPlayerEquipment(index, Weapon)).Tool <> Resource(Resource_Index).ToolRequired Then
                    PlayerMsg index, "You have the wrong type of item equipped to use this resource!", BrightRed
                    Exit Sub
                End If
            End If
                
            ' Enough space in inventory?
            If Resource(Resource_Index).ItemReward > 0 Then
                If FindOpenInvSlot(index, Resource(Resource_Index).ItemReward) = 0 Then
                    PlayerMsg index, "You do not have enough inventory space!", BrightRed
                    Exit Sub
                End If
            End If

            ' Check if the resource has already been deplenished
            If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceState = 0 Then
                rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).x
                rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Y
            
                ' Reduce weapon's durability
                Call DamagePlayerEquipment(index, Equipment.Weapon)
                
                ' Give the reward random when they deal damage
                RndNum = Random(Resource(Resource_Index).LowChance, Resource(Resource_Index).HighChance)
                  
                If Not RndNum = Resource(Resource_Index).LowChance Then
                    ' Subtract the RndNum by the random value of the weapon's chance modifier
                    RndNum = RndNum - Round(Random((Item(GetPlayerEquipment(index, Weapon)).ChanceModifier / 2), Item(GetPlayerEquipment(index, Weapon)).ChanceModifier))
                    
                    ' If value is less than the resource low chance then set it to it
                    If RndNum < Resource(Resource_Index).LowChance Then
                        RndNum = Resource(Resource_Index).LowChance
                    End If
                End If
                
                If RndNum = Resource(Resource_Index).LowChance Then
                    ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Cur_Reward = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Cur_Reward - 1
                    GiveInvItem index, Resource(Resource_Index).ItemReward, 1
                    
                    If GetPlayerSkill(index, Resource(Resource_Index).Skill) < Options.MaxLevel Then
                        ' Add the experience to the skill
                        Call SetPlayerSkillExp(index, GetPlayerSkillExp(index, Resource(Resource_Index).Skill) + Resource(Resource_Index).Exp * EXP_RATE, Resource(Resource_Index).Skill)
                        
                        ' Check for skill level up
                        Call CheckPlayerSkillLevelUp(index, Resource(Resource_Index).Skill)
                    End If
                    
                    ' Send message if it exists
                    If Len(Trim$(Resource(Resource_Index).SuccessMessage)) > 0 Then
                        SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_Index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                    End If
                    
                    ' If the resource is empty then clear it
                    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Cur_Reward = 0 Then
                        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceState = 1
                        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceTimer = timeGetTime
                        SendResourceCacheToMap GetPlayerMap(index)
                    End If
                Else
                    ' Send message if it exists
                    If Len(Trim$(Resource(Resource_Index).FailMessage)) > 0 Then
                        SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_Index).FailMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                    End If
                End If
                
                SendAnimation GetPlayerMap(index), Resource(Resource_Index).Animation, rX, rY
                
                ' Send the sound
                SendMapSound GetPlayerMap(index), index, rX, rY, SoundEntity.seResource, Resource_Index
            Else
                ' Send message if it exists
                If Len(Trim$(Resource(Resource_Index).EmptyMessage)) > 0 Then
                    SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_Index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal InvSlot As Byte, ByVal Amount As Long, Optional ByVal Durability As Integer = 0)
    Dim BankSlot As Long
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, InvSlot))
        
    If BankSlot > 0 And BankSlot <= MAX_BANK Then
        If Item(GetPlayerInvItemNum(index, InvSlot)).stackable = 1 Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, InvSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), Amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, InvSlot))
                Call SetPlayerBankItemValue(index, BankSlot, Amount)
                Call SetPlayerBankItemBind(index, BankSlot, GetPlayerInvItemBind(index, InvSlot))
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, InvSlot) And Not Item(GetPlayerInvItemNum(index, InvSlot)).Type = ITEM_TYPE_EQUIPMENT Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, InvSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call SetPlayerBankItemBind(index, BankSlot, GetPlayerInvItemBind(index, InvSlot))
                Call SetPlayerBankItemDur(index, BankSlot, Durability)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), 0)
            End If
        End If
    End If
    
    SendBank index
End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Byte, ByVal Amount As Long)
    Dim InvSlot As Long

    If BankSlot < 1 Or BankSlot > MAX_BANK Then Exit Sub
    If GetPlayerBankItemNum(index, BankSlot) < 1 Or GetPlayerBankItemNum(index, BankSlot) > MAX_ITEMS Then Exit Sub
    
    ' Hack prevention
    If Item(GetPlayerBankItemNum(index, BankSlot)).stackable = 1 Then
        If GetPlayerBankItemValue(index, BankSlot) < Amount Then Amount = GetPlayerBankItemValue(index, BankSlot)
        If Amount < 1 Then Exit Sub
    Else
        If Not Amount = 1 Then Exit Sub
    End If
    
    InvSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If InvSlot > 0 And InvSlot <= MAX_ITEMS Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).stackable = 1 Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), Amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - Amount)
            
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
                Call SetPlayerBankItemBind(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0, GetPlayerBankItemDur(index, BankSlot), GetPlayerBankItemBind(index, BankSlot))
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
                Call SetPlayerBankItemDur(index, BankSlot, 0)
                Call SetPlayerBankItemBind(index, BankSlot, 0)
            End If
        End If
    End If
    
    SendBank index
End Sub

Function TakeBankSlot(ByVal index As Long, ByVal ItemNum As Integer, ByVal ItemVal As Long) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

    For i = 1 To MAX_BANK
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(index, i) = ItemNum Then
            If Item(ItemNum).stackable = 1 Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerBankItemValue(index, i) Then
                    TakeBankSlot = True
                Else
                    Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) - ItemVal)
                    Exit Function
                End If
            Else
                TakeBankSlot = True
            End If

            If TakeBankSlot Then
                Call SetPlayerBankItemNum(index, i, 0)
                Call SetPlayerBankItemValue(index, i, 0)
                Call SetPlayerBankItemDur(index, i, 0)
                Call SetPlayerBankItemBind(index, i, 0)
                Exit For
            End If
        End If
    Next
    
    SendBank index
End Function

Public Sub UseItem(ByVal index As Long, ByVal InvNum As Byte)
    Dim n As Long, i As Long, x As Long, Y As Long, TotalPoints As Integer, EquipSlot As Byte
    Dim Item1 As Long
    Dim Item2 As Long
    Dim Result As Long
    Dim Skill As Byte
    Dim SkillExp As Integer
    Dim SkillLevelReq As Byte
    Dim ToolReq As Long
    Dim ItemNum As Integer
    
    ItemNum = GetPlayerInvItemNum(index, InvNum)
    
    ' Check subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' Check if they can use the item
    If Not CanPlayerUseItem(index, ItemNum) Then Exit Sub
    
    n = Item(ItemNum).Data2

    ' Set the bind
    If Item(ItemNum).Type = ITEM_TYPE_EQUIPMENT Then
        If Item(ItemNum).BindType = BIND_ON_EQUIP Then
            Call SetPlayerInvItemBind(index, InvNum, BIND_ON_PICKUP)
        End If
    End If
            
    ' Find out what kind of item it is
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_EQUIPMENT
            EquipSlot = Item(ItemNum).EquipSlot
            
            If EquipSlot >= 1 And EquipSlot <= Equipment_Count - 1 Then
                If Item(ItemNum).Indestructable = 1 Or GetPlayerInvItemDur(index, InvNum) > 0 Then
                    Call PlayerUnequipItem(index, EquipSlot, False, False, True)
                    
                    PlayerMsg index, "You equip " & CheckGrammar(Trim$(Item(ItemNum).Name)) & ".", BrightGreen
                    SetPlayerEquipment index, ItemNum, EquipSlot
                    SetPlayerEquipmentDur index, GetPlayerInvItemDur(index, InvNum), EquipSlot
                    SetPlayerEquipmentBind index, GetPlayerInvItemBind(index, InvNum), EquipSlot
                    TakeInvSlot index, InvNum, 0, True
                    
                    ' Send update
                    SendInventoryUpdate index, InvNum
                    Call SendWornEquipment(index)
                    Call SendMapEquipment(index)
                    SendPlayerStats index
                    
                    ' Send vitals
                    For i = 1 To Vitals.Vital_Count - 1
                        Call SendVital(index, i)
                    Next
                    
                    ' Send vitals to party if in one
                    If tempplayer(index).InParty > 0 Then SendPartyVitals tempplayer(index).InParty, index
                    
                     ' Send the sound
                    SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
                Else
                    If Item(ItemNum).Data1 = 0 Then
                        Call PlayerMsg(index, "This item lacks durability, report it to a staff member!", 12)
                    Else
                        Call PlayerMsg(index, "The item you are trying to equip is broken!", 12)
                    End If
                End If
            End If
        
        Case ITEM_TYPE_CONSUME
            If GetPlayerLevel(index) = Options.MaxLevel And Item(ItemNum).AddEXP > 0 Then
                Call PlayerMsg(index, "You can't use items which modify your experience when your at the max level!", BrightRed)
                Exit Sub
            End If
            
            ' Add HP
            If Item(ItemNum).AddHP > 0 Then
                If Not GetPlayerVital(index, HP) = GetPlayerMaxVital(index, HP) Then
                    If tempplayer(index).VitalPotionTimer(HP) > timeGetTime Then
                        Call PlayerMsg(index, "You must wait before you can use another potion that modifies your health!", BrightRed)
                        Exit Sub
                    Else
                        If Item(ItemNum).HoT = 1 Then
                            tempplayer(index).VitalCycle(HP) = Item(ItemNum).Data1
                            tempplayer(index).VitalPotion(HP) = ItemNum
                            tempplayer(index).VitalPotionTimer(HP) = timeGetTime + (Item(ItemNum).Data1 * 1000)
                        Else
                            Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.HP) = Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.HP) + Item(ItemNum).AddHP
                            SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                            SendVital index, HP
                            tempplayer(index).VitalPotionTimer(HP) = timeGetTime + PotionWaitTimer
                        End If
                    End If
                Else
                    Call PlayerMsg(index, "Using this item will have no effect!", BrightRed)
                    Exit Sub
                End If
            End If
            
            ' Add MP
            If Item(ItemNum).AddMP > 0 Then
                If Not GetPlayerVital(index, MP) = GetPlayerMaxVital(index, MP) Then
                    If tempplayer(index).VitalPotionTimer(MP) > timeGetTime And Item(ItemNum).AddHP < 1 Then
                        Call PlayerMsg(index, "You must wait before you can use another potion that modifies your mana!", BrightRed)
                        Exit Sub
                    Else
                        If Item(ItemNum).HoT = 1 Then
                            tempplayer(index).VitalCycle(MP) = Item(ItemNum).Data1
                            tempplayer(index).VitalPotion(MP) = ItemNum
                            tempplayer(index).VitalPotionTimer(MP) = timeGetTime + (Item(ItemNum).Data1 * 1000)
                        Else
                            Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.MP) = Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.MP) + Item(ItemNum).AddMP
                            SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                            SendVital index, MP
                            tempplayer(index).VitalPotionTimer(MP) = timeGetTime + PotionWaitTimer
                        End If
                    End If
                Else
                    Call PlayerMsg(index, "Using this item will have no effect!", BrightRed)
                    Exit Sub
                End If
            End If
            
            ' Add exp
            If Item(ItemNum).AddEXP > 0 Then
                SetPlayerExp index, GetPlayerExp(index) + Item(ItemNum).AddEXP
                SendPlayerExp index
                CheckPlayerLevelUp index
                SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddEXP & " Exp", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
            End If
            
            Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, GetPlayerX(index), GetPlayerY(index))
            
            ' Send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            
            ' Is it reusable, if not take the item away
            If Item(ItemNum).IsReusable = False Then
                Call TakeInvSlot(index, InvNum, 1)
            End If
        
        Case ITEM_TYPE_SPELL
            ' Get the spell number
            n = Item(ItemNum).Data1

            If n > 0 Then
                i = FindOpenSpellSlot(index)

                ' Make sure they have an open spell slot
                If i > 0 Then
                    ' Make sure they don't already have the spell
                    If Not HasSpell(index, n) Then
                        ' Make sure it's a valid name and their is an icon
                        If Not Trim$(Spell(n).Name) = vbNullString And Not Spell(n).Icon = 0 Then
                            ' Send the sound
                            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
                            Call SetPlayerSpell(index, i, n)
                            Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, GetPlayerX(index), GetPlayerY(index))
                            Call TakeInvSlot(index, InvNum, 1)
                            Call PlayerMsg(index, "You have learned a new spell!", BrightGreen)
                            Call SendPlayerSpell(index, i)
                        Else
                            Call PlayerMsg(index, "This spell either does not have a name or icon, report this to a staff member.", BrightRed)
                            Exit Sub
                        End If
                    Else
                        Call PlayerMsg(index, "You have already learned this spell!", BrightRed)
                        Exit Sub
                    End If
                Else
                    Call PlayerMsg(index, "You have learned all that you can learn!", BrightRed)
                    Exit Sub
                End If
            Else
                Call PlayerMsg(index, "This item does not have a spell, please inform a staff member!", BrightRed)
                Exit Sub
            End If
        
        Case ITEM_TYPE_TELEPORT
            If Moral(Map(GetPlayerMap(index)).Moral).CanPK = 1 Then
                Call PlayerMsg(index, "You can't teleport while in a PvP area!", BrightRed)
                Exit Sub
            End If
            
            Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, GetPlayerX(index), GetPlayerY(index))
            Call PlayerWarp(index, Item(ItemNum).Data1, Item(ItemNum).Data2, Item(ItemNum).Data3)
            
            ' Send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            
            ' Is it reusable, if not take item away
            If Item(ItemNum).IsReusable = False Then
                Call TakeInvSlot(index, InvNum, 1)
            End If
            
        Case ITEM_TYPE_RESETSTATS
            TotalPoints = GetPlayerPoints(index)
            
            For i = 1 To Stats.Stat_count - 1
                TotalPoints = TotalPoints + (GetPlayerRawStat(index, i) - Class(GetPlayerClass(index)).Stat(i))
                Call SetPlayerStat(index, i, Class(GetPlayerClass(index)).Stat(i))
            Next
            
            ' Send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            
            Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, GetPlayerX(index), GetPlayerY(index))
            Call SetPlayerPoints(index, TotalPoints)
            Call SendPlayerStats(index)
            Call SendPlayerPoints(index)
            Call PlayerMsg(index, "Your stats have been reset!", Yellow)
            Call TakeInvSlot(index, InvNum, 1)

        Case ITEM_TYPE_SPRITE
            Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, GetPlayerX(index), GetPlayerY(index))
            Call SetPlayerSprite(index, Item(ItemNum).Data1)
            Call SendPlayerSprite(index)
            
            ' Send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
        
            ' Is it reusable, if not take item away
            If Item(ItemNum).IsReusable = False Then
                Call TakeInvSlot(index, InvNum, 1)
            End If
            
        Case ITEM_TYPE_TITLE
            Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, GetPlayerX(index), GetPlayerY(index))
            Call AddPlayerTitle(index, Item(ItemNum).Data1, InvNum)
            
            ' Send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
        
        Case ITEM_TYPE_RECIPE
            ' Get the recipe information
            Item1 = Item(ItemNum).Data1
            Item2 = Item(ItemNum).Data2
            Result = Item(ItemNum).Data3
            Skill = Item(ItemNum).Skill
            SkillExp = Item(ItemNum).SkillExp
            SkillLevelReq = Item(ItemNum).SkillLevelReq
            ToolReq = Item(ItemNum).ToolRequired
            
            ' Perform Recipe checks
            If Item1 <= 0 Or Item2 <= 0 Or Result <= 0 Or Skill <= 0 Then
                Call PlayerMsg(index, "This is an incomplete recipe...", BrightRed)
                Exit Sub
            End If
            
            If GetPlayerEquipment(index, Weapon) <> ToolReq And HasItem(index, ToolReq) = 0 And ToolReq <> 0 Then
                Call PlayerMsg(index, "You don't have the proper tool required to craft this item!", BrightRed)
                Exit Sub
            End If
            
            If GetPlayerSkill(index, Skill) < SkillLevelReq Then
                Call PlayerMsg(index, "Your " & GetSkillName(Skill) & " skill isn't high enough to craft this item (" & SkillLevelReq & ")!", BrightRed)
                Exit Sub
            End If
            
            ' Give the resulting item
            If HasItem(index, Item1) Then
                If HasItem(index, Item2) Then
                    Call TakeInvSlot(index, Item1, 1)
                    Call TakeInvSlot(index, Item2, 1)
                    Call GiveInvItem(index, Result, 1)
                    Call PlayerMsg(index, "You have successfully created " & Trim$(Item(Result).Name) & " and earned " & SkillExp & " experience for the skill " & GetSkillName(Skill) & ".", BrightGreen)
                    
                    If GetPlayerSkill(index, Skill) < Options.MaxLevel Then
                        ' Add the experience to the skill
                        Call SetPlayerSkillExp(index, GetPlayerSkillExp(index, Skill) + SkillExp, Skill)
                        
                        ' Check for skill level up
                        Call CheckPlayerSkillLevelUp(index, Skill)
                    End If
                    
                    Call SendPlayerData(index)
                Else
                    Call PlayerMsg(index, "You do not have all of the ingredients!", BrightRed)
                    Exit Sub
                End If
            Else
                Call PlayerMsg(index, "You do not have all of the ingredients!", BrightRed)
                Exit Sub
            End If
    End Select
End Sub

Public Sub SetCheckpoint(ByVal index As Long, ByVal MapNum As Integer, ByVal x As Long, ByVal Y As Long)
    ' Check if their checkpoint is already set here
    If Account(index).Chars(GetPlayerChar(index)).CheckPointMap = MapNum And Account(index).Chars(GetPlayerChar(index)).CheckPointX = x And Account(index).Chars(GetPlayerChar(index)).CheckPointY = Y Then
        Call PlayerMsg(index, "Your checkpoint is already saved here!", BrightRed)
        Exit Sub
    End If
   
    PlayerMsg index, "Your checkpoint has been saved.", BrightGreen
    
    ' Save the Checkpoint
    Account(index).Chars(GetPlayerChar(index)).CheckPointMap = MapNum
    Account(index).Chars(GetPlayerChar(index)).CheckPointX = x
    Account(index).Chars(GetPlayerChar(index)).CheckPointY = Y
End Sub

Public Sub UpdatePlayerEquipmentItems(ByVal index As Long)
    Dim i As Long
    
    If GetPlayerEquipment(index, Shield) > 0 And GetPlayerEquipment(index, Weapon) > 0 Then
        If Item(GetPlayerEquipment(index, Weapon)).TwoHanded = 1 Then
            Call PlayerUnequipItem(index, Weapon, True, True, True)
        End If
    End If
    
    For i = 1 To Equipment_Count - 1
        If GetPlayerEquipment(index, i) > 0 Then
            If Item(GetPlayerEquipment(index, i)).EquipSlot <> i Then
                Call PlayerUnequipItem(index, i, True, True, True)
            End If
        End If
    Next
End Sub

Public Sub UpdateAllPlayerEquipmentItems()
    Dim n As Long, i As Long
    
    For n = 1 To Player_HighIndex
        If IsPlaying(n) Then
            If GetPlayerEquipment(n, Shield) > 0 And GetPlayerEquipment(n, Weapon) > 0 Then
                If Item(GetPlayerEquipment(n, Weapon)).TwoHanded = 1 Then
                    Call PlayerUnequipItem(n, Weapon, True, True, True)
                End If
            End If
            
            For i = 1 To Equipment_Count - 1
                If GetPlayerEquipment(n, i) > 0 Then
                    If Item(GetPlayerEquipment(n, i)).EquipSlot <> i Then
                        Call PlayerUnequipItem(n, i, True, True, True)
                    End If
                End If
            Next
        End If
    Next
End Sub

Public Sub UpdatePlayerItems(ByVal index As Long)
    Dim TmpItem As Long, TmpAmount As Long
    Dim i As Byte, InvAmount As Long, BankAmount As Long, x As Long

    ' Make sure the inventory items are not cached as a currency
    For i = 1 To MAX_INV
        TmpItem = GetPlayerInvItemNum(index, i)
        InvAmount = CheckInventorySlots(index, TmpItem)
        
        If TmpItem > 0 And TmpItem <= MAX_ITEMS Then
            If GetPlayerInvItemValue(index, i) > 1 And Item(TmpItem).stackable = 0 Then
                TmpAmount = GetPlayerInvItemValue(index, i)
                Call TakeInvSlot(index, i, GetPlayerInvItemValue(index, i))
                
                For x = 1 To TmpAmount
                    Call GiveInvItem(index, TmpItem, 1)
                Next
            End If
            
            If GetPlayerInvItemValue(index, i) = 0 And Item(TmpItem).Type <> ITEM_TYPE_EQUIPMENT Then
                Call TakeInvSlot(index, i, 0)
                Call GiveInvItem(index, TmpItem, 1)
            ElseIf InvAmount > 1 And Item(TmpItem).stackable = 1 And GetPlayerInvItemValue(index, i) <= 1 Then
                Call TakeInvSlot(index, i, 1)
                Call GiveInvItem(index, TmpItem, 1)
            End If
        End If
    Next
    
    ' Make sure the Bank items are not cached as a currency
   ' For i = 1 To MAX_BANK
   '     TmpItem = GetPlayerBankItemNum(index, i)
   '     BankAmount = CheckBankSlots(index, TmpItem)
        
   '     If TmpItem > 0 And TmpItem <= MAX_ITEMS Then
   '         If GetPlayerBankItemValue(index, i) > 1 And Item(TmpItem).Stackable = 0 Then
   '             TmpAmount = GetPlayerBankItemValue(index, i)
   '             Call TakeBankSlot(index, i, GetPlayerBankItemValue(index, i))
                
   '             For X = 1 To TmpAmount
   '                 Call GiveBankItem(index, TmpItem, 1)
   '             Next
   '         End If
            
   '         If GetPlayerBankItemValue(index, i) = 0 And Item(TmpItem).Type <> ITEM_TYPE_EQUIPMENT Then
   '             Call TakeBankSlot(index, i, 0)
   '             Call GiveBankItem(index, TmpItem, 1)
   '         ElseIf BankAmount > 1 And Item(TmpItem).Stackable = 1 And GetPlayerBankItemValue(index, i) <= 1 Then
   '             Call TakeBankSlot(index, i, 1)
   '             Call GiveBankItem(index, TmpItem, 1)
   '         End If
   '     End If
   ' Next
End Sub

Public Sub UpdateAllPlayerItems(ByVal ItemNum As Integer)
    Dim TmpItem As Long
    Dim n As Long, i As Byte, x As Byte

    For n = 1 To Player_HighIndex
        If IsPlaying(n) Then
            UpdatePlayerItems n
        End If
    Next
End Sub

Public Sub UpdateClassData(ByVal index As Long)
     Dim i As Long
    Dim TotalPoints As Long
    Dim TotalPoints2 As Long
    
    If GetPlayerAccess(index) > STAFF_MODERATOR Then Exit Sub
    
    For i = 1 To Stats.Stat_count - 1
        TotalPoints = TotalPoints + Class(GetPlayerClass(index)).Stat(i)
        TotalPoints2 = TotalPoints2 + GetPlayerRawStat(index, i)
    Next

    TotalPoints = TotalPoints + ((GetPlayerLevel(index) - 1) * Options.StatsLevel)
    TotalPoints2 = TotalPoints2 + GetPlayerPoints(index)

    ' Verify incorrect class data
    If TotalPoints <> TotalPoints2 Then

        For i = 1 To Stats.Stat_count - 1
            Call SetPlayerStat(index, i, Class(GetPlayerClass(index)).Stat(i))
        Next

        Call SetPlayerPoints(index, (GetPlayerLevel(index) - 1) * Options.StatsLevel)
    End If

   If GetPlayerSprite(index) = 0 Then
        If GetPlayerGender(index) = GENDER_MALE Then
            Call SetPlayerSprite(index, Class(GetPlayerClass(index)).MaleSprite)
        Else
            Call SetPlayerSprite(index, Class(GetPlayerClass(index)).FemaleSprite)
        End If
        
        ' Sprite still nothing?
        If GetPlayerSprite(index) = 0 Then
            Call SetPlayerSprite(index, 1)
        End If
    End If
    
    If GetPlayerFace(index) = 0 Then
        If GetPlayerGender(index) = GENDER_MALE Then
            Call SetPlayerFace(index, Class(GetPlayerClass(index)).MaleFace)
        Else
            Call SetPlayerFace(index, Class(GetPlayerClass(index)).FemaleFace)
        End If
        
        ' Face still nothing?
        If GetPlayerFace(index) = 0 Then
            Call SetPlayerFace(index, 1)
        End If
    End If
End Sub

Public Sub UpdateAllClassData()
    Dim i As Long, x As Long
    
    For x = 1 To Player_HighIndex
        If GetPlayerAccess(x) <= STAFF_MODERATOR Then
            ' Verify incorrect class data
            For i = 1 To Stats.Stat_count - 1
                If Not Class(GetPlayerClass(x)).Stat(i) = GetPlayerStat(x, i) - ((GetPlayerLevel(x) - 1) * 5) Then
                    Call SetPlayerStat(x, i, Class(GetPlayerClass(x)).Stat(i) + ((GetPlayerLevel(x) - 1) * 5))
                End If
            Next
        
            If GetPlayerSprite(x) = 0 Then
                If GetPlayerGender(x) = GENDER_MALE Then
                    Call SetPlayerSprite(x, Class(GetPlayerClass(x)).MaleSprite)
                Else
                    Call SetPlayerSprite(x, Class(GetPlayerClass(x)).FemaleSprite)
                End If
                
                ' Sprite still nothing?
                If GetPlayerSprite(x) = 0 Then
                    Call SetPlayerSprite(x, 1)
                End If
            End If
            
            If GetPlayerFace(x) = 0 Then
                If GetPlayerGender(x) = GENDER_MALE Then
                    Call SetPlayerFace(x, Class(GetPlayerClass(x)).MaleFace)
                Else
                    Call SetPlayerFace(x, Class(GetPlayerClass(x)).FemaleFace)
                End If
                
                ' Face still nothing?
                If GetPlayerFace(x) = 0 Then
                    Call SetPlayerFace(x, 1)
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerTrade(ByVal index As Long, ByVal TradeTarget As Long) As Boolean
    Dim sX As Long, sY As Long, tX As Long, tY As Long
    
    ' Can't trade with yourself
    If TradeTarget = index Then
        PlayerMsg index, "You can't trade with yourself.", BrightRed
        Exit Function
    End If
    
    ' Make sure they're on the same map
    If Not Account(TradeTarget).Chars(GetPlayerChar(TradeTarget)).Map = Account(index).Chars(GetPlayerChar(index)).Map Then Exit Function
    
    ' Make sure they are allowed to trade
    If Account(TradeTarget).Chars(GetPlayerChar(index)).CanTrade = False Then
        PlayerMsg index, Trim$(GetPlayerName(TradeTarget)) & " has their trading turned off.", BrightRed
        Exit Function
    End If

    ' Make sure they're stood next to each other
    tX = Account(TradeTarget).Chars(GetPlayerChar(TradeTarget)).x
    tY = Account(TradeTarget).Chars(GetPlayerChar(TradeTarget)).Y
    sX = Account(index).Chars(GetPlayerChar(index)).x
    sY = Account(index).Chars(GetPlayerChar(index)).Y
    
    ' Within range?
    If tX < sX - 1 Or tX > sX + 1 And tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request or accept a trade.", BrightRed
        Exit Function
    End If
    
    CanPlayerTrade = True
End Function

Function CanPlayerUseItem(ByVal index As Long, ByVal ItemNum As Long, Optional Message As Boolean = True) As Boolean
    Dim LevelReq As Byte
    Dim AccessReq As Byte
    Dim ClassReq As Byte
    Dim GenderReq As Byte
    Dim i As Long

    ' Can't use items while in a map that doesn't allow it
    If Moral(Map(GetPlayerMap(index)).Moral).CanUseItem = 0 Then
        Call PlayerMsg(index, "You can't use items here!", BrightRed)
        Exit Function
    End If
    
    LevelReq = Item(ItemNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        If Message Then
            Call PlayerMsg(index, "You must be level " & LevelReq & " to use this item.", BrightRed)
        End If
        Exit Function
    End If
    
    AccessReq = Item(ItemNum).AccessReq
    
    ' Make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        If Message Then
            Call PlayerMsg(index, "You must be a staff member to use this item.", BrightRed)
        End If
        Exit Function
    End If
    
    ClassReq = Item(ItemNum).ClassReq
    
    ' Make sure the Classes req > 0
    If ClassReq > 0 Then ' 0 = no req
        If Not ClassReq = GetPlayerClass(index) Then
            If Message Then
                Call PlayerMsg(index, "You must be " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this item!", BrightRed)
            End If
            Exit Function
        End If
    End If
    
    GenderReq = Item(ItemNum).GenderReq
    
    ' Make sure the Gender req > 0
    If GenderReq > 0 Then ' 0 = no req
        If Not GenderReq - 1 = GetPlayerGender(index) Then
            If Message Then
                If GetPlayerGender(index) = 0 Then
                    Call PlayerMsg(index, "You need to be a female to use this item!", BrightRed)
                Else
                    Call PlayerMsg(index, "You need to be a male to use this item!", BrightRed)
                End If
            End If
            Exit Function
        End If
    End If
    
    ' Check if they have the stats required to use this item
    For i = 1 To Stats.Stat_count - 1
        If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
            If Message Then
                PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
            End If
            Exit Function
        End If
    Next
    
    ' Check if they have the proficiency required to use this item
    If Item(ItemNum).ProficiencyReq > 0 Then
        If GetPlayerProficiency(index, Item(ItemNum).ProficiencyReq) = 0 Then
            If Message Then
                PlayerMsg index, "You lack the proficiency to use this item!", BrightRed
            End If
            Exit Function
        End If
    End If
    
    ' Don't let them equip a two handed weapon if they have a shield on
     If Item(ItemNum).TwoHanded = 1 Then
        If GetPlayerEquipment(index, Shield) > 0 Then
            PlayerMsg index, "You must unequip your shield before equipping a two-handed weapon!", BrightRed
            Exit Function
        End If
    End If
    
    ' Don't let them use a tool they don't meet the level requirement to
    If Item(ItemNum).SkillReq > 0 Then
        If GetPlayerSkill(index, Item(ItemNum).SkillReq) < Item(ItemNum).LevelReq Then
            PlayerMsg index, "Your " & CheckGrammar(GetSkillName(Item(ItemNum).SkillReq)) & " skill level does not meet the requirement to use this item!", BrightRed
            Exit Function
        End If
    End If
    
    CanPlayerUseItem = True
End Function

Public Function CanPlayerCastSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < Spell(SpellNum).MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Function
    End If
    
    ' Make sure they are the right level
    If Spell(SpellNum).LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & Spell(SpellNum).LevelReq & " to cast this spell.", BrightRed)
        Exit Function
    End If
    
    ' Make sure they have the right access
    If Spell(SpellNum).AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be a staff member to cast this spell.", BrightRed)
        Exit Function
    End If
    
    ' Make sure the ClassReq > 0
    If Spell(SpellNum).ClassReq > 0 Then ' 0 = no req
        If Spell(SpellNum).ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(Spell(SpellNum).ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Function
        End If
    End If
    
    CanPlayerCastSpell = True
End Function

Public Sub DamagePlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Byte)
    Dim ItemNum As Long, RandomNum As Byte
    
    ItemNum = GetPlayerEquipment(index, EquipmentSlot)
    
    If ItemNum = 0 Then Exit Sub
    
    ' Make sure the item isn't indestructable
    If Item(ItemNum).Indestructable = 1 Then Exit Sub
    
    ' Don't subtract past 0
    If GetPlayerEquipmentDur(index, EquipmentSlot) = 0 Then Exit Sub
    
    RandomNum = Random(1, 7)
    
    ' 1 in 7 chance it will actually damage the equipment if it's not a shield type item
    If RandomNum = 1 Or EquipmentSlot = Shield Then
        If Item(ItemNum).Type = ITEM_TYPE_EQUIPMENT Then
        
            ' Take away 1 durability
            Call SetPlayerEquipmentDur(index, GetPlayerEquipmentDur(index, EquipmentSlot) - 1, EquipmentSlot)
            Call SendWornEquipment(index)
            Call SendMapEquipment(index)
                
            If GetPlayerEquipmentDur(index, EquipmentSlot) < 1 Then
                Call PlayerMsg(index, "Your " & Trim$(Item(ItemNum).Name) & " has broken.", BrightRed)
                Call PlayerUnequipItem(index, EquipmentSlot, True, True, True)
            ElseIf GetPlayerEquipmentDur(index, EquipmentSlot) = 10 Then
                Call PlayerMsg(index, "Your " & Trim$(Item(ItemNum).Name) & " is about to break!", BrightRed)
            End If
        End If
    End If
End Sub

Public Sub WarpToCheckPoint(index As Long)
    Dim MapNum As Integer
    Dim x As Long, Y As Long
    
    MapNum = Account(index).Chars(GetPlayerChar(index)).CheckPointMap
    x = Account(index).Chars(GetPlayerChar(index)).CheckPointX
    Y = Account(index).Chars(GetPlayerChar(index)).CheckPointY
    
    PlayerWarp index, MapNum, x, Y
End Sub

Function IsAFriend(ByVal index As Long, ByVal OtherPlayer As Long) As Boolean
    Dim i As Long
    
    ' Are they on the user's friend list
    For i = 1 To Account(OtherPlayer).Friends.AmountOfFriends
        If Trim$(Account(OtherPlayer).Friends.Members(i)) = GetPlayerName(index) Then
            IsAFriend = True
            Exit Function
        End If
    Next
End Function

Function IsAFoe(ByVal index As Long, ByVal OtherPlayer As Long) As Boolean
    Dim i As Long
    
    ' Are they on the user's foe list
    For i = 1 To Account(OtherPlayer).Foes.Amount
        If Trim$(Account(OtherPlayer).Foes.Members(i)) = GetPlayerName(index) Then
            Call PlayerMsg(index, "You are being ignored by " & GetPlayerName(OtherPlayer) & "!", BrightRed)
            IsAFoe = True
            Exit Function
        End If
    Next
End Function

Function IsPlayerBusy(ByVal index As Long, ByVal OtherPlayer As Long) As Boolean
    ' Make sure they're not busy doing something else
    If IsPlaying(OtherPlayer) Then
        If tempplayer(OtherPlayer).InBank Or tempplayer(OtherPlayer).InShop > 0 Or tempplayer(OtherPlayer).InTrade > 0 Or (tempplayer(OtherPlayer).PartyInvite > 0 And tempplayer(OtherPlayer).PartyInvite <> index) Or (tempplayer(OtherPlayer).TradeRequest > 0 And tempplayer(OtherPlayer).TradeRequest <> index) Or (tempplayer(OtherPlayer).GuildInvite > 0 And tempplayer(OtherPlayer).GuildInvite <> index) Then
            IsPlayerBusy = True
            PlayerMsg index, GetPlayerName(OtherPlayer) & " is busy!", BrightRed
            Exit Function
        End If
    End If
End Function

