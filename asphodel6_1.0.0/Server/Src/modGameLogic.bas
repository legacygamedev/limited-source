Attribute VB_Name = "modGameLogic"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

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

Public Function IsWithinPVPLimit(ByVal Index As Long, ByVal Target As Long) As Boolean

    If frmServer.chkPVPLevel.Value <> vbChecked Then
        IsWithinPVPLimit = True
        Exit Function
    End If
    
    If GetPlayerLevel(Index) < frmServer.scrlLevelLimit.Value Then
        PlayerMsg Index, "You must be level " & frmServer.scrlLevelLimit.Value & " to be able to attack another player!", Color.BrightRed
        Exit Function
    End If
    
    If GetPlayerLevel(Target) < frmServer.scrlLevelLimit.Value Then
        PlayerMsg Index, "You cannot attack " & GetPlayerName(Target) & "!", Color.BrightRed
        Exit Function
    End If
    
    IsWithinPVPLimit = True
    
End Function

Public Sub AdjustVitalBonus(ByVal Index As Long, ByVal Vital As Vitals, ByVal Amount As Long, Optional ByVal AddorSubtract As Boolean = True, Optional ByVal Clear As Boolean = False)

    If Clear Then
        VitalBonus(Index, Vital) = 0
        GoTo Skipper
    End If
    
    Select Case AddorSubtract
        Case True
            VitalBonus(Index, Vital) = VitalBonus(Index, Vital) + Amount
        Case False
            VitalBonus(Index, Vital) = VitalBonus(Index, Vital) - Amount
    End Select
    
Skipper:
    
    If GetPlayerVital_withBonus(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then SetPlayerVital Index, Vital, GetPlayerMaxVital(Index, Vital): UpdatePlayerVital Index
    
    SendVital Index, Vital
    
End Sub

Public Sub AdjustStatBonus(ByVal Index As Long, ByVal Stat As Stats, ByVal Amount As Long, Optional ByVal AddorSubtract As Boolean = True, Optional ByVal Clear As Boolean = False)

    If Clear Then
        StatBonus(Index, Stat) = 0
        GoTo Skipper
    End If
    
    Select Case AddorSubtract
        Case True
            StatBonus(Index, Stat) = StatBonus(Index, Stat) + Amount
        Case False
            StatBonus(Index, Stat) = StatBonus(Index, Stat) - Amount
    End Select
    
Skipper:
    
    SendStats Index
    
End Sub

Function CanMakeGuild(ByVal Index As Long, ByVal SpeakerName As String) As Boolean
Dim LoopI As Long
Dim Speaker As String
Dim UseQuote As String

    CanMakeGuild = True
    
    If LenB(Trim$(SpeakerName)) > 0 Then
        UseQuote = vbQuote
        Speaker = SpeakerName & " says, " & UseQuote
    End If
    
    If FindOpenGuildSlot = 0 Then
        PlayerMsg Index, Speaker & "Sorry, but there isn't anymore room for guilds right now!" & UseQuote, Color.BrightRed
        CanMakeGuild = False
        Exit Function
    End If
    
    If Player(Index).Char(TempPlayer(Index).CharNum).Guild > 0 Then
        PlayerMsg Index, Speaker & "You are already in a guild, you can't make another until you disband your current one!" & UseQuote, Color.BrightRed
        CanMakeGuild = False
        Exit Function
    End If
    
    If Guild_Creation_Item > 0 Then
        For LoopI = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, LoopI) > 0 Then
                If GetPlayerInvItemNum(Index, LoopI) = Guild_Creation_Item Then
                    Exit For
                End If
            End If
        Next
    End If
    
    If LoopI = MAX_INV + 1 Then
        PlayerMsg Index, Speaker & "Sorry, but to start a guild it costs " & Guild_Creation_Cost & " " & Trim$(Item(Guild_Creation_Item).Name) & "!" & UseQuote, Color.BrightRed
        CanMakeGuild = False
        Exit Function
    Else
        If Guild_Creation_Item > 0 Then
            If GetPlayerInvItemValue(Index, LoopI) < Guild_Creation_Cost Then
                PlayerMsg Index, Speaker & "Sorry, but to start a guild it costs " & Guild_Creation_Cost & " " & Trim$(Item(Guild_Creation_Item).Name) & "!" & UseQuote, Color.BrightRed
                CanMakeGuild = False
                Exit Function
            End If
        End If
    End If
    
End Function

Function FindOpenGuildMemberSlot(ByVal GuildNum As Long) As Long
Dim LoopI As Long

    For LoopI = 1 To UBound(Guild(GuildNum).Member_Account)
        If LenB(Trim$(Guild(GuildNum).Member_Account(LoopI))) < 1 Then
            FindOpenGuildMemberSlot = LoopI
            Exit Function
        End If
    Next
    
End Function

Function FindOpenGuildSlot() As Long
Dim LoopI As Long

    For LoopI = 1 To MAX_GUILDS
        If LenB(Trim$(Guild(LoopI).Name)) < 1 Then
            FindOpenGuildSlot = LoopI
            Exit Function
        End If
    Next
    
End Function

Public Sub CheckIfGuildStillExists(ByVal Index As Long)

    If Player(Index).Char(TempPlayer(Index).CharNum).Guild > 0 Then
        If LenB(Trim$(Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Name)) < 1 Then
            Player(Index).Char(TempPlayer(Index).CharNum).Guild = 0
            Player(Index).Char(TempPlayer(Index).CharNum).GuildRank = 0
        End If
    End If
    
End Sub

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
Dim i As Long

    FindOpenMapItemSlot = 0
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Function
    
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next
    
End Function

Function TotalOnlinePlayers() As Long
Dim i As Long

    TotalOnlinePlayers = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then TotalOnlinePlayers = TotalOnlinePlayers + 1
    Next
    
End Function

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    FindPlayer = 0

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
    
End Function

Public Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Durability, MapNum, X, Y)
    
End Sub

Public Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim i As Long

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    
    i = MapItemSlot
    
    If i <> 0 Then
        If ItemNum >= 0 Then
            If ItemNum <= MAX_ITEMS Then
            
                MapItem(MapNum, i).Num = ItemNum
                MapItem(MapNum, i).Value = ItemVal
                
                If ItemNum <> 0 Then
                    If (Item(ItemNum).Type >= ItemType.Weapon_) And (Item(ItemNum).Type <= ItemType.Shield_) Then
                        MapItem(MapNum, i).Dur = ItemDur
                    Else
                        MapItem(MapNum, i).Dur = 0
                    End If
                Else
                    MapItem(MapNum, i).Dur = 0
                End If
                
                MapItem(MapNum, i).X = X
                MapItem(MapNum, i).Y = Y
                MapItem(MapNum, i).Anim = Map(MapNum).Tile(X, Y).Data3
                
                Packet = SSpawnItem & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & MapItem(MapNum, i).Anim & SEP_CHAR & ItemNum & END_CHAR
                Call SendDataToMap(MapNum, Packet)
                
            End If
        End If
    End If
    
End Sub

Public Sub SpawnAllMapsItems()
Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next
    
End Sub

Public Sub SpawnMapItems(ByVal MapNum As Long)
Dim X As Long
Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    
    ' Spawn what we have
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(X, Y).Type = Tile_Type.Item_) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(X, Y).Data1).Type = ItemType.Currency_ And Map(MapNum).Tile(X, Y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, 1, MapNum, X, Y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, Map(MapNum).Tile(X, Y).Data2, MapNum, X, Y)
                End If
            End If
        Next
    Next
    
End Sub

Public Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Packet As String
Dim NpcNum As Long
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > UBound(MapSpawn(MapNum).Npc) Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    
    NpcNum = MapSpawn(MapNum).Npc(MapNpcNum).Num
    
    If NpcNum > 0 Then
        
        If MapSpawn(MapNum).Npc(MapNpcNum).X = -1 Then
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                X = Random(0, MAX_MAPX)
                Y = Random(0, MAX_MAPY)
               
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, X, Y) Then
                    MapNpc(MapNum).MapNpc(MapNpcNum).X = X
                    MapNpc(MapNum).MapNpc(MapNpcNum).Y = Y
                    Spawned = True
                    Exit For
                End If
            Next
            
            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then
                For X = 0 To MAX_MAPX
                    For Y = 0 To MAX_MAPY
                        If NpcTileIsOpen(MapNum, X, Y) Then
                            MapNpc(MapNum).MapNpc(MapNpcNum).X = X
                            MapNpc(MapNum).MapNpc(MapNpcNum).Y = Y
                            Spawned = True
                        End If
                    Next
                Next
            End If
        Else
            X = MapSpawn(MapNum).Npc(MapNpcNum).X
            Y = MapSpawn(MapNum).Npc(MapNpcNum).Y
            
            ' Check if the tile is walkable
            If NpcTileIsOpen(MapNum, X, Y) Then
                MapNpc(MapNum).MapNpc(MapNpcNum).X = X
                MapNpc(MapNum).MapNpc(MapNpcNum).Y = Y
                Spawned = True
            Else
                Spawned = False
            End If
        End If
        
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
        
            MapNpc(MapNum).MapNpc(MapNpcNum).Num = NpcNum
            MapNpc(MapNum).MapNpc(MapNpcNum).Target = 0
            MapNpc(MapNum).MapNpc(MapNpcNum).Dir = E_Direction.Down_
            MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
            MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(NpcNum, Vitals.MP)
            MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.SP) = GetNpcMaxVital(NpcNum, Vitals.SP)
            SendNPCVital MapNum, MapNpcNum
            
            Packet = SSpawnNpc & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Num & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).X & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Dir & END_CHAR
            Call SendDataToMap(MapNum, Packet)
        End If
        
    End If
    
End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
Dim LoopI As Long

    NpcTileIsOpen = True
    
    If PlayersOnMap(MapNum) Then
        For LoopI = 1 To MAX_PLAYERS
            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = X Then
                    If GetPlayerY(LoopI) = Y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
    
    For LoopI = 1 To UBound(MapSpawn(MapNum).Npc)
        If MapNpc(MapNum).MapNpc(LoopI).Num > 0 Then
            If MapNpc(MapNum).MapNpc(LoopI).X = X Then
                If MapNpc(MapNum).MapNpc(LoopI).Y = Y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If
    Next
    
    If Map(MapNum).Tile(X, Y).Type <> Tile_Type.None_ Then NpcTileIsOpen = False
    
End Function

Public Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To UBound(MapSpawn(MapNum).Npc)
        Call SpawnNpc(i, MapNum)
    Next
    
End Sub

Public Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next
    
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean

    ' Check attack timer
    If GetTickCountNew < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
    
    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function
    
    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function
   
    ' Check if at same coordinates
    Select Case GetPlayerDir(Attacker)
        Case E_Direction.Up_
            If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
        Case E_Direction.Down_
            If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
        Case E_Direction.Left_
            If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
        Case E_Direction.Right_
            If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
        Case Else
            Exit Function
    End Select
    
    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", Color.BrightRed)
            Exit Function
        End If
    End If
    
    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function
    
    If AdminSafety(Attacker, Victim) Then Exit Function
    
    ' Make sure attacker is high enough level
    If Not IsWithinPVPLimit(Attacker, Victim) Then Exit Function
    
    CanAttackPlayer = True
    
End Function

Function AdminSafety(ByVal Index As Long, ByVal Target As Long) As Boolean

    AdminSafety = True
    
    ' Check to make sure that they dont have access
    If GetPlayerAccess(Index) > StaffType.Monitor And frmServer.chkAdminSafety.Value Then
        Call PlayerMsg(Index, "You cannot attack any player as a staff member!", Color.BrightBlue)
        Exit Function
    End If
    
    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Target) > StaffType.Monitor And frmServer.chkAdminSafety.Value Then
        Call PlayerMsg(Index, "You are not allowed to attack staff members!", Color.BrightRed)
        Exit Function
    End If
    
    AdminSafety = False
    
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim MapNum As Long
Dim NpcNum As Long
Dim NpcX As Long
Dim NpcY As Long
    
    ' Check for subscript out of range
    If Not IsPlaying(Attacker) Or MapNpcNum <= 0 Or MapNpcNum > UBound(MapSpawn(GetPlayerMap(Attacker)).Npc) Then Exit Function
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).MapNpc(MapNpcNum).Num <= 0 Then Exit Function
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).MapNpc(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Function
    
    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If NpcNum > 0 Then
            If GetTickCountNew > TempPlayer(Attacker).AttackTimer + 1000 Then
                ' Check if at same coordinates
                Select Case GetPlayerDir(Attacker)
                    Case E_Direction.Up_
                        NpcX = MapNpc(MapNum).MapNpc(MapNpcNum).X
                        NpcY = MapNpc(MapNum).MapNpc(MapNpcNum).Y + 1
                    Case E_Direction.Down_
                        NpcX = MapNpc(MapNum).MapNpc(MapNpcNum).X
                        NpcY = MapNpc(MapNum).MapNpc(MapNpcNum).Y - 1
                    Case E_Direction.Left_
                        NpcX = MapNpc(MapNum).MapNpc(MapNpcNum).X + 1
                        NpcY = MapNpc(MapNum).MapNpc(MapNpcNum).Y
                    Case E_Direction.Right_
                        NpcX = MapNpc(MapNum).MapNpc(MapNpcNum).X - 1
                        NpcY = MapNpc(MapNum).MapNpc(MapNpcNum).Y
                End Select
                
                If NpcX = GetPlayerX(Attacker) Then
                    If NpcY = GetPlayerY(Attacker) Then
                        If Npc(NpcNum).Behavior <> NPC_Behavior.Friendly And Npc(NpcNum).Behavior <> NPC_Behavior.ShopKeeper Then
                            CanAttackNpc = True
                        Else
                            Call PlayerMsg(Attacker, "You cannot attack a " & Trim$(Npc(NpcNum).Name) & "!", BrightBlue)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long
Dim NpcNum As Long
    
    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > UBound(MapSpawn(GetPlayerMap(Index)).Npc) Or Not IsPlaying(Index) Then Exit Function
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).MapNpc(MapNpcNum).Num <= 0 Then Exit Function
    
    MapNum = GetPlayerMap(Index)
    NpcNum = MapNpc(MapNum).MapNpc(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Function
    
    ' Make sure npcs dont attack more then once a second
    If GetTickCountNew < MapNpc(MapNum).MapNpc(MapNpcNum).AttackTimer Then Exit Function
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then Exit Function
    
    MapNpc(MapNum).MapNpc(MapNpcNum).AttackTimer = GetTickCountNew + 1000
    
    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NpcNum > 0 Then
            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum).MapNpc(MapNpcNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).MapNpc(MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum).MapNpc(MapNpcNum).Y) And (GetPlayerX(Index) = MapNpc(MapNum).MapNpc(MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum).MapNpc(MapNpcNum).Y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum).MapNpc(MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum).MapNpc(MapNpcNum).Y) Then
                            If (GetPlayerX(Index) - 1 = MapNpc(MapNum).MapNpc(MapNpcNum).X) Then
                                CanNpcAttackPlayer = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal Reflection As Boolean = False)
Dim Name As String
Dim Exp As Long
Dim MapNum As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > UBound(MapSpawn(GetPlayerMap(Victim)).Npc) Or Not IsPlaying(Victim) Or Damage < 0 Then Exit Sub
    
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).MapNpc(MapNpcNum).Num <= 0 Then Exit Sub
    
    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(MapNpc(MapNum).MapNpc(MapNpcNum).Num).Name)
    
    ' Send this packet so they can see the person attacking
    Call SendDataToMap(MapNum, SNpcAttack & SEP_CHAR & MapNpcNum & END_CHAR)
    
    ' reduce dur. on victims equipment
    Call DamageEquipment(Victim, Armor)
    Call DamageEquipment(Victim, Helmet)
    
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
    
        If Not Reflection Then
            Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            Call PlayerMsg(Victim, "A " & Name & " reflected " & Damage & " points of damage!", BrightRed)
        End If
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by a " & Name, BrightRed)
                
        ' Calculate exp to give attacker
        Exp = Int(GetPlayerExp(Victim) * 0.33)
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then Exp = 0
        
        If Exp = 0 Then
            Call PlayerMsg(Victim, "You lost no experience points.", BrightRed)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            Call PlayerMsg(Victim, "You lost " & Exp & " experience points.", BrightRed)
        End If
        
        ' Set NPC target to 0
        MapNpc(MapNum).MapNpc(MapNpcNum).Target = 0
        
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        If Not Reflection Then
            Call PlayerMsg(Victim, "A " & Name & " hit you for " & Damage & " hit points.", BrightRed)
        Else
            Call PlayerMsg(Victim, "A " & Name & " reflected " & Damage & " points of damage!", BrightRed)
        End If
        
    End If
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
Dim i As Long
Dim n As Long
Dim X As Long
Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > UBound(MapSpawn(MapNum).Npc) Or Dir < E_Direction.Up_ Or Dir > E_Direction.Right_ Then Exit Function
    
    X = MapNpc(MapNum).MapNpc(MapNpcNum).X
    Y = MapNpc(MapNum).MapNpc(MapNpcNum).Y
    
    CanNpcMove = True
    
    Select Case Dir
        Case E_Direction.Up_
            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = Map(MapNum).Tile(X, Y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> Tile_Type.None_ Then
                    If n <> Tile_Type.Item_ Then
                        CanNpcMove = False
                        Exit Function
                    End If
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) Then
                            If (GetPlayerX(i) = MapNpc(MapNum).MapNpc(MapNpcNum).X) Then
                                If (GetPlayerY(i) = MapNpc(MapNum).MapNpc(MapNpcNum).Y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To UBound(MapSpawn(MapNum).Npc)
                    If (i <> MapNpcNum) Then
                        If (MapNpc(MapNum).MapNpc(i).Num > 0) Then
                            If (MapNpc(MapNum).MapNpc(i).X = MapNpc(MapNum).MapNpc(MapNpcNum).X) Then
                                If (MapNpc(MapNum).MapNpc(i).Y = MapNpc(MapNum).MapNpc(MapNpcNum).Y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
            Else
                CanNpcMove = False
            End If
                
        Case E_Direction.Down_
            ' Check to make sure not outside of boundries
            If Y < MAX_MAPY Then
                n = Map(MapNum).Tile(X, Y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> Tile_Type.None_ Then
                    If n <> Tile_Type.Item_ Then
                        CanNpcMove = False
                        Exit Function
                    End If
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) Then
                            If (GetPlayerX(i) = MapNpc(MapNum).MapNpc(MapNpcNum).X) Then
                                If (GetPlayerY(i) = MapNpc(MapNum).MapNpc(MapNpcNum).Y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To UBound(MapSpawn(MapNum).Npc)
                    If (i <> MapNpcNum) Then
                        If (MapNpc(MapNum).MapNpc(i).Num > 0) Then
                            If (MapNpc(MapNum).MapNpc(i).X = MapNpc(MapNum).MapNpc(MapNpcNum).X) Then
                                If (MapNpc(MapNum).MapNpc(i).Y = MapNpc(MapNum).MapNpc(MapNpcNum).Y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
            Else
                CanNpcMove = False
            End If
                
        Case E_Direction.Left_
            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> Tile_Type.None_ Then
                    If n <> Tile_Type.Item_ Then
                        CanNpcMove = False
                        Exit Function
                    End If
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) Then
                            If (GetPlayerX(i) = MapNpc(MapNum).MapNpc(MapNpcNum).X - 1) Then
                                If (GetPlayerY(i) = MapNpc(MapNum).MapNpc(MapNpcNum).Y) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To UBound(MapSpawn(MapNum).Npc)
                    If (i <> MapNpcNum) Then
                        If (MapNpc(MapNum).MapNpc(i).Num > 0) Then
                            If (MapNpc(MapNum).MapNpc(i).X = MapNpc(MapNum).MapNpc(MapNpcNum).X - 1) Then
                                If (MapNpc(MapNum).MapNpc(i).Y = MapNpc(MapNum).MapNpc(MapNpcNum).Y) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
            Else
                CanNpcMove = False
            End If
                
        Case E_Direction.Right_
            ' Check to make sure not outside of boundries
            If X < MAX_MAPX Then
                n = Map(MapNum).Tile(X + 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> Tile_Type.None_ Then
                    If n <> Tile_Type.Item_ Then
                        CanNpcMove = False
                        Exit Function
                    End If
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) Then
                            If (GetPlayerX(i) = MapNpc(MapNum).MapNpc(MapNpcNum).X + 1) Then
                                If (GetPlayerY(i) = MapNpc(MapNum).MapNpc(MapNpcNum).Y) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To UBound(MapSpawn(MapNum).Npc)
                    If (i <> MapNpcNum) Then
                        If (MapNpc(MapNum).MapNpc(i).Num > 0) Then
                            If (MapNpc(MapNum).MapNpc(i).X = MapNpc(MapNum).MapNpc(MapNpcNum).X + 1) Then
                                If (MapNpc(MapNum).MapNpc(i).Y = MapNpc(MapNum).MapNpc(MapNpcNum).Y) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Public Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Packet As String

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > UBound(MapSpawn(MapNum).Npc) Or Dir < E_Direction.Up_ Or Dir > E_Direction.Right_ Or Movement < 1 Or Movement > 2 Then Exit Sub
    
    MapNpc(MapNum).MapNpc(MapNpcNum).Dir = Dir
    
    Select Case Dir
        Case E_Direction.Up_
            MapNpc(MapNum).MapNpc(MapNpcNum).Y = MapNpc(MapNum).MapNpc(MapNpcNum).Y - 1
            Packet = SNpcMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).X & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case E_Direction.Down_
            MapNpc(MapNum).MapNpc(MapNpcNum).Y = MapNpc(MapNum).MapNpc(MapNpcNum).Y + 1
            Packet = SNpcMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).X & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case E_Direction.Left_
            MapNpc(MapNum).MapNpc(MapNpcNum).X = MapNpc(MapNum).MapNpc(MapNpcNum).X - 1
            Packet = SNpcMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).X & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    
        Case E_Direction.Right_
            MapNpc(MapNum).MapNpc(MapNpcNum).X = MapNpc(MapNum).MapNpc(MapNpcNum).X + 1
            Packet = SNpcMove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).X & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMap(MapNum, Packet)
    End Select
End Sub

Public Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
Dim Packet As String

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > UBound(MapSpawn(MapNum).Npc) Or Dir < E_Direction.Down_ Or Dir > E_Direction.Right_ Then Exit Sub
    
    MapNpc(MapNum).MapNpc(MapNpcNum).Dir = Dir
    Packet = SNpcDir & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
Dim i As Long
Dim n As Long

    n = 0
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                n = n + 1
            End If
        End If
    Next
    
    GetTotalMapPlayers = n
    
End Function

Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If
    
    Select Case Vital
        Case HP
            GetNpcMaxVital = Npc(NpcNum).HP
        Case MP
            GetNpcMaxVital = Npc(NpcNum).Stat(Stats.Magic) * 2
        Case SP
            GetNpcMaxVital = Npc(NpcNum).Stat(Stats.SPEED) * 2
    End Select
End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If
    
    Select Case Vital
        Case HP
            i = Int(Npc(NpcNum).Stat(Stats.Defense) * 0.33)
            If i < 1 Then i = 1
            GetNpcVitalRegen = i
        'Case MP
        
        'Case SP
    
    End Select
End Function

Public Function IsInRange(ByVal X As Integer, ByVal Y As Integer, ByVal TargetX As Byte, ByVal TargetY As Byte, ByVal Range As Byte) As Boolean

    X = Sqr(((X - TargetX) * (X - TargetX)) + ((Y - TargetY) * (Y - TargetY)))
    
    If X <= Range Then
        IsInRange = True
        Exit Function
    End If
    
End Function

Public Function Meets_ItemRequired(ByVal Index As Long, ByVal InvNum As Long) As Boolean
Dim i As Long

    Meets_ItemRequired = True
    
    For i = 0 To Item_Requires.Count - 1
        If Item(GetPlayerInvItemNum(Index, InvNum)).Required(i) > 0 Then
            Select Case i
            
                Case 0 To 3
                    If GetPlayerStat(Index, i + 1) < Item(GetPlayerInvItemNum(Index, InvNum)).Required(i) Then
                        Meets_ItemRequired = False
                        Exit Function
                    End If
                    
                Case Item_Requires.Class_
                    If GetPlayerClass(Index) <> Item(GetPlayerInvItemNum(Index, InvNum)).Required(i) Then
                        Meets_ItemRequired = False
                        Exit Function
                    End If
                    
                Case Item_Requires.Level_
                    If GetPlayerLevel(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).Required(i) Then
                        Meets_ItemRequired = False
                        Exit Function
                    End If
                    
                Case Item_Requires.Access_
                    If GetPlayerAccess(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).Required(i) Then
                        Meets_ItemRequired = False
                        Exit Function
                    End If
                    
            End Select
        End If
    Next
    
End Function
