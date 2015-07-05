Attribute VB_Name = "modGameLogic"
Option Explicit

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

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapnum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

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

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If itemnum < 1 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapnum)
    Call SpawnItemSlot(i, itemnum, ItemVal, mapnum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemnum < 0 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemnum >= 0 And itemnum <= MAX_ITEMS Then
            MapItem(mapnum, i).playerName = playerName
            MapItem(mapnum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(mapnum, i).canDespawn = canDespawn
            MapItem(mapnum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(mapnum, i).Num = itemnum
            MapItem(mapnum, i).Value = ItemVal
            MapItem(mapnum, i).x = x
            MapItem(mapnum, i).y = y
            ' send to map
            SendSpawnItemToMap mapnum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapnum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapnum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(mapnum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, 1, mapnum, x, y)
                Else
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, Map(mapnum).Tile(x, y).Data2, mapnum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal mapnum As Long, Optional ForcedSpawn As Boolean = False)
    Dim Buffer As clsBuffer
    Dim npcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Sub
    npcNum = Map(mapnum).Npc(mapNpcNum)
    If ForcedSpawn = False And Map(mapnum).NpcSpawnType(mapNpcNum) = 1 Then npcNum = 0
    If npcNum > 0 Then
    
        MapNpc(mapnum).Npc(mapNpcNum).Num = npcNum
        MapNpc(mapnum).Npc(mapNpcNum).target = 0
        MapNpc(mapnum).Npc(mapNpcNum).targetType = 0 ' clear
        
        MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
        MapNpc(mapnum).Npc(mapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(npcNum, Vitals.MP)
        
        MapNpc(mapnum).Npc(mapNpcNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(mapnum).MaxX
            For y = 0 To Map(mapnum).MaxY
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(mapnum).Tile(x, y).Data1 = mapNpcNum Then
                        MapNpc(mapnum).Npc(mapNpcNum).x = x
                        MapNpc(mapnum).Npc(mapNpcNum).y = y
                        MapNpc(mapnum).Npc(mapNpcNum).Dir = Map(mapnum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, Map(mapnum).MaxX)
                y = Random(0, Map(mapnum).MaxY)
    
                If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
                If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapnum, x, y) Then
                    MapNpc(mapnum).Npc(mapNpcNum).x = x
                    MapNpc(mapnum).Npc(mapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(mapnum).MaxX
                For y = 0 To Map(mapnum).MaxY

                    If NpcTileIsOpen(mapnum, x, y) Then
                        MapNpc(mapnum).Npc(mapNpcNum).x = x
                        MapNpc(mapnum).Npc(mapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).Num
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).Dir
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
            UpdateMapBlock mapnum, MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y, True
        End If
        
        SendMapNpcVitals mapnum, mapNpcNum
    Else
        MapNpc(mapnum).Npc(mapNpcNum).Num = 0
        MapNpc(mapnum).Npc(mapNpcNum).target = 0
        MapNpc(mapnum).Npc(mapNpcNum).targetType = 0 ' clear
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong mapNpcNum
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
    End If

End Sub

Public Sub SpawnMapEventsFor(index As Long, mapnum As Long)
Dim i As Long, x As Long, y As Long, z As Long, spawncurrentevent As Boolean, p As Long
Dim Buffer As clsBuffer
    
    TempPlayer(index).EventMap.CurrentEvents = 0
    ReDim TempPlayer(index).EventMap.EventPages(0)
    
    If Map(mapnum).EventCount <= 0 Then Exit Sub
    For i = 1 To Map(mapnum).EventCount
        If Map(mapnum).Events(i).PageCount > 0 Then
            For z = Map(mapnum).Events(i).PageCount To 1 Step -1
                With Map(mapnum).Events(i).Pages(z)
                    spawncurrentevent = True
                    
                    If .chkVariable = 1 Then
                        If Player(index).Variables(.VariableIndex) < .VariableCondition Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSwitch = 1 Then
                        If Player(index).Switches(.SwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkHasItem = 1 Then
                        If HasItem(index, .HasItemIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSelfSwitch = 1 Then
                        If Map(mapnum).Events(i).SelfSwitches(.SelfSwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If spawncurrentevent = True Or (spawncurrentevent = False And z = 1) Then
                        'spawn the event... send data to player
                        TempPlayer(index).EventMap.CurrentEvents = TempPlayer(index).EventMap.CurrentEvents + 1
                        ReDim Preserve TempPlayer(index).EventMap.EventPages(TempPlayer(index).EventMap.CurrentEvents)
                        With TempPlayer(index).EventMap.EventPages(TempPlayer(index).EventMap.CurrentEvents)
                            If Map(mapnum).Events(i).Pages(z).GraphicType = 1 Then
                                Select Case Map(mapnum).Events(i).Pages(z).GraphicY
                                    Case 0
                                        .Dir = DIR_DOWN
                                    Case 1
                                        .Dir = DIR_LEFT
                                    Case 2
                                        .Dir = DIR_RIGHT
                                    Case 3
                                        .Dir = DIR_UP
                                End Select
                            Else
                                .Dir = 0
                            End If
                            .GraphicNum = Map(mapnum).Events(i).Pages(z).Graphic
                            .GraphicType = Map(mapnum).Events(i).Pages(z).GraphicType
                            .GraphicX = Map(mapnum).Events(i).Pages(z).GraphicX
                            .GraphicY = Map(mapnum).Events(i).Pages(z).GraphicY
                            .GraphicX2 = Map(mapnum).Events(i).Pages(z).GraphicX2
                            .GraphicY2 = Map(mapnum).Events(i).Pages(z).GraphicY2
                            Select Case Map(mapnum).Events(i).Pages(z).MoveSpeed
                                Case 0
                                    .movementspeed = 2
                                Case 1
                                    .movementspeed = 3
                                Case 2
                                    .movementspeed = 4
                                Case 3
                                    .movementspeed = 6
                                Case 4
                                    .movementspeed = 12
                                Case 5
                                    .movementspeed = 24
                            End Select
                            If Map(mapnum).Events(i).Global Then
                                .x = TempEventMap(mapnum).Events(i).x
                                .y = TempEventMap(mapnum).Events(i).y
                                .Dir = TempEventMap(mapnum).Events(i).Dir
                                .MoveRouteStep = TempEventMap(mapnum).Events(i).MoveRouteStep
                            Else
                                .x = Map(mapnum).Events(i).x
                                .y = Map(mapnum).Events(i).y
                                .MoveRouteStep = 0
                            End If
                            .Position = Map(mapnum).Events(i).Pages(z).Position
                            .eventID = i
                            .pageID = z
                            If spawncurrentevent = True Then
                                .Visible = 1
                            Else
                                .Visible = 0
                            End If
                            
                            .MoveType = Map(mapnum).Events(i).Pages(z).MoveType
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(mapnum).Events(i).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(mapnum).Events(i).Pages(z).MoveRouteCount)
                                If Map(mapnum).Events(i).Pages(z).MoveRouteCount > 0 Then
                                    For p = 0 To Map(mapnum).Events(i).Pages(z).MoveRouteCount
                                        .MoveRoute(p) = Map(mapnum).Events(i).Pages(z).MoveRoute(p)
                                    Next
                                End If
                            End If
                            
                            .RepeatMoveRoute = Map(mapnum).Events(i).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(mapnum).Events(i).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(mapnum).Events(i).Pages(z).MoveFreq
                            .MoveSpeed = Map(mapnum).Events(i).Pages(z).MoveSpeed
                            
                            .WalkingAnim = Map(mapnum).Events(i).Pages(z).WalkAnim
                            .WalkThrough = Map(mapnum).Events(i).Pages(z).WalkThrough
                            .ShowName = Map(mapnum).Events(i).Pages(z).ShowName
                            .FixedDir = Map(mapnum).Events(i).Pages(z).DirFix
                            
                        End With
                        GoTo nextevent
                    End If
                End With
            Next
        End If
nextevent:
    Next
    
    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
        For i = 1 To TempPlayer(index).EventMap.CurrentEvents
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnEvent
            Buffer.WriteLong i
            With TempPlayer(index).EventMap.EventPages(i)
                Buffer.WriteString Map(GetPlayerMap(index)).Events(i).Name
                Buffer.WriteLong .Dir
                Buffer.WriteLong .GraphicNum
                Buffer.WriteLong .GraphicType
                Buffer.WriteLong .GraphicX
                Buffer.WriteLong .GraphicX2
                Buffer.WriteLong .GraphicY
                Buffer.WriteLong .GraphicY2
                Buffer.WriteLong .movementspeed
                Buffer.WriteLong .x
                Buffer.WriteLong .y
                Buffer.WriteLong .Position
                Buffer.WriteLong .Visible
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkAnim
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).DirFix
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkThrough
                Buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).ShowName
            End With
            SendDataTo index, Buffer.ToArray
            Set Buffer = Nothing
        Next
    End If
End Sub

Public Function NpcTileIsOpen(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapnum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapnum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(mapnum).Npc(LoopI).Num > 0 Then
            If MapNpc(mapnum).Npc(LoopI).x = x Then
                If MapNpc(mapnum).Npc(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next
    
    For LoopI = 1 To TempEventMap(mapnum).EventCount
        If TempEventMap(mapnum).Events(LoopI).active = 1 Then
            If MapNpc(mapnum).Npc(LoopI).x = TempEventMap(mapnum).Events(LoopI).x Then
                If MapNpc(mapnum).Npc(LoopI).y = TempEventMap(mapnum).Events(LoopI).y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If
    Next

    If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapnum)
    Next
    
    CacheMapBlocks mapnum

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Sub SpawnAllMapGlobalEvents()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnGlobalEvents(i)
    Next

End Sub

Sub SpawnGlobalEvents(ByVal mapnum As Long)
    Dim i As Long, z As Long
    
    If Map(mapnum).EventCount > 0 Then
        TempEventMap(mapnum).EventCount = 0
        ReDim TempEventMap(mapnum).Events(0)
        For i = 1 To Map(mapnum).EventCount
            TempEventMap(mapnum).EventCount = TempEventMap(mapnum).EventCount + 1
            ReDim Preserve TempEventMap(mapnum).Events(0 To TempEventMap(mapnum).EventCount)
            If Map(mapnum).Events(i).PageCount > 0 Then
                If Map(mapnum).Events(i).Global = 1 Then
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).x = Map(mapnum).Events(i).x
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).y = Map(mapnum).Events(i).y
                    If Map(mapnum).Events(i).Pages(1).GraphicType = 1 Then
                        Select Case Map(mapnum).Events(i).Pages(1).GraphicY
                            Case 0
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                            Case 1
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_LEFT
                            Case 2
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_RIGHT
                            Case 3
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_UP
                        End Select
                    Else
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                    End If
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).active = 1
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = Map(mapnum).Events(i).Pages(1).MoveType
                    
                    If TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = 2 Then
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRouteCount = Map(mapnum).Events(i).Pages(1).MoveRouteCount
                        ReDim TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(0 To Map(mapnum).Events(i).Pages(1).MoveRouteCount)
                        For z = 0 To Map(mapnum).Events(i).Pages(1).MoveRouteCount
                            TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(z) = Map(mapnum).Events(i).Pages(1).MoveRoute(z)
                        Next
                    End If
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).RepeatMoveRoute = Map(mapnum).Events(i).Pages(1).RepeatMoveRoute
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).IgnoreIfCannotMove = Map(mapnum).Events(i).Pages(1).IgnoreMoveRoute
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveFreq = Map(mapnum).Events(i).Pages(1).MoveFreq
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveSpeed = Map(mapnum).Events(i).Pages(1).MoveSpeed
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkThrough = Map(mapnum).Events(i).Pages(1).WalkThrough
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).FixedDir = Map(mapnum).Events(i).Pages(1).DirFix
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkingAnim = Map(mapnum).Events(i).Pages(1).WalkAnim
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).ShowName = Map(mapnum).Events(i).Pages(1).ShowName
                    
                End If
            End If
        Next
    End If

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapnum).Npc(mapNpcNum).x
    y = MapNpc(mapnum).Npc(mapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).Npc(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(mapNpcNum).x) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                n = Map(mapnum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).Npc(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(mapNpcNum).x) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapnum).Npc(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(mapNpcNum).x - 1) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                n = Map(mapnum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapnum).Npc(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(mapNpcNum).x + 1) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapnum).Npc(mapNpcNum).Dir = Dir
    UpdateMapBlock mapnum, MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y, False

    Select Case Dir
        Case DIR_UP
            MapNpc(mapnum).Npc(mapNpcNum).y = MapNpc(mapnum).Npc(mapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapnum).Npc(mapNpcNum).y = MapNpc(mapnum).Npc(mapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapnum).Npc(mapNpcNum).x = MapNpc(mapnum).Npc(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapnum).Npc(mapNpcNum).x = MapNpc(mapnum).Npc(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select
    
    UpdateMapBlock mapnum, MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y, True

End Sub

Sub NpcDir(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapnum).Npc(mapNpcNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong Dir
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal mapnum As Long)
    Dim y As Long
    Dim x As Long
    temptile(mapnum).DoorTimer = 0
    ReDim temptile(mapnum).DoorOpen(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            temptile(mapnum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapnum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapnum).ResourceData(0 To Resource_Count)
                ResourceCache(mapnum).ResourceData(Resource_Count).x = x
                ResourceCache(mapnum).ResourceData(Resource_Count).y = y
                ResourceCache(mapnum).ResourceData(Resource_Count).cur_health = Resource(Map(mapnum).Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapnum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0
        PlayerMsg index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    rand = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal index As Long)
Dim partyNum As Long, i As Long

    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
        
            ' check if leader
            If Party(partyNum).Leader = index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(i) & " é o lider do grupo.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(index) & " saiu do grupo.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        TempPlayer(index).inParty = 0
                        TempPlayer(index).partyInvite = 0
                        Exit For
                        End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(index) & " saiu do grupo.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        TempPlayer(index).inParty = 0
                        TempPlayer(index).partyInvite = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Grupo desfeito.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(partyNum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).partyInvite = 0
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        partyNum = TempPlayer(index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = index
                    ' let them know
                    PlayerMsg index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = index
        ' let them know
        PlayerMsg index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                'add to the party
                Party(partyNum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                SendPartyVitals partyNum, targetPlayer
                ' let everyone know they've joined
                PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = index
        Party(partyNum).Member(1) = index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, index
        SendPartyVitals partyNum, targetPlayer
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal targetPlayer As Long)
    PlayerMsg index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, x As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal exp As Long, ByVal index As Long, ByVal mapnum As Long)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long, LoseMemberCount As Byte

    ' check if it's worth sharing
    If Not exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP index, exp
        Exit Sub
    End If
    
    ' check members in outhers maps
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        If tmpIndex > 0 Then
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) <> mapnum Then
                    LoseMemberCount = LoseMemberCount + 1
                End If
            End If
        End If
    Next i
    
    ' find out the equal share
    expShare = exp \ (Party(partyNum).MemberCount - LoseMemberCount)
    leftOver = exp Mod (Party(partyNum).MemberCount - LoseMemberCount)
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) = mapnum Then
                    ' give them their share
                    GivePlayerEXP tmpIndex, expShare
                End If
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(rand(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal index As Long, ByVal exp As Long)
    ' give the exp
    Call SetPlayerExp(index, GetPlayerExp(index) + exp)
    SendEXP index
    SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp index
End Sub

Function CanEventMove(index As Long, ByVal mapnum As Long, x As Long, y As Long, eventID As Long, WalkThrough As Long, ByVal Dir As Byte, Optional globalevent As Boolean = False) As Boolean
    Dim i As Long
    Dim n As Long, z As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    CanEventMove = True
    
    

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).Tile(x, y - 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If
                
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y - 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).Npc(i).x = x) And (MapNpc(mapnum).Npc(i).y = y - 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_UP + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                n = Map(mapnum).Tile(x, y + 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = y + 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).Npc(i).x = x) And (MapNpc(mapnum).Npc(i).y = y + 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_DOWN + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).Tile(x - 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x - 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).Npc(i).x = x - 1) And (MapNpc(mapnum).Npc(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x - 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x - 1) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_LEFT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                n = Map(mapnum).Tile(x + 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = x + 1) And (GetPlayerY(i) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).Npc(i).x = x + 1) And (MapNpc(mapnum).Npc(i).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x + 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x + 1) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_RIGHT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

    End Select

End Function

Sub EventDir(playerindex As Long, ByVal mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, Optional globalevent As Boolean = False)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(playerindex).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(playerindex).EventMap.EventPages(eventID).Dir = Dir
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEventDir
    Buffer.WriteLong eventID
    If globalevent Then
        Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
    Else
        Buffer.WriteLong TempPlayer(playerindex).EventMap.EventPages(eventID).Dir
    End If
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub EventMove(index As Long, mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, movementspeed As Long, Optional globalevent As Boolean = False)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
        UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, False
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(index).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(index).EventMap.EventPages(eventID).Dir = Dir
    End If

    Select Case Dir
        Case DIR_UP
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y - 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).y = TempPlayer(index).EventMap.EventPages(eventID).y - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
            
        Case DIR_DOWN
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y + 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).y = TempPlayer(index).EventMap.EventPages(eventID).y + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
        Case DIR_LEFT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x - 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).x = TempPlayer(index).EventMap.EventPages(eventID).x - 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
        Case DIR_RIGHT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x + 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).x = TempPlayer(index).EventMap.EventPages(eventID).x + 1
                Set Buffer = New clsBuffer
                Buffer.WriteLong SEventMove
                Buffer.WriteLong eventID
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                Buffer.WriteLong Dir
                Buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                Buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, Buffer.ToArray()
                Else
                    SendDataTo index, Buffer.ToArray
                End If
                Set Buffer = Nothing
            End If
    End Select

End Sub
