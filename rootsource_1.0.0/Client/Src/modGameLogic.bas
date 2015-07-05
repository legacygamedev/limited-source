Attribute VB_Name = "modGameLogic"
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************

Public Sub GameLoop()
    Dim i, n As Long
    
    Dim TickFPS As Long
    Dim FPS As Long
    
    Dim Tick As Long
    
    Dim WalkTimer As Long
    Dim tmr25 As Long
    Dim tmr10000 As Long
    
    InitTileSurf 1
    Pixelation = 16
    ' *** Start GameLoop ***
    Do While InGame
        Tick = GetTickCount
        
        
        
        If tmr25 < Tick Then
                        
            If Pixelation <> 0 Then Pixelation = Pixelation - 1
                        
            InGame = IsConnected
            
            If SentSync = False Then Call SyncPacket
  
            If GetForegroundWindow() = frmMainGame.hwnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If

            ' Change map animation every 250 milliseconds
            If MapAnimTimer < Tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = Tick + 250
            End If

            tmr25 = Tick + 25
        End If
                         
        ' Process movements (actually move them)
        If WalkTimer < Tick Then
            For i = 1 To High_Index
                If Player(i).Moving > 0 Then
                    Call ProcessMovement(i)
                End If
            Next
            
            
            For n = 1 To 9
                If tMap(n) <> 0 Then
                    For i = 1 To MAX_MAP_NPCS
                        If map(n).Npc(i) > 0 Then
                            If MapNpc(i, tMap(n)).Moving > 0 Then
                                Call ProcessNpcMovement(i, tMap(n))
                            End If
                        End If
                    Next
                End If
            Next
            
            WalkTimer = Tick + 30 ' edit this value to change WalkTimer
        End If
        
        ' Check if surface is ready to be unloaded
        If tmr10000 < Tick Then
            ' Sprites
            For i = 1 To NumSprites
                Call DD_CheckSurfTimer(DDS_Sprite(i))
            Next
            
            ' Spells
            For i = 1 To NumSpells
                Call DD_CheckSurfTimer(DDS_Spell(i))
            Next
            
            ' Items
            For i = 1 To NumItems
                Call DD_CheckSurfTimer(DDS_Item(i))
            Next
            
            tmr10000 = Tick + 10000
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        If Editor = EDITOR_NONE Then
            NewPlayerX = Player(MyIndex).x - 8
            NewPlayerY = Player(MyIndex).y - 5
            
            NewXOffset = Player(MyIndex).XOffset
            NewYOffset = Player(MyIndex).YOffset
            
            StaticX = -1 * NewPlayerX * PIC_X - NewXOffset
            StaticY = -1 * NewPlayerY * PIC_Y - NewYOffset
        ElseIf Editor = EDITOR_MAP Then
            If ReCalcTiles Then
                CalcTilePositions
                ReCalcTiles = False
            End If
            
            NewPlayerX = 0
            NewPlayerY = 0
            
            NewXOffset = 0
            NewYOffset = 0
            
            StaticX = 0
            StaticY = 0
        End If
            
        Call Render_Graphics
        
        ' Lock fps
'        Do While GetTickCount < Tick + 15
'            DoEvents
'            Sleep 1
'        Loop
        DoEvents
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
        
        
    Loop
    
    frmMainGame.Visible = False
    
    If isLogging Then
        frmMainGame.txtChat = vbNullString
        frmMainGame.txtMyChat = vbNullString
        isLogging = False
        frmMainMenu.Visible = False
        GettingMap = True
    Else
         ' Shutdown the game
        frmSendGetData.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If
    
End Sub

Private Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long

    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving
        Case MOVING_WALKING
            MovementSpeed = WALK_SPEED
        Case MOVING_RUNNING
            MovementSpeed = RUN_SPEED
    End Select

    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).YOffset = Player(Index).YOffset - MovementSpeed
            If Player(Index).YOffset <= 0 Then Player(Index).YOffset = 0
        Case DIR_DOWN
            Player(Index).YOffset = Player(Index).YOffset + MovementSpeed
            If Player(Index).YOffset >= 0 Then Player(Index).YOffset = 0
        Case DIR_LEFT
            Player(Index).XOffset = Player(Index).XOffset - MovementSpeed
            If Player(Index).XOffset <= 0 Then Player(Index).XOffset = 0
        Case DIR_RIGHT
            Player(Index).XOffset = Player(Index).XOffset + MovementSpeed
            If Player(Index).XOffset >= 0 Then Player(Index).XOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).XOffset = 0 Then
        If Player(Index).YOffset = 0 Then
            Player(Index).Moving = 0
        End If
    End If
    
    If PlayerMapBounds(MyIndex) And (Index = MyIndex) Then
        If Player(MyIndex).XOffset = 0 Then
            If Player(MyIndex).YOffset = 0 Then
                If map(5).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                    Dim tempx, tempy As Long
                    
                    tempx = GetPlayerX(MyIndex)
                    tempy = GetPlayerY(MyIndex)
                        
                    Call SetPlayerX(MyIndex, map(5).Tile(tempx, tempy).Data2)
                    Call SetPlayerY(MyIndex, map(5).Tile(tempx, tempy).Data3)
                    Player(MyIndex).YOffset = 0
                    Player(MyIndex).XOffset = 0
                    Call SetPlayerMap(MyIndex, map(5).Tile(tempx, tempy).Data1)
                    Call CalcTilePositions
                End If
            End If
        End If
    End If

End Sub

Private Sub ProcessNpcMovement(ByVal MapNpcNum As Long, ByVal MapNum As Long)
    ' Check if NPC is walking, and if so process moving them over
    
    'If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
    
        Select Case MapNpc(MapNpcNum, MapNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum, MapNum).YOffset = MapNpc(MapNpcNum, MapNum).YOffset - WALK_SPEED
                If MapNpc(MapNpcNum, MapNum).YOffset <= 0 Then MapNpc(MapNpcNum, MapNum).YOffset = 0
            Case DIR_DOWN
                MapNpc(MapNpcNum, MapNum).YOffset = MapNpc(MapNpcNum, MapNum).YOffset + WALK_SPEED
                If MapNpc(MapNpcNum, MapNum).YOffset >= 0 Then MapNpc(MapNpcNum, MapNum).YOffset = 0
            Case DIR_LEFT
                MapNpc(MapNpcNum, MapNum).XOffset = MapNpc(MapNpcNum, MapNum).XOffset - WALK_SPEED
                If MapNpc(MapNpcNum, MapNum).XOffset <= 0 Then MapNpc(MapNpcNum, MapNum).XOffset = 0
            Case DIR_RIGHT
                MapNpc(MapNpcNum, MapNum).XOffset = MapNpc(MapNpcNum, MapNum).XOffset + WALK_SPEED
                If MapNpc(MapNpcNum, MapNum).XOffset >= 0 Then MapNpc(MapNpcNum, MapNum).XOffset = 0
        End Select
        
        ' Check if completed walking over to the next tile
        If MapNpc(MapNpcNum, MapNum).XOffset = 0 Then
            If MapNpc(MapNpcNum, MapNum).YOffset = 0 Then
                MapNpc(MapNpcNum, MapNum).Moving = 0
            End If
        End If
        
    'End If
End Sub

Private Function IsTryingToMove() As Boolean
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If
End Function

Public Sub CheckAttack()
Dim Buffer As clsBuffer
    
    
    
    If ControlDown Then
        If Player(MyIndex).AttackTimer + 1000 < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then
                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                End With
                
                Set Buffer = New clsBuffer
                Buffer.PreAllocate 2
                Buffer.WriteInteger CAttack
                Call SendData(Buffer.ToArray())
            End If
        End If
    End If
End Sub

Private Function CanMove() As Boolean
Dim d As Long

    CanMove = True
   
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
   
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False
            Exit Function
        End If
    End If
   
    d = GetPlayerDir(MyIndex)
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False
               
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If (map(5).Up > 0) And (map(2).Tile(GetPlayerX(MyIndex), MAX_MAPY).Type <> TILE_TYPE_BLOCKED) Then
                Call SetPlayerMap(MyIndex, map(5).Up)
                Call SetPlayerY(MyIndex, MAX_MAPY + 1)
                Call LoadMaps(GetPlayerMap(MyIndex))
                Call MapEditorLeaveMap
                'Call SendPlayerRequestNewMap
            Else
                CanMove = False
            End If

            Exit Function
        End If
    End If
           
    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False
               
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If (map(5).Down > 0) And (map(8).Tile(GetPlayerX(MyIndex), 0).Type <> TILE_TYPE_BLOCKED) Then
                Call SetPlayerMap(MyIndex, map(5).Down)
                Call SetPlayerY(MyIndex, -1)
                Call LoadMaps(GetPlayerMap(MyIndex))
                Call MapEditorLeaveMap
                'Call SendPlayerRequestNewMap
            Else
                CanMove = False
            End If
        End If
    End If
               
    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False
               
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If (map(5).Left > 0) And (map(4).Tile(MAX_MAPX, GetPlayerY(MyIndex)).Type <> TILE_TYPE_BLOCKED) Then
                Call SetPlayerMap(MyIndex, map(5).Left)
                Call SetPlayerX(MyIndex, MAX_MAPX + 1)
                Call LoadMaps(GetPlayerMap(MyIndex))
                Call MapEditorLeaveMap
                'Call SendPlayerRequestNewMap
            Else
                CanMove = False
            End If
        End If
    End If
       
    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If (map(5).Right > 0) And (map(6).Tile(0, GetPlayerY(MyIndex)).Type <> TILE_TYPE_BLOCKED) Then
                Call SetPlayerMap(MyIndex, map(5).Right)
                Call SetPlayerX(MyIndex, -1)
                Call LoadMaps(GetPlayerMap(MyIndex))
                Call MapEditorLeaveMap
                'Call SendPlayerRequestNewMap
            Else
                CanMove = False
            End If
        End If
    End If
End Function

Private Function CheckDirection(ByVal Direction As Byte) As Boolean
Dim x As Long
Dim y As Long
Dim i As Long
Dim n As Long

    CheckDirection = False
   
    Select Case Direction
        Case DIR_UP
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            x = GetPlayerX(MyIndex) - 1
            y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            x = GetPlayerX(MyIndex) + 1
            y = GetPlayerY(MyIndex)
    End Select
   
    ' Check to see if the map tile is blocked or not
    If PlayerMapBounds(MyIndex) Then
    If map(5).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If
    End If
                               
    ' Check to see if the key door is open or not
    If map(5).Tile(x, y).Type = TILE_TYPE_KEY Then
        ' This actually checks if its open or not
        If TempTile(x, y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
   
    ' Check to see if a player is already on that tile
    For i = 1 To PlayersOnMapHighIndex
        If GetPlayerX(PlayersOnMap(i)) = x Then
            If GetPlayerY(PlayersOnMap(i)) = y Then
                CheckDirection = True
                Exit Function
            End If
        End If
    Next

    ' Check to see if a npc is already on that tile
    For i = 1 To MAX_MAP_NPCS
        For n = 1 To 9
            If tMap(n) = Player(MyIndex).map Then
                If (tMap(n) <> 0) And (tMap(n) < MAX_MAPS + 1) Then
                    If MapNpc(i, tMap(n)).Num > 0 Then
                        If MapNpc(i, tMap(n)).x = x Then
                            If MapNpc(i, tMap(n)).y = y Then
                                CheckDirection = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Next
    
End Function

Private Sub CheckMovement()
    If IsTryingToMove Then
        If CanMove Then
            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MOVING_RUNNING
            Else
                Player(MyIndex).Moving = MOVING_WALKING
            End If
        
            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                    Call CalcTilePositions
                    
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                    Call CalcTilePositions
                    
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                    Call CalcTilePositions
            
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                    Call CalcTilePositions
                    
            End Select
            
        End If
    End If
    

End Sub

Public Function PlayerMapBounds(ByVal Index As Long) As Boolean
Dim x, y As Long

    PlayerMapBounds = True

    x = GetPlayerX(Index)
    y = GetPlayerY(Index)
    
    If x < 0 Then PlayerMapBounds = False
    If x > MAX_MAPX Then PlayerMapBounds = False
    If y < 0 Then PlayerMapBounds = False
    If y > MAX_MAPY Then PlayerMapBounds = False
End Function

Public Sub UpdateInventory()
Dim i As Long

    frmMainGame.lstInv.Clear
    
    ' Show the inventory
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmMainGame.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                ' Check if this item is being worn
                If GetPlayerEquipmentSlot(MyIndex, Weapon) = i Or GetPlayerEquipmentSlot(MyIndex, Armor) = i Or GetPlayerEquipmentSlot(MyIndex, Helmet) = i Or GetPlayerEquipmentSlot(MyIndex, Shield) = i Then
                    frmMainGame.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                Else
                    frmMainGame.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                End If
            End If
        Else
            frmMainGame.lstInv.AddItem "<free inventory slot>"
        End If
    Next
    
    frmMainGame.lstInv.ListIndex = 0
End Sub

Public Sub GetPlayersOnMap()
Dim i As Long

    PlayersOnMapHighIndex = 1

    ReDim PlayersOnMap(1 To MAX_PLAYERS)
        
    For i = 1 To High_Index
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                PlayersOnMap(PlayersOnMapHighIndex) = i
                PlayersOnMapHighIndex = PlayersOnMapHighIndex + 1
            End If
        End If
    Next
    
    PlayersOnMapHighIndex = PlayersOnMapHighIndex - 1
End Sub

Public Sub PlayerSearch(ByVal CurX As Integer, ByVal CurY As Integer)
Dim Buffer As clsBuffer
    
    
    If isInBounds Then
        Set Buffer = New clsBuffer
        Buffer.PreAllocate 10
        Buffer.WriteInteger CSearch
        Buffer.WriteLong CurX
        Buffer.WriteLong CurY
        Call SendData(Buffer.ToArray())
    End If
End Sub

Public Function isInBounds()
    If (CurX >= 0) Then
        If (CurX <= MAX_MAPX) Then
            If (CurY >= 0) Then
                If (CurY <= MAX_MAPY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If
End Function

Public Sub UpdateDrawMapName()
    DrawMapNameX = (MAX_MAPX + 1) * PIC_X \ 2 - ((Len(Trim$(map(5).Name)) \ 2) * 8)
    DrawMapNameY = 1
     
    Select Case map(5).Moral
        Case MAP_MORAL_NONE
            DrawMapNameColor = QBColor(BrightRed)
            
        Case MAP_MORAL_SAFE
            DrawMapNameColor = QBColor(White)
            
        Case MAP_MORAL_INN
            DrawMapNameColor = RGB(220, 192, 0)
            
        Case MAP_MORAL_ARENA
            DrawMapNameColor = RGB(174, 174, 174)
            
        Case Else
            DrawMapNameColor = QBColor(White)
    End Select
    
End Sub

Public Sub UseItem()
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
End Sub

Public Sub CastSpell()
Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If SpellSelected < 1 Or SpellSelected > MAX_SPELLS Then
        Exit Sub
    End If
    
    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(SpellSelected).MPReq Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(SpellSelected).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(SpellSelected) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.PreAllocate 6
                Buffer.WriteInteger CCast
                Buffer.WriteLong SpellSelected
                Call SendData(Buffer.ToArray())
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If
End Sub

Public Sub ClearTempTile()
Dim x As Long
Dim y As Long

    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            TempTile(x, y).DoorOpen = NO
        Next
    Next
End Sub

Public Sub CalcTilePositions()
Dim x As Long
Dim y As Long
Dim tx As Long
Dim ty As Long
Dim x1 As Long
Dim y1 As Long
Dim n As Long
Dim m As Long

If Editor = EDITOR_NONE Then
    For x1 = Player(MyIndex).x - 9 To (MAX_MAPX + 1) + Player(MyIndex).x - 8
        For y1 = Player(MyIndex).y - 6 To (MAX_MAPY + 1) + Player(MyIndex).y - 5
            m = Lookup(x1, y1).map
            x = Lookup(x1, y1).x
            y = Lookup(x1, y1).y
            tx = (x1 - Player(MyIndex).x + 8)
            ty = (y1 - Player(MyIndex).y + 5)
            
            MapTilePosition(tx, ty).PosX = tx * PIC_X
            MapTilePosition(tx, ty).PosY = ty * PIC_Y
            
            For n = 0 To 8
                MapTilePosition(tx, ty).Layer(n).Top = (map(m).Tile(x, y).Num(n) \ TILESHEET_WIDTH) * PIC_Y
                MapTilePosition(tx, ty).Layer(n).Bottom = MapTilePosition(tx, ty).Layer(n).Top + PIC_Y
                MapTilePosition(tx, ty).Layer(n).Left = (map(m).Tile(x, y).Num(n) Mod TILESHEET_WIDTH) * PIC_X
                MapTilePosition(tx, ty).Layer(n).Right = MapTilePosition(tx, ty).Layer(n).Left + PIC_X
            Next n
        Next y1
    Next x1
ElseIf Editor = EDITOR_MAP Then
    For x1 = 0 To MAX_MAPX
        For y1 = 0 To MAX_MAPY
            m = Lookup(x1, y1).map
            x = Lookup(x1, y1).x
            y = Lookup(x1, y1).y
            tx = (x1 - Player(MyIndex).x + 8)
            ty = (y1 - Player(MyIndex).y + 5)
            
            MapTilePosition(x, y).PosX = x * PIC_X
            MapTilePosition(x, y).PosY = y * PIC_Y
            
            For n = 0 To 8
                MapTilePosition(x, y).Layer(n).Top = (map(m).Tile(x, y).Num(n) \ TILESHEET_WIDTH) * PIC_Y
                MapTilePosition(x, y).Layer(n).Bottom = MapTilePosition(x, y).Layer(n).Top + PIC_Y
                MapTilePosition(x, y).Layer(n).Left = (map(m).Tile(x, y).Num(n) Mod TILESHEET_WIDTH) * PIC_X
                MapTilePosition(x, y).Layer(n).Right = MapTilePosition(x, y).Layer(n).Left + PIC_X
            Next n
        Next
    Next
End If
End Sub

Public Sub InitMapTables()
Dim MapNum As Long
Dim x As Long
Dim y As Long
Dim x1 As Long
Dim y1 As Long

    ReDim Lookup(-MAX_MAPX - 1 To MAX_MAPX * 2 + 1, -MAX_MAPY - 1 To MAX_MAPY * 2 + 1) As TileLookupStruct
    
    For x = -MAX_MAPX - 1 To MAX_MAPX * 2 + 1
        For y = -MAX_MAPY - 1 To MAX_MAPY * 2 + 1
            x1 = x
            y1 = y
            If x1 >= 0 And x1 <= MAX_MAPX Then
                If y1 >= 0 And y1 <= MAX_MAPY Then
                      MapNum = 5
                ElseIf y1 > MAX_MAPY Then
                      y1 = y1 - (MAX_MAPY + 1)
                      MapNum = 8
                ElseIf y1 < 0 Then
                      y1 = (MAX_MAPY + 1) + y1
                      MapNum = 2
                End If
            ElseIf x1 > MAX_MAPX Then
                x1 = x1 - (MAX_MAPX + 1)
                If y1 >= 0 And y1 <= MAX_MAPY Then
                      MapNum = 6
                ElseIf y1 > MAX_MAPY Then
                      y1 = y1 - (MAX_MAPY + 1)
                      MapNum = 9
                ElseIf y1 < 0 Then
                      y1 = (MAX_MAPY + 1) + y1
                      MapNum = 3
                End If
            ElseIf x1 < 0 Then
                x1 = x1 + (MAX_MAPX + 1)
                If y1 >= 0 And y1 <= MAX_MAPY Then
                      MapNum = 4
                ElseIf y1 > MAX_MAPY Then
                      y1 = y1 - (MAX_MAPY + 1)
                      MapNum = 7
                ElseIf y1 < 0 Then
                      y1 = (MAX_MAPY + 1) + y1
                      MapNum = 1
                End If
            End If
            
            Lookup(x, y).map = MapNum
            Lookup(x, y).x = x1
            Lookup(x, y).y = y1
        Next
    Next
End Sub

Public Sub InitMapData()
Dim i As Long
Dim MusicFile As String

    MusicFile = Trim$(CStr(map(5).Music)) & ".mid"
    
'    ' get high NPC index
    High_Npc_Index = 0
    For i = 1 To MAX_MAP_NPCS
        If map(5).Npc(i) > 0 Then
            High_Npc_Index = High_Npc_Index + 1
        Else
            Exit For
        End If
    Next
    
    ' Play music
    If map(5).Music > 0 Then
        If MusicFile <> CurrentMusic Then
            DirectMusic_StopMidi
            Call DirectMusic_PlayMidi(MusicFile)
            CurrentMusic = MusicFile
        End If
    Else
        DirectMusic_StopMidi
        CurrentMusic = 0
    End If
    
    Call UpdateDrawMapName
    
    If map(5).Shop = 0 Then
        frmMainGame.picTradeButton.Visible = False
    Else
        frmMainGame.picTradeButton.Visible = True
    End If
    
    Call CalcTilePositions
    
    Call InitTileSurf(map(5).TileSet)

End Sub

Public Sub DevMsg(ByVal Text As String, ByVal Color As Byte)
    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(Text, Color)
        End If
    End If
    Debug.Print Text
End Sub
