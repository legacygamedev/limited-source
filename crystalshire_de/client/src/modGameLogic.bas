Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
Dim FrameTime As Long, Tick As Long, TickFPS As Long, FPS As Long, i As Long, WalkTimer As Long, x As Long, y As Long
Dim tmr25 As Long, tmr10000 As Long, mapTimer As Long, chatTmr As Long, targetTmr As Long, fogTmr As Long, barTmr As Long
Dim barDifference As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' *** Start GameLoop ***
    Do While InGame
        Tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < Tick Then
            ' check ping
            Call GetPing
            tmr10000 = Tick + 10000
        End If

        If tmr25 < Tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hWnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
            If Count_Spellicon > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i).Spell > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i).Spell).CDTime * 1000) < Tick Then
                                SpellCD(i) = 0
                            End If
                        End If
                    End If
                Next
            End If
            
            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000) < Tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If
            
            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
            
            tmr25 = Tick + 25
        End If
        
        If chatTmr < Tick Then
            If ChatButtonUp Then
                ScrollChatBox 0
            End If
            If ChatButtonDown Then
                ScrollChatBox 1
            End If
            chatTmr = Tick + 50
        End If
        
        ' targetting
        If targetTmr < Tick Then
            If tabDown Then
                FindNearestTarget
            End If
            targetTmr = Tick + 50
        End If
        
        ' fog scrolling
        If fogTmr < Tick Then
            ' move
            fogOffsetX = fogOffsetX - 1
            fogOffsetY = fogOffsetY - 1
            ' reset
            If fogOffsetX < -256 Then fogOffsetX = 0
            If fogOffsetY < -256 Then fogOffsetY = 0
            fogTmr = Tick + 20
        End If
        
        ' elastic bars
        If barTmr < Tick Then
            SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
            SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
            SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    SetBarWidth BarWidth_NpcHP_Max(i), BarWidth_NpcHP(i)
                End If
            Next
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    SetBarWidth BarWidth_PlayerHP_Max(i), BarWidth_PlayerHP(i)
                End If
            Next
            
            ' reset timer
            barTmr = Tick + 10
        End If
        
        ' Animations!
        If mapTimer < Tick Then
            ' animate waterfalls
            Select Case waterfallFrame
                Case 0
                    waterfallFrame = 1
                Case 1
                    waterfallFrame = 2
                Case 2
                    waterfallFrame = 0
            End Select
            
            ' animate autotiles
            Select Case autoTileFrame
                Case 0
                    autoTileFrame = 1
                Case 1
                    autoTileFrame = 2
                Case 2
                    autoTileFrame = 0
            End Select
            
            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If
            
            ' re-set timer
            mapTimer = Tick + 500
        End If

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then

            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex
                If Map.Npc(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i

            WalkTimer = Tick + 30 ' edit this value to change WalkTimer
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        DoEvents
        
        ' fader logic
        If canFade Then
            If faderAlpha <= 0 Then
                canFade = False
                faderAlpha = 0
            Else
                faderAlpha = faderAlpha - faderSpeed
            End If
        End If

        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + 20
                DoEvents
                Sleep 1
            Loop
        End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
    Loop

    frmMain.visible = False

    If isLogging Then
        isLogging = False
        MenuLoop
        GettingMap = True
        Stop_Music
        Play_Music Options.MenuMusic
    Else
        ' Shutdown the game
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MenuLoop()
Dim FrameTime As Long
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim faderTimer As Long
Dim tmr500 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' *** Start GameLoop ***
    Do While inMenu
        Tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Menu
        DoEvents
        
        ' fader logic
        ' 0, 1, 2, 3 = Fading in/out of intro
        ' 4 = fading in to main menu
        ' 5 = fading out of main menu
        ' 6 = fading in to game
        If canFade Then
            If faderTimer = 0 Then
                Select Case faderState
                    Case 0, 2, 4, 6 ' fading in
                        If faderAlpha <= 0 Then
                            faderTimer = Tick + 1000
                        Else
                            ' fade out a bit
                            faderAlpha = faderAlpha - faderSpeed
                        End If
                    Case 1, 3, 5 ' fading out
                        If faderAlpha >= 254 Then
                            If faderState < 5 Then
                                faderState = faderState + 1
                            ElseIf faderState = 5 Then
                                ' fading out to game - make game load during fade
                                faderAlpha = 254
                                ShowGame
                                HideMenu
                                Call GameInit
                                Call GameLoop
                                Exit Sub
                            End If
                        Else
                            ' fade in a bit
                            faderAlpha = faderAlpha + faderSpeed
                        End If
                End Select
            Else
                If faderTimer < Tick Then
                    ' change the speed
                    If faderState > 2 Then faderSpeed = 15
                    ' normal fades
                    If faderState < 4 Then
                        faderState = faderState + 1
                        faderTimer = 0
                    Else
                        faderTimer = 0
                    End If
                End If
            End If
        End If
        
        If tmr500 < Tick Then
            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If
            tmr500 = Tick + 500
        End If

        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + 20
                DoEvents
                Sleep 1
            Loop
        End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving
        Case MOVING_WALKING: MovementSpeed = RUN_SPEED '((ElapsedTime / 1000) * (RUN_SPEED * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = WALK_SPEED ' ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0
        Case DIR_DOWN
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0
        Case DIR_LEFT
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0
        Case DIR_RIGHT
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        Else
            If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
Dim MovementSpeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        MovementSpeed = RUN_SPEED
    Else
        Exit Sub
    End If

    Select Case MapNpc(MapNpcNum).dir
        Case DIR_UP
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - MovementSpeed
            If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0
            
        Case DIR_DOWN
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + MovementSpeed
            If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0
            
        Case DIR_LEFT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - MovementSpeed
            If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0
            
        Case DIR_RIGHT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + MovementSpeed
            If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If MapNpc(MapNpcNum).Moving > 0 Then
        If MapNpc(MapNpcNum).dir = DIR_RIGHT Or MapNpc(MapNpcNum).dir = DIR_DOWN Then
            If (MapNpc(MapNpcNum).xOffset >= 0) And (MapNpc(MapNpcNum).yOffset >= 0) Then
                MapNpc(MapNpcNum).Moving = 0
                If MapNpc(MapNpcNum).Step = 0 Then
                    MapNpc(MapNpcNum).Step = 2
                Else
                    MapNpc(MapNpcNum).Step = 0
                End If
            End If
        Else
            If (MapNpc(MapNpcNum).xOffset <= 0) And (MapNpc(MapNpcNum).yOffset <= 0) Then
                MapNpc(MapNpcNum).Moving = 0
                If MapNpc(MapNpcNum).Step = 0 Then
                    MapNpc(MapNpcNum).Step = 2
                Else
                    MapNpc(MapNpcNum).Step = 0
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckMapGetItem()
Dim Buffer As New clsBuffer, tmpIndex As Long, i As Long, x As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer

    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then
        ' find out if we want to pick it up
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).x = Player(MyIndex).x And MapItem(i).y = Player(MyIndex).y Then
                If MapItem(i).num > 0 Then
                    If Item(MapItem(i).num).BindType = 1 Then
                        ' make sure it's not a party drop
                        If Party.Leader > 0 Then
                            For x = 1 To MAX_PARTY_MEMBERS
                                tmpIndex = Party.Member(x)
                                If tmpIndex > 0 Then
                                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(i).playerName) Then
                                        If Item(MapItem(i).num).ClassReq > 0 Then
                                            If Item(MapItem(i).num).ClassReq <> Player(MyIndex).Class Then
                                                Dialogue "Loot Check", "This item is BoP and is not for your class. Are you sure you want to pick it up?", DIALOGUE_LOOT_ITEM, True
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Else
                        'not bound
                        Exit For
                    End If
                End If
            End If
        Next
        ' nevermind, pick it up
        Player(MyIndex).MapGetTimer = GetTickCount
        Buffer.WriteLong CMapGetItem
        SendData Buffer.ToArray()
    End If

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()
Dim Buffer As clsBuffer
Dim attackspeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ControlDown Then
    
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                End With

                Set Buffer = New clsBuffer
                Buffer.WriteLong CAttack
                SendData Buffer.ToArray()
                Set Buffer = Nothing
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsTryingToMove() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'If DirUp Or DirDown Or DirLeft Or DirRight Then
    If wDown Or sDown Or aDown Or dDown Or upDown Or leftDown Or downDown Or rightDown Then
        IsTryingToMove = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanMove() As Boolean
Dim d As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' not in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        frmMain.picBank.visible = False
    End If
    
    If inTutorial Then
        CanMove = False
        Exit Function
    End If

    d = GetPlayerDir(MyIndex)

    If wDown Or upDown Then
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
            If Map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If downDown Or sDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY Then
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
            If Map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If aDown Or leftDown Then
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
            If Map.left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If dDown Or rightDown Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MaxX Then
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
            If Map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CheckDirection(ByVal direction As Byte) As Boolean
Dim x As Long
Dim y As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    CheckDirection = False
    
    ' check directional blocking
    If isDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case direction
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
    If Map.Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the key door is open or not
    If Map.Tile(x, y).Type = TILE_TYPE_KEY Then

        ' This actually checks if its open or not
        If TempTile(x, y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
    
    ' Check to see if a player is already on that tile
    If Map.Moral = 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If

    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If MapNpc(i).x = x Then
                If MapNpc(i).y = y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "checkDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub CheckMovement()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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
                    Player(MyIndex).yOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select
            
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                GettingMap = True
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal x As Long, ByVal y As Long) As Boolean
    IsValidMapPoint = False

    If x < 0 Then Exit Function
    If y < 0 Then Exit Function
    If x > Map.MaxX Then Exit Function
    If y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
End Function

Public Sub UseItem()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ForgetSpell(ByVal spellSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellSlot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot).Spell > 0 Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CForgetSpell
        Buffer.WriteLong spellSlot
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellSlot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot).Spell = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellSlot).Spell).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellSlot).Spell).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellSlot).Spell > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong CCast
                Buffer.WriteLong spellSlot
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                SpellBuffer = spellSlot
                SpellBufferTimer = GetTickCount
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearTempTile()
Dim x As Long
Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ReDim TempTile(0 To Map.MaxX, 0 To Map.MaxY)

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            TempTile(x, y).DoorOpen = NO
            If Not GettingMap Then cacheRenderState x, y, MapLayer.Mask
        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearTempTile", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DevMsg(ByVal Text As String, ByVal Color As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(Text, Color)
        End If
    End If

    Debug.Print Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "TwipsToPixels", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub CacheResources()
Dim x As Long, y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource_Count = 0

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            If Map.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).x = x
                MapResource(Resource_Count).y = y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateActionMsg(ByVal message As String, ByVal Color As Integer, ByVal MsgType As Byte, ByVal x As Long, ByVal y As Long)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .message = message
        .Color = Color
        .Type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .x = x
        .y = y
        .alpha = 255
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).y = ActionMsg(ActionMsgIndex).y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).x = ActionMsg(ActionMsgIndex).x + Rand(-8, 8)
    End If
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsg(Index).message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).Color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).x = 0
    ActionMsg(Index).y = 0
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
Dim looptime As Long
Dim Layer As Long
Dim FrameCount As Long
Dim lockindex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).FrameIndex(Layer) = 0 Then AnimInstance(Index).FrameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).timer(Layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(Index).FrameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).FrameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).FrameIndex(Layer) = AnimInstance(Index).FrameIndex(Layer) + 1
                End If
                AnimInstance(Index).timer(Layer) = GetTickCount
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InShop = shopnum

    GUIWindow(GUI_SHOP).visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(bankslot).num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).num = itemnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GetBankItemValue = Bank.Item(bankslot).value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef dir As Byte, ByVal block As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If block Then
        blockvar = blockvar Or (2 ^ dir)
    Else
        blockvar = blockvar And Not (2 ^ dir)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsHotbarSlot(ByVal x As Single, ByVal y As Single) As Long
Dim top As Long, left As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    IsHotbarSlot = 0

    For i = 1 To MAX_HOTBAR
        top = GUIWindow(GUI_HOTBAR).y + HotbarTop
        left = GUIWindow(GUI_HOTBAR).x + HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
        If x >= left And x <= left + PIC_X Then
            If y >= top And y <= top + PIC_Y Then
                IsHotbarSlot = i
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub PlayMapSound(ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If entityNum <= 0 Then Exit Sub
    
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).sound)
        ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(Npc(entityNum).sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).sound)
        ' other
        Case Else
            Exit Sub
    End Select
    
    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    Play_Sound soundName
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub
    
    ' set global dialogue index
    dialogueIndex = diIndex
    
    ' set the global dialogue data
    dialogueData1 = Data1

    ' set the captions
    frmMain.lblDialogue_Title.Caption = diTitle
    frmMain.lblDialogue_Text.Caption = diText
    
    ' show/hide buttons
    If Not isYesNo Then
        frmMain.lblDialogue_Button(1).visible = True ' Okay button
        frmMain.lblDialogue_Button(2).visible = False ' Yes button
        frmMain.lblDialogue_Button(3).visible = False ' No button
    Else
        frmMain.lblDialogue_Button(1).visible = False ' Okay button
        frmMain.lblDialogue_Button(2).visible = True ' Yes button
        frmMain.lblDialogue_Button(3).visible = True ' No button
    End If
    
    ' show the dialogue box
    frmMain.picDialogue.visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
Dim Buffer As New clsBuffer
    Set Buffer = New clsBuffer
    
    ' find out which button
    If Index = 1 Then ' okay button
        ' dialogue index
        Select Case dialogueIndex
        
        End Select
    ElseIf Index = 2 Then ' yes button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendAcceptTradeRequest
            Case DIALOGUE_TYPE_FORGET
                ForgetSpell dialogueData1
            Case DIALOGUE_TYPE_PARTY
                SendAcceptParty
            Case DIALOGUE_LOOT_ITEM
                ' send the packet
                Player(MyIndex).MapGetTimer = GetTickCount
                Buffer.WriteLong CMapGetItem
                SendData Buffer.ToArray()
        End Select
    ElseIf Index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                SendDeclineParty
        End Select
    End If
End Sub

Public Function ConvertMapX(ByVal x As Long) As Long
    ConvertMapX = x - (TileView.left * PIC_X) - Camera.left
End Function

Public Function ConvertMapY(ByVal y As Long) As Long
    ConvertMapY = y - (TileView.top * PIC_Y) - Camera.top
End Function

Public Sub UpdateCamera()
Dim offsetX As Long, offsetY As Long, StartX As Long, StartY As Long, EndX As Long, EndY As Long

    offsetX = Player(MyIndex).xOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y
    StartX = GetPlayerX(MyIndex) - ((MAX_MAPX + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((MAX_MAPY + 1) \ 2) - 1

    If StartX < 0 Then
        offsetX = 0

        If StartX = -1 Then
            If Player(MyIndex).xOffset > 0 Then
                offsetX = Player(MyIndex).xOffset
            End If
        End If

        StartX = 0
    End If

    If StartY < 0 Then
        offsetY = 0

        If StartY = -1 Then
            If Player(MyIndex).yOffset > 0 Then
                offsetY = Player(MyIndex).yOffset
            End If
        End If

        StartY = 0
    End If

    EndX = StartX + (MAX_MAPX + 1) + 1
    EndY = StartY + (MAX_MAPY + 1) + 1

    If EndX > Map.MaxX Then
        offsetX = 32

        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).xOffset < 0 Then
                offsetX = Player(MyIndex).xOffset + PIC_X
            End If
        End If

        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If

    If EndY > Map.MaxY Then
        offsetY = 32

        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).yOffset < 0 Then
                offsetY = Player(MyIndex).yOffset + PIC_Y
            End If
        End If

        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .top = StartY
        .bottom = EndY
        .left = StartX
        .Right = EndX
    End With

    With Camera
        .top = offsetY
        .bottom = .top + ScreenY
        .left = offsetX
        .Right = .left + ScreenX
    End With

    CurX = TileView.left + ((GlobalX + Camera.left) \ PIC_X)
    CurY = TileView.top + ((GlobalY + Camera.top) \ PIC_Y)
    GlobalX_Map = GlobalX + (TileView.left * PIC_X) + Camera.left
    GlobalY_Map = GlobalY + (TileView.top * PIC_Y) + Camera.top
End Sub

Public Function IsBankItem(ByVal x As Single, ByVal y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsBankItem = 0
    
    For i = 1 To MAX_BANK
        If Not emptySlot Then
            If GetBankItemNum(i) <= 0 And GetBankItemNum(i) > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
            With tempRec
                .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .bottom = .top + PIC_Y
                .left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .left + PIC_X
            End With
            
            If x >= tempRec.left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.bottom Then
                    
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsShopItem(ByVal x As Single, ByVal y As Single) As Long
Dim i As Long, top As Long, left As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For i = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            top = GUIWindow(GUI_SHOP).y + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            left = GUIWindow(GUI_SHOP).x + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))

            If x >= left And x <= left + 32 Then
                If y >= top And y <= top + 32 Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsEqItem(ByVal x As Single, ByVal y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsEqItem = 0

    For i = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, i) > 0 And GetPlayerEquipment(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .top = GUIWindow(GUI_CHARACTER).y + EqTop
                .bottom = .top + PIC_Y
                .left = GUIWindow(GUI_CHARACTER).x + EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .left + PIC_X
            End With

            If x >= tempRec.left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsEqItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsInvItem(ByVal x As Single, ByVal y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV
        
        If Not emptySlot Then
            If GetPlayerInvItemNum(MyIndex, i) <= 0 Or GetPlayerInvItemNum(MyIndex, i) > MAX_ITEMS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .top = GUIWindow(GUI_INVENTORY).y + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .bottom = .top + PIC_Y
                .left = GUIWindow(GUI_INVENTORY).x + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .left + PIC_X
            End With
    
            If x >= tempRec.left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsPlayerSpell(ByVal x As Single, ByVal y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If Not emptySlot Then
            If PlayerSpells(i).Spell <= 0 And PlayerSpells(i).Spell > MAX_PLAYER_SPELLS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .top = GUIWindow(GUI_SPELLS).y + SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .bottom = .top + PIC_Y
                .left = GUIWindow(GUI_SPELLS).x + SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .left + PIC_X
            End With
    
            If x >= tempRec.left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsTradeItem(ByVal x As Single, ByVal y As Single, ByVal Yours As Boolean) As Long
    Dim tempRec As RECT
    Dim i As Long
    Dim itemnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV
    
        If Yours Then
            itemnum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        Else
            itemnum = TradeTheirOffer(i).num
        End If

        If itemnum > 0 And itemnum <= MAX_ITEMS Then

            With tempRec
                .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .bottom = .top + PIC_Y
                .left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .left + PIC_X
            End With

            If x >= tempRec.left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTradeItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CensorWord(ByVal sString As String) As String
    CensorWord = String(Len(sString), "*")
End Function

Public Sub placeAutotile(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(x, y).Layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .x = autoInner(1).x
                .y = autoInner(1).y
            Case "b"
                .x = autoInner(2).x
                .y = autoInner(2).y
            Case "c"
                .x = autoInner(3).x
                .y = autoInner(3).y
            Case "d"
                .x = autoInner(4).x
                .y = autoInner(4).y
            Case "e"
                .x = autoNW(1).x
                .y = autoNW(1).y
            Case "f"
                .x = autoNW(2).x
                .y = autoNW(2).y
            Case "g"
                .x = autoNW(3).x
                .y = autoNW(3).y
            Case "h"
                .x = autoNW(4).x
                .y = autoNW(4).y
            Case "i"
                .x = autoNE(1).x
                .y = autoNE(1).y
            Case "j"
                .x = autoNE(2).x
                .y = autoNE(2).y
            Case "k"
                .x = autoNE(3).x
                .y = autoNE(3).y
            Case "l"
                .x = autoNE(4).x
                .y = autoNE(4).y
            Case "m"
                .x = autoSW(1).x
                .y = autoSW(1).y
            Case "n"
                .x = autoSW(2).x
                .y = autoSW(2).y
            Case "o"
                .x = autoSW(3).x
                .y = autoSW(3).y
            Case "p"
                .x = autoSW(4).x
                .y = autoSW(4).y
            Case "q"
                .x = autoSE(1).x
                .y = autoSE(1).y
            Case "r"
                .x = autoSE(2).x
                .y = autoSE(2).y
            Case "s"
                .x = autoSE(3).x
                .y = autoSE(3).y
            Case "t"
                .x = autoSE(4).x
                .y = autoSE(4).y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim x As Long, y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).x = 32
    autoInner(1).y = 0
    
    ' NE - b
    autoInner(2).x = 48
    autoInner(2).y = 0
    
    ' SW - c
    autoInner(3).x = 32
    autoInner(3).y = 16
    
    ' SE - d
    autoInner(4).x = 48
    autoInner(4).y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).x = 0
    autoNW(1).y = 32
    
    ' NE - f
    autoNW(2).x = 16
    autoNW(2).y = 32
    
    ' SW - g
    autoNW(3).x = 0
    autoNW(3).y = 48
    
    ' SE - h
    autoNW(4).x = 16
    autoNW(4).y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).x = 32
    autoNE(1).y = 32
    
    ' NE - g
    autoNE(2).x = 48
    autoNE(2).y = 32
    
    ' SW - k
    autoNE(3).x = 32
    autoNE(3).y = 48
    
    ' SE - l
    autoNE(4).x = 48
    autoNE(4).y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).x = 0
    autoSW(1).y = 64
    
    ' NE - n
    autoSW(2).x = 16
    autoSW(2).y = 64
    
    ' SW - o
    autoSW(3).x = 0
    autoSW(3).y = 80
    
    ' SE - p
    autoSW(4).x = 16
    autoSW(4).y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).x = 32
    autoSE(1).y = 64
    
    ' NE - r
    autoSE(2).x = 48
    autoSE(2).y = 64
    
    ' SW - s
    autoSE(3).x = 32
    autoSE(3).y = 80
    
    ' SE - t
    autoSE(4).x = 48
    autoSE(4).y = 80
    
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                calculateAutotile x, y, layerNum
                ' cache the rendering state of the tiles and set them
                cacheRenderState x, y, layerNum
            Next
        Next
    Next
End Sub

Public Sub cacheRenderState(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If x < 0 Or x > Map.MaxX Or y < 0 Or y > Map.MaxY Then Exit Sub

    With Map.Tile(x, y)
        ' check if the tile can be rendered
        If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > Count_Tileset Then
            Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it's a key - hide mask if key is closed
        If layerNum = MapLayer.Mask Then
            If .Type = TILE_TYPE_KEY Then
                If TempTile(x, y).DoorOpen = NO Then
                    Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_NONE
                    Exit Sub
                Else
                    Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_NORMAL
                    Exit Sub
                End If
            End If
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Or Options.noAuto = 1 Then
            ' default to... default
            Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_NORMAL
        Else
            Autotile(x, y).Layer(layerNum).renderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(x, y).Layer(layerNum).srcX(quarterNum) = (Map.Tile(x, y).Layer(layerNum).x * 32) + Autotile(x, y).Layer(layerNum).QuarterTile(quarterNum).x
                Autotile(x, y).Layer(layerNum).srcY(quarterNum) = (Map.Tile(x, y).Layer(layerNum).y * 32) + Autotile(x, y).Layer(layerNum).QuarterTile(quarterNum).y
            Next
        End If
    End With
End Sub

Public Sub calculateAutotile(ByVal x As Long, ByVal y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(x, y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(x, y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, x, y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, x, y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, x, y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, x, y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, x, y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, x, y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, x, y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, x, y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, x, y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, x, y, x - 1, y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, x, y, x, y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, x, y, x + 1, y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, x, y, x - 1, y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, x, y, x - 1, y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal x As Long, ByVal y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, x, y, x, y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, x, y, x + 1, y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, x, y, x + 1, y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, x, y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, x, y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, x, y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, x, y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If x2 < 0 Or x2 > Map.MaxX Or y2 < 0 Or y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If Map.Tile(x2, y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(x2, y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(x1, y1).Layer(layerNum).Tileset <> Map.Tile(x2, y2).Layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(x1, y1).Layer(layerNum).x <> Map.Tile(x2, y2).Layer(layerNum).x Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(x1, y1).Layer(layerNum).y <> Map.Tile(x2, y2).Layer(layerNum).y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Sub OpenNpcChat(ByVal npcNum As Long, ByVal mT As String, ByVal o1 As String, ByVal o2 As String, ByVal o3 As String, ByVal o4 As String)
    ' set the shit
    chatNpc = npcNum
    chatText = mT
    chatOpt(1) = o1
    chatOpt(2) = o2
    chatOpt(3) = o3
    chatOpt(4) = o4
    ' we're in chat now boy
    inChat = True
End Sub

Public Sub SetTutorialState(ByVal stateNum As Byte)
Dim i As Long
    Select Case stateNum
        Case 1 ' introduction
            chatText = "Ah, so you have appeared at last my dear. Please, listen to what I have to say."
            chatOpt(1) = "*sigh* I suppose I should..."
            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next
        Case 2 ' next
            chatText = "There are some important things you need to know. Here they are. To move, use W, A, S and D. To attack or to talk to someone, press CTRL. To initiate chat press ENTER."
            chatOpt(1) = "Go on..."
            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next
        Case 3 ' chatting
            chatText = "When chatting you can talk in different channels. By default you're talking in the map channel. To talk globally append an apostrophe (') to the start of your message. To perform an emote append a hyphen (-) to the start of your message."
            chatOpt(1) = "Wait, what about combat?"
            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next
        Case 4 ' combat
            chatText = "Combat can be done through melee and skills. You can melee an enemy by facing them and pressing CTRL. To use a skill you can double click it in your skill menu, double click it in the hotbar or use the number keys. (1, 2, 3, etc.)"
            chatOpt(1) = "Oh! What do stats do?"
            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next
        Case 5 ' stats
            chatText = "Strength increases damage and allows you to equip better weaponry. Endurance increases your maximum health. Intelligence increases your maximum spirit. Agility allows you to reduce damage received and also increases critical hit chances. Willpower increase regeneration abilities."
            chatOpt(1) = "Thanks. See you later."
            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next
        Case Else ' goodbye
            chatText = vbNullString
            For i = 1 To 4
                chatOpt(i) = vbNullString
            Next
            SendFinishTutorial
            inTutorial = False
            AddText "Well done, you finished the tutorial.", BrightGreen
            Exit Sub
    End Select
    ' set the state
    tutorialState = stateNum
End Sub

Public Sub setOptionsState()
    ' music
    If Options.Music = 1 Then
        Buttons(26).state = 2
        Buttons(27).state = 0
    Else
        Buttons(26).state = 0
        Buttons(27).state = 2
    End If
    
    ' sound
    If Options.sound = 1 Then
        Buttons(28).state = 2
        Buttons(29).state = 0
    Else
        Buttons(28).state = 0
        Buttons(29).state = 2
    End If
    
    ' debug
    If Options.Debug = 1 Then
        Buttons(30).state = 2
        Buttons(31).state = 0
    Else
        Buttons(30).state = 0
        Buttons(31).state = 2
    End If
    
    ' autotile
    If Options.noAuto = 0 Then
        Buttons(32).state = 2
        Buttons(33).state = 0
    Else
        Buttons(32).state = 0
        Buttons(33).state = 2
    End If
End Sub

Public Sub ScrollChatBox(ByVal direction As Byte)
    ' do a quick exit if we don't have enough text to scroll
    If totalChatLines < 8 Then
        ChatScroll = 8
        UpdateChatArray
        Exit Sub
    End If
    ' actually scroll
    If direction = 0 Then ' up
        ChatScroll = ChatScroll + 1
    Else ' down
        ChatScroll = ChatScroll - 1
    End If
    ' scrolling down
    If ChatScroll < 8 Then ChatScroll = 8
    ' scrolling up
    If ChatScroll > totalChatLines Then ChatScroll = totalChatLines
    ' update the array
    UpdateChatArray
End Sub

Public Sub ClearMapCache()
Dim i As Long, FileName As String

    For i = 1 To MAX_MAPS
        FileName = App.path & "\data files\maps\map" & i & ".map"
        If FileExist(FileName) Then
            Kill FileName
        End If
    Next
    AddText "Map cache destroyed.", BrightGreen
End Sub

Public Sub AddChatBubble(ByVal target As Long, ByVal targetType As Byte, ByVal Msg As String, ByVal colour As Long)
Dim i As Long, Index As Long

    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    
    ' default to new bubble
    Index = chatBubbleIndex
    
    ' loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If chatBubble(i).targetType = targetType Then
            If chatBubble(i).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = i
                Exit For
            End If
        End If
    Next
    
    ' set the bubble up
    With chatBubble(Index)
        .target = target
        .targetType = targetType
        .Msg = Msg
        .colour = colour
        .timer = GetTickCount
        .active = True
    End With
End Sub

Public Sub FindNearestTarget()
Dim i As Long, x As Long, y As Long, x2 As Long, y2 As Long, xDif As Long, yDif As Long
Dim bestX As Long, bestY As Long, bestIndex As Long

    x2 = GetPlayerX(MyIndex)
    y2 = GetPlayerY(MyIndex)
    
    bestX = 255
    bestY = 255
    
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then
            x = MapNpc(i).x
            y = MapNpc(i).y
            ' find the difference - x
            If x < x2 Then
                xDif = x2 - x
            ElseIf x > x2 Then
                xDif = x - x2
            Else
                xDif = 0
            End If
            ' find the difference - y
            If y < y2 Then
                yDif = y2 - y
            ElseIf y > y2 Then
                yDif = y - y2
            Else
                yDif = 0
            End If
            ' best so far?
            If (xDif + yDif) < (bestX + bestY) Then
                bestX = xDif
                bestY = yDif
                bestIndex = i
            End If
        End If
    Next
    
    ' target the best
    If bestIndex > 0 And bestIndex <> myTarget Then PlayerTarget bestIndex, TARGET_TYPE_NPC
End Sub

Public Sub FindTarget()
Dim i As Long, x As Long, y As Long

    ' check players
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            x = (GetPlayerX(i) * 32) + Player(i).xOffset + 32
            y = (GetPlayerY(i) * 32) + Player(i).yOffset + 32
            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_PLAYER
                    Exit Sub
                End If
            End If
        End If
    Next
    
    ' check npcs
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then
            x = (MapNpc(i).x * 32) + MapNpc(i).xOffset + 32
            y = (MapNpc(i).y * 32) + MapNpc(i).yOffset + 32
            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_NPC
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

Public Sub SetBarWidth(ByRef MaxWidth As Long, ByRef Width As Long)
Dim barDifference As Long
    If MaxWidth < Width Then
        ' find out the amount to increase per loop
        barDifference = ((Width - MaxWidth) / 100) * 10
        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width - barDifference
    ElseIf MaxWidth > Width Then
        ' find out the amount to increase per loop
        barDifference = ((MaxWidth - Width) / 100) * 10
        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width + barDifference
    End If
End Sub
