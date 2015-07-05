Attribute VB_Name = "modGameLogic"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Public Sub GameLoop()
    Dim Tick As Long
    Dim TickFPS As Long
    Dim FPS As Long
    Dim i As Long
    Dim WalkTimer As Long
    Dim tmr25 As Long
    Dim tmr10000 As Long

    ' *** Start GameLoop ***
    Do While InGame
        Tick = GetTickCount
        
        ' * Check surface timers *
        ' Sprites
        If tmr10000 < Tick Then
            For i = 1 To NumSprites    'Check to unload surfaces
                If SpriteTimer(i) > 0 Then 'Only update surfaces in use
                    If SpriteTimer(i) < Tick Then   'Unload the surface
                        Call ZeroMemory(ByVal VarPtr(DDSD_Sprite(i)), LenB(DDSD_Sprite(i)))
                        Set DDS_Sprite(i) = Nothing
                        SpriteTimer(i) = 0
                    End If
                End If
            Next
            
            ' Spells
            For i = 1 To NumSpells    'Check to unload surfaces
                If SpellTimer(i) > 0 Then 'Only update surfaces in use
                    If SpellTimer(i) < Tick Then   'Unload the surface
                        Call ZeroMemory(ByVal VarPtr(DDSD_Spell(i)), LenB(DDSD_Spell(i)))
                        Set DDS_Spell(i) = Nothing
                        SpellTimer(i) = 0
                    End If
                End If
            Next
            
            ' Items
            For i = 1 To NumItems    'Check to unload surfaces
                If ItemTimer(i) > 0 Then 'Only update surfaces in use
                    If ItemTimer(i) < Tick Then   'Unload the surface
                        Call ZeroMemory(ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i)))
                        Set DDS_Item(i) = Nothing
                        ItemTimer(i) = 0
                    End If
                End If
            Next
            
            tmr10000 = Tick + 10000
        End If
        
        
        If tmr25 < Tick Then
                        
            InGame = IsConnected
            
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything
            
            Call CheckInputKeys ' Check which keys were pressed

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
        
        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then
            ' Process player movements (actually move them)
            For i = 1 To PlayersOnMapHighIndex
                If Player(PlayersOnMap(i)).Moving > 0 Then
                    Call ProcessMovement(PlayersOnMap(i))
                End If
            Next
            
            ' Process npc movements (actually move them)
            For i = 1 To High_Npc_Index
                If Map.Npc(i) > 0 Then
                    If MapNpc(i).Moving > 0 Then
                        Call ProcessNpcMovement(i)
                    End If
                End If
            Next
            
            WalkTimer = Tick + 30 ' edit this value to change WalkTimer
            
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        DoEvents
        
        ' Lock fps
        Do While GetTickCount < Tick + 30
            DoEvents
            Sleep 1
        Loop
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If

    Loop
    
    frmMirage.Visible = False
    
    If isLogging Then
        isLogging = False
        frmMirage.picScreen.Visible = False
        frmMainMenu.Visible = True
        GettingMap = True
    Else
         ' Shutdown the game
        frmSendGetData.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If
End Sub

Sub ProcessMovement(ByVal Index As Long)
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
        Case DIR_DOWN
            Player(Index).YOffset = Player(Index).YOffset + MovementSpeed
        Case DIR_LEFT
            Player(Index).XOffset = Player(Index).XOffset - MovementSpeed
        Case DIR_RIGHT
            Player(Index).XOffset = Player(Index).XOffset + MovementSpeed
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).XOffset = 0 Then
        If Player(Index).YOffset = 0 Then
            Player(Index).Moving = 0
        End If
    End If

End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if NPC is walking, and if so process moving them over
    
    'If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
    
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - WALK_SPEED
            Case DIR_DOWN
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + WALK_SPEED
            Case DIR_LEFT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - WALK_SPEED
            Case DIR_RIGHT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If MapNpc(MapNpcNum).XOffset = 0 Then
            If MapNpc(MapNpcNum).YOffset = 0 Then
                MapNpc(MapNpcNum).Moving = 0
            End If
        End If
        
    'End If
End Sub

Sub CheckMapGetItem()
Dim Buffer As New clsBuffer

    Set Buffer = New clsBuffer
    
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = GetTickCount
            
            Buffer.WriteLong CMapGetItem
            
            SendData Buffer.ToArray()
            
        End If
    End If
    
    Set Buffer = Nothing
    
End Sub

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
                
                Buffer.WriteLong CAttack
                
                SendData Buffer.ToArray()
                
                Set Buffer = Nothing
                
            End If
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If
End Function

Function CanMove() As Boolean
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
           
    If DirDown Then
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
            If Map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If
            CanMove = False
            Exit Function
        End If
    End If
       
    If DirRight Then
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
End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim i As Long

    CheckDirection = False
   
    Select Case Direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
    End Select
   
    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If
                               
    ' Check to see if the key door is open or not
    If Map.Tile(X, Y).Type = TILE_TYPE_KEY Then
        ' This actually checks if its open or not
        If TempTile(X, Y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
   
    ' Check to see if a player is already on that tile
    For i = 1 To PlayersOnMapHighIndex
        If GetPlayerX(PlayersOnMap(i)) = X Then
            If GetPlayerY(PlayersOnMap(i)) = Y Then
                CheckDirection = True
                Exit Function
            End If
        End If
    Next

    ' Check to see if a npc is already on that tile
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
    
End Function

Sub CheckMovement()
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
            
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
            
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
            
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select

            If Player(MyIndex).XOffset = 0 Then
                If Player(MyIndex).YOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
            
        End If
    End If
                
'    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
'        CanMoveNow = False
'    End If
    
End Sub

Public Sub UpdateInventory()
Dim i As Long

    frmMirage.lstInv.Clear
    
    ' Show the inventory
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmMirage.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                ' Check if this item is being worn
                If GetPlayerEquipmentSlot(MyIndex, Weapon) = i Or GetPlayerEquipmentSlot(MyIndex, Armor) = i Or GetPlayerEquipmentSlot(MyIndex, Helmet) = i Or GetPlayerEquipmentSlot(MyIndex, Shield) = i Then
                    frmMirage.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                Else
                    frmMirage.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                End If
            End If
        Else
            frmMirage.lstInv.AddItem "<free inventory slot>"
        End If
    Next
    
    frmMirage.lstInv.ListIndex = 0
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
    
    ' Subtract 1 to prevent subscript out of range, because we have one array that starts
    '' at 0, and another that starts at 1
    PlayersOnMapHighIndex = PlayersOnMapHighIndex - 1
    
End Sub

Sub PlayerSearch(ByVal CurX As Integer, ByVal CurY As Integer)
Dim Buffer As clsBuffer
    
    If isInBounds Then
        
        Set Buffer = New clsBuffer
        
        Buffer.WriteLong CSearch
        Buffer.WriteLong CurX
        Buffer.WriteLong CurY
        
        SendData Buffer.ToArray()
        
        Set Buffer = Nothing
        
    End If
End Sub

Public Function isInBounds()
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If
End Function

Public Sub UpdateDrawMapName()
    DrawMapNameX = (MAX_MAPX + 1) * PIC_X \ 2 - ((Len(Trim$(Map.Name)) \ 2) * 8)
    DrawMapNameY = 1
     
    Select Case Map.Moral
        Case MAP_MORAL_NONE
            DrawMapNameColor = QBColor(BrightRed)
            
        Case MAP_MORAL_SAFE
            DrawMapNameColor = QBColor(White)
            
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
                
                Buffer.WriteLong CCast
                Buffer.WriteLong SpellSelected
                
                SendData Buffer.ToArray()
                
                Set Buffer = Nothing
                
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

Sub ClearTempTile()
Dim X As Long
Dim Y As Long

    ReDim TempTile(0 To Map.MaxX, 0 To Map.MaxY)
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            TempTile(X, Y).DoorOpen = NO
        Next
    Next
End Sub

Public Sub DevMsg(ByVal Text As String, ByVal Color As Byte)
    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(Text, Color)
        End If
    End If
    Debug.Print Text
End Sub
