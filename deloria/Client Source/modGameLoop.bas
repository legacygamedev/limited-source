Attribute VB_Name = "modGameLoop"
Sub GameLoop()
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim x As Long
Dim Y As Long
Dim i As Long
Dim rec_back As RECT
    
    ' Set the focus
    frmMirage.picScreen.SetFocus
    
    ' Set font
    'Call SetFont("Fixedsys", 18)
    GameFontSize = 16
    Call SetFont("Verdana", GameFontSize)
    GameFontSize = 8
                
    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0
    
    Do While InGame
        Tick = GetTickCount
        
        ' Check to make sure they aren't trying to auto do anything
        If GetAsyncKeyState(VK_UP) >= 0 And DirUp = True Then DirUp = False
        If GetAsyncKeyState(VK_DOWN) >= 0 And DirDown = True Then DirDown = False
        If GetAsyncKeyState(VK_LEFT) >= 0 And DirLeft = True Then DirLeft = False
        If GetAsyncKeyState(VK_RIGHT) >= 0 And DirRight = True Then DirRight = False
        If GetAsyncKeyState(VK_CONTROL) >= 0 And ControlDown = True Then ControlDown = False
        If GetAsyncKeyState(VK_SHIFT) >= 0 And ShiftDown = True Then ShiftDown = False
        
        ' Check to make sure we are still connected
        If Not IsConnected Then InGame = False
        
        ' Check if we need to restore surfaces
        If NeedToRestoreSurfaces Then
            DD.RestoreAllSurfaces
            Call InitSurfaces
        End If
                
        ' Visual Inventory
        Dim Q As Long
        Dim Qq As Long
        Dim IT As Long
               
        If GetTickCount > IT + 500 And frmMirage.picInv3.Visible = True Then
            For Q = 0 To MAX_INV - 1
                Qq = Player(MyIndex).Inv(Q + 1).Num
               
                If frmMirage.picInv(Q).Picture <> LoadPicture() Then
                    frmMirage.picInv(Q).Picture = LoadPicture()
                Else
                    If Qq = 0 Then
                        frmMirage.picInv(Q).Picture = LoadPicture()
                    Else
                        Call BitBlt(frmMirage.picInv(Q).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq).Pic - Int(Item(Qq).Pic / 6) * 6) * PIC_X, Int(Item(Qq).Pic / 6) * PIC_Y, SRCCOPY)
                    End If
                End If
            Next Q
        End If
        
        If ID = True Then Call CheckInput2
        
    If GettingMap = False Then
                        
        NewX = 8
        NewY = 6
                          
        NewPlayerY = Player(MyIndex).Y - NewY
        NewPlayerX = Player(MyIndex).x - NewX
        
        NewX = NewX * PIC_X
        NewY = NewY * PIC_Y
        
        NewXOffset = Player(MyIndex).XOffset
        NewYOffset = Player(MyIndex).YOffset
        
        If Player(MyIndex).Y - 6 < 1 Then
            If CheckMap(GetPlayerMap(MyIndex)).Up <= 0 Then
                NewY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).YOffset
                NewYOffset = 0
                NewPlayerY = 0
                If Player(MyIndex).Y = 6 And Player(MyIndex).Dir = DIR_UP Then
                    NewPlayerY = Player(MyIndex).Y - 6
                    NewY = 6 * PIC_Y
                    NewYOffset = Player(MyIndex).YOffset
                End If
            End If
        ElseIf Player(MyIndex).Y + 8 > MAX_MAPY + 1 Then
            If CheckMap(GetPlayerMap(MyIndex)).Down <= 0 Then
                NewY = (Player(MyIndex).Y - 18) * PIC_Y + Player(MyIndex).YOffset
                NewYOffset = 0
                NewPlayerY = MAX_MAPY - 12
                If Player(MyIndex).Y = 24 And Player(MyIndex).Dir = DIR_DOWN Then
                    NewPlayerY = Player(MyIndex).Y - 6
                    NewY = 6 * PIC_Y
                    NewYOffset = Player(MyIndex).YOffset
                End If
            End If
        End If
        
        If Player(MyIndex).x - 8 < 1 Then
            If CheckMap(GetPlayerMap(MyIndex)).Left <= 0 Then
                NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
                NewXOffset = 0
                NewPlayerX = 0
                If Player(MyIndex).x = 8 And Player(MyIndex).Dir = DIR_LEFT Then
                    NewPlayerX = Player(MyIndex).x - 8
                    NewX = 8 * PIC_X
                    NewXOffset = Player(MyIndex).XOffset
                End If
            End If
        ElseIf Player(MyIndex).x + 10 > MAX_MAPX + 1 Then
            If CheckMap(GetPlayerMap(MyIndex)).Right <= 0 Then
                NewX = (Player(MyIndex).x - 14) * PIC_X + Player(MyIndex).XOffset
                NewXOffset = 0
                NewPlayerX = MAX_MAPX - 16
                If Player(MyIndex).x = 22 And Player(MyIndex).Dir = DIR_RIGHT Then
                    NewPlayerX = Player(MyIndex).x - 8
                    NewX = 8 * PIC_X
                    NewXOffset = Player(MyIndex).XOffset
                End If
            End If
        End If
        
        sx = 32
        If MAX_MAPX = 19 Then
            NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = 0
            NewY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = 0
            sx = 0
        End If
        
        ' Blit out tiles layers ground/anim1/anim2
        For Y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If CheckMap(GetPlayerMap(MyIndex)).Up > 0 Then
                    If MapsAvailable(CheckMap(GetPlayerMap(MyIndex)).Up) = True Then
                        If GetPlayerY(MyIndex) - 7 < 0 Then
                            If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Left > 0 Then Call BltTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Left, x, Y, 0)
                            If CheckMap(GetPlayerMap(MyIndex)).Left > 0 And CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Left <= 0 Then
                                If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Left).Up > 0 Then Call BltTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Left).Up, x, Y, 0)
                            End If
                            
                            Call BltTile(CheckMap(GetPlayerMap(MyIndex)).Up, x, Y, 1)
                            
                            If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Right > 0 Then Call BltTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Right, x, Y, 2)
                            If CheckMap(GetPlayerMap(MyIndex)).Right > 0 And CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Right <= 0 Then
                                If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Right).Up > 0 Then Call BltTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Right).Up, x, Y, 2)
                            End If
                        End If
                    End If
                End If
                
                If CheckMap(GetPlayerMap(MyIndex)).Left > 0 Then
                    If MapsAvailable(CheckMap(GetPlayerMap(MyIndex)).Left) = True Then
                        If GetPlayerX(MyIndex) - 9 < 0 Then
                            Call BltTile(CheckMap(GetPlayerMap(MyIndex)).Left, x, Y, 3)
                        End If
                    End If
                End If
                
                Call BltTile(GetPlayerMap(MyIndex), x, Y, 4)
                
                If CheckMap(GetPlayerMap(MyIndex)).Right > 0 Then
                    If MapsAvailable(CheckMap(GetPlayerMap(MyIndex)).Right) = True Then
                        If GetPlayerX(MyIndex) + 9 > MAX_MAPX Then
                            Call BltTile(CheckMap(GetPlayerMap(MyIndex)).Right, x, Y, 5)
                        End If
                    End If
                End If
                
                If CheckMap(GetPlayerMap(MyIndex)).Down > 0 Then
                    If MapsAvailable(CheckMap(GetPlayerMap(MyIndex)).Down) = True Then
                        If GetPlayerY(MyIndex) + 7 > MAX_MAPX Then
                            If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Left > 0 Then Call BltTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Left, x, Y, 6)
                            
                            If CheckMap(GetPlayerMap(MyIndex)).Left > 0 And CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Left <= 0 Then
                                If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Left).Down > 0 Then Call BltTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Left).Down, x, Y, 6)
                            End If
                            
                            Call BltTile(CheckMap(GetPlayerMap(MyIndex)).Down, x, Y, 7)
                            
                            If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Right > 0 Then Call BltTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Right, x, Y, 8)
                            If CheckMap(GetPlayerMap(MyIndex)).Right > 0 And CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Right <= 0 Then
                                If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Right).Down > 0 Then Call BltTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Right).Down, x, Y, 8)
                            End If
                        End If
                    End If
                End If
            Next x
        Next Y
            
       
        ' Blit out the items
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call BltItem(i)
            End If
        Next i
        
    If ScreenMode = 0 Then
        ' Blit out the sprite change attribute
        For Y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltSpriteChange(x, Y)
            Next x
        Next Y

        ' Blit out players
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltPlayer(i)
            End If
        Next i
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpc(i)
        Next i
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(i).Y - 1 >= 0 Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).Y - 1).Fringe > 1 Or CheckMap(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).Y - 1).FAnim > 1 Or CheckMap(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).Y - 1).Fringe2 > 1 Or CheckMap(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).Y - 1).F2Anim > 1 Then
                    Call BltNpcTop(i)
                End If
            End If
        Next i
        
        ' Blit out the sprite change attribute
        For Y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltSpriteChange2(x, Y)
            Next x
        Next Y
    End If
                
        ' Blit out tile layer fringe
        For Y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If CheckMap(GetPlayerMap(MyIndex)).Up > 0 Then
                    If MapsAvailable(CheckMap(GetPlayerMap(MyIndex)).Up) = True Then
                        If GetPlayerY(MyIndex) - 7 < 0 Then
                            If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Left > 0 Then Call BltFringeTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Left, x, Y, 0)
                            If CheckMap(GetPlayerMap(MyIndex)).Left > 0 And CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Left <= 0 Then
                                If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Left).Up > 0 Then Call BltFringeTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Left).Up, x, Y, 0)
                            End If
                            
                            Call BltFringeTile(CheckMap(GetPlayerMap(MyIndex)).Up, x, Y, 1)
                            
                            If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Right > 0 Then Call BltFringeTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Right, x, Y, 2)
                            If CheckMap(GetPlayerMap(MyIndex)).Right > 0 And CheckMap(CheckMap(GetPlayerMap(MyIndex)).Up).Right <= 0 Then
                                If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Right).Up > 0 Then Call BltFringeTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Right).Up, x, Y, 2)
                            End If
                        End If
                    End If
                End If
                
                If CheckMap(GetPlayerMap(MyIndex)).Left > 0 Then
                    If MapsAvailable(CheckMap(GetPlayerMap(MyIndex)).Left) = True Then
                        If GetPlayerX(MyIndex) - 9 < 0 Then
                            Call BltFringeTile(CheckMap(GetPlayerMap(MyIndex)).Left, x, Y, 3)
                        End If
                    End If
                End If
                
                Call BltFringeTile(GetPlayerMap(MyIndex), x, Y, 4)
                
                If CheckMap(GetPlayerMap(MyIndex)).Right > 0 Then
                    If MapsAvailable(CheckMap(GetPlayerMap(MyIndex)).Right) = True Then
                        If GetPlayerX(MyIndex) + 9 > MAX_MAPX Then
                            Call BltFringeTile(CheckMap(GetPlayerMap(MyIndex)).Right, x, Y, 5)
                        End If
                    End If
                End If
                
                If CheckMap(GetPlayerMap(MyIndex)).Down > 0 Then
                    If MapsAvailable(CheckMap(GetPlayerMap(MyIndex)).Down) = True Then
                        If GetPlayerY(MyIndex) + 7 > MAX_MAPX Then
                            If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Left > 0 Then Call BltFringeTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Left, x, Y, 6)
                            
                            If CheckMap(GetPlayerMap(MyIndex)).Left > 0 And CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Left <= 0 Then
                                If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Left).Down > 0 Then Call BltFringeTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Left).Down, x, Y, 6)
                            End If
                            
                            Call BltFringeTile(CheckMap(GetPlayerMap(MyIndex)).Down, x, Y, 7)
                            
                            If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Right > 0 Then Call BltFringeTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Right, x, Y, 8)
                            If CheckMap(GetPlayerMap(MyIndex)).Right > 0 And CheckMap(CheckMap(GetPlayerMap(MyIndex)).Down).Right <= 0 Then
                                If CheckMap(CheckMap(GetPlayerMap(MyIndex)).Right).Down > 0 Then Call BltFringeTile(CheckMap(CheckMap(GetPlayerMap(MyIndex)).Right).Down, x, Y, 8)
                            End If
                        End If
                    End If
                End If
            Next x
        Next Y
      
    If ScreenMode = 0 Then
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            If CheckMap(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).Y).Fringe < 1 Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).Y).FAnim < 1 Then
                    If CheckMap(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).Y).Fringe2 < 1 Then
                        If CheckMap(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).Y).F2Anim < 1 Then
                            Call BltNpcTop(i)
                        End If
                    End If
                End If
            End If
        Next i
    End If

        If InEditor = True And frmMapEditor.optMapGrid.Value = Checked Then
            For Y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Call BltTile2(x * 32, Y * 32, 1)
                Next x
            Next Y
        End If
        
        If frmMirage.chkWeather.Value = Checked Then
            If CheckMap(GetPlayerMap(MyIndex)).Indoor = 0 Then Call BltWeather
        End If
    End If

        ' Lock the backbuffer so we can draw text and names
        TexthDC = DD_BackBuffer.GetDC
        
    If GettingMap = False Then
        If frmMirage.chkTime.Value = Checked Then
            If CheckMap(GetPlayerMap(MyIndex)).Indoor = 0 Then
                Dim fadenight As Long
                If GameTime = TIME_NIGHT Then
                    If fadenight <> 150 Then
                        fadenight = fadenight + 5
                        If fadenight > 150 Then fadenight = 100
                    End If
                Else
                    If fadenight <> 0 Then
                        fadenight = fadenight - 5
                        If fadenight < 0 Then fadenight = 0
                    End If
                End If
                Call SquareAlphaBlend(frmMirage.picScreen.Width, frmMirage.picScreen.Height, frmMirage.picDayNight.hDC, 0, 0, TexthDC, (PIC_X * (MAX_MAPX + 1)) + 32, (PIC_Y * (MAX_MAPY + 1)) + 32, fadenight)
            End If
        End If
        
        If ScreenMode = 0 Then
            If frmMirage.chkplayername.Value = Checked Then
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        Call BltPlayerGuildName(i)
                        Call BltPlayerName(i)
                    End If
                Next i
            End If
                            
            ' Blit out attribs if in editor
            If InEditor Then
                For Y = 0 To MAX_MAPY
                    For x = 0 To MAX_MAPX
                        With CheckMap(GetPlayerMap(MyIndex)).Tile(x, Y)
                            If .Type = TILE_TYPE_BLOCKED Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "B", QBColor(BrightRed))
                            If .Type = TILE_TYPE_WARP Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "W", QBColor(BrightBlue))
                            If .Type = TILE_TYPE_ITEM Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "I", QBColor(White))
                            If .Type = TILE_TYPE_NPCAVOID Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "N", QBColor(White))
                            If .Type = TILE_TYPE_KEY Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "K", QBColor(White))
                            If .Type = TILE_TYPE_KEYOPEN Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "O", QBColor(White))
                            If .Type = TILE_TYPE_HEAL Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "H", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_KILL Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "K", QBColor(BrightRed))
                            If .Type = TILE_TYPE_SHOP Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "S", QBColor(Yellow))
                            If .Type = TILE_TYPE_CBLOCK Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "CB", QBColor(Black))
                            If .Type = TILE_TYPE_ARENA Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "A", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_SOUND Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "PS", QBColor(Yellow))
                            If .Type = TILE_TYPE_SPRITE_CHANGE Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "SC", QBColor(Grey))
                            If .Type = TILE_TYPE_SIGN Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "SI", QBColor(Yellow))
                            If .Type = TILE_TYPE_DOOR Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "D", QBColor(Black))
                            If .Type = TILE_TYPE_NOTICE Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "N", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_CHEST Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "C", QBColor(Brown))
                            If .Type = TILE_TYPE_CLASS_CHANGE Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "CG", QBColor(White))
                            If .Type = TILE_TYPE_SCRIPTED Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "SC", QBColor(Yellow))
                            If .Type = TILE_TYPE_SEX_BLOCK Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 2 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "SEX", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_LEVEL_BLOCK Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 2 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "LVL", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_BANK Then Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx + 1 - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset), "BANK", QBColor(BrightRed))
                        End With
                    Next x
                Next Y
            End If
    
            ' Blit the text they are putting in
            MyText = frmMirage.txtMyTextBox.Text
            
            Dim s As String
            s = Trim(CheckMap(GetPlayerMap(MyIndex)).Name)
            If Len(s) > 6 Then
                If Trim(CheckMap(GetPlayerMap(MyIndex)).Name) <> "" Then If LCase(Mid(s, Len(s) - 6, Len(s))) = "colling" Then s = CheckMap(GetPlayerMap(MyIndex)).Name & "wood"
            End If
            frmMirage.lblMapName.Caption = Trim(s)
            If CheckMap(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NONE Then
                frmMirage.lblMapName.ForeColor = RGB(255, 0, 0)
            ElseIf CheckMap(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_SAFE Then
                frmMirage.lblMapName.ForeColor = &HFFFFC0
            ElseIf CheckMap(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NO_PENALTY Then
                frmMirage.lblMapName.ForeColor = RGB(240, 240, 240)
            ElseIf CheckMap(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_TRAINING Then
                frmMirage.lblMapName.ForeColor = &HFFFFC0
            End If
            
            'Draw party name text
            If PartyMems(0).Name <> "" Then
                Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (5 + sx), (PIC_Y * (MAX_MAPY + 1)) + sx, "Party Members:", QBColor(BrightCyan))
                For i = 0 To MAX_PARTY_MEMS
                    If PartyMems(i).Name = "" Then
                        Exit For
                    Else
                        Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (5 + sx), (PIC_Y * (MAX_MAPY + 1)) + sx + PartyMems(i).Y, Trim(PartyMems(i).Name), QBColor(BrightCyan))
                    End If
                Next i
            End If
            
            For i = 1 To MAX_BLT_LINE
                If BattleMsg(i).Index > 0 Then
                    If BattleMsg(i).Time + 7000 > GetTickCount Then
                        Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + (5 + sx), (PIC_Y * (MAX_MAPY + 1)) + (BattleMsg(i).Y + frmMirage.picScreen.Height - 20 + sx), Trim(BattleMsg(i).Msg), QBColor(Yellow))
                    Else
                        BattleMsg(i).Done = 0
                    End If
                End If
            Next i
        End If
    End If
    
        If frmMirage.chkFPS.Value = Checked Then
            Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + ((frmMirage.picScreen.Width - (Len("FPS:" & GameFPS) * GameFontSize)) + sx - 1), (PIC_Y * (MAX_MAPY + 1)) + (frmMirage.picScreen.Height - 14 + sx), "FPS:" & GameFPS, QBColor(White))
        End If

        ' Check if we are getting a map, and if we are tell them so
        If GettingMap = True Then
            Call DrawText(TexthDC, (PIC_X * (MAX_MAPX + 1)) + 36, (PIC_Y * (MAX_MAPY + 1)) + 36, "Receiving Map...", QBColor(BrightCyan))
        End If
                        
        ' Release DC
        Call DD_BackBuffer.ReleaseDC(TexthDC)
        
        ' Blit out emoticons
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltEmoticons(i)
            End If
        Next i
        
        ' Get the rect for the back buffer to blit from
        rec.Top = (PIC_Y * (MAX_MAPY + 1))
        rec.Bottom = rec.Top + (MAX_MAPY + 1) * PIC_Y
        rec.Left = (PIC_X * (MAX_MAPX + 1))
        rec.Right = rec.Left + (MAX_MAPX + 1) * PIC_X
        
        ' Get the rect to blit to
        Call DX.GetWindowRect(frmMirage.picScreen.hWnd, rec_pos)
        rec_pos.Bottom = rec_pos.Top - sx + ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Right = rec_pos.Left - sx + ((MAX_MAPX + 1) * PIC_X)
        rec_pos.Top = rec_pos.Bottom - ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Left = rec_pos.Right - ((MAX_MAPX + 1) * PIC_X)
        
        ' Blit the backbuffer
        Call DD_PrimarySurf.Blt(rec_pos, DD_BackBuffer, rec, DDBLT_WAIT)
        
        ' Check if player is trying to move
        Call CheckMovement
        
        ' Check to see if player is trying to attack
        Call CheckAttack
        
        ' Process player movements (actually move them)
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call ProcessMovement(i)
            End If
        Next i
        
        ' Process npc movements (actually move them)
        For i = 1 To MAX_MAP_NPCS
            If CheckMap(GetPlayerMap(MyIndex)).Npc(i) > 0 Then
                Call ProcessNpcMovement(i)
            End If
        Next i
  
        ' Change map animation every 250 milliseconds
        If GetTickCount > MapAnimTimer + 250 Then
            If MapAnim = 0 Then
                MapAnim = 1
            Else
                MapAnim = 0
            End If
            MapAnimTimer = GetTickCount
        End If
                
        ' Lock fps
        Do While GetTickCount < Tick + 35
            DoEvents
        Loop
        
        ' Calculate fps
        If GetTickCount > TickFPS + 1000 Then
            GameFPS = FPS
            TickFPS = GetTickCount
            FPS = 0
        Else
            FPS = FPS + 1
        End If
        
        Call MakeMidiLoop
        
        DoEvents
    Loop
    
    frmMirage.Visible = False
    Call SetStatus("Destroying game data...")
    
    ' Shutdown the game
    Call GameDestroy
    
    ' Report disconnection if server disconnects
    If IsConnected = False Then
        Call MsgBox("Thank you for playing " & GAME_NAME & "!", vbOKOnly, GAME_NAME)
    End If
End Sub

