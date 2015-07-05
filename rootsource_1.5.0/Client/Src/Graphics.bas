Attribute VB_Name = "Graphics"
Option Explicit

Public Sub Render_Graphics()
Dim x As Long
Dim y As Long
Dim i As Long
Dim n As Long

    DX8.BeginScene
    
        For x = -1 To MAX_MAPX + 1
            For y = -1 To MAX_MAPY + 1
                Call BltMapTile(x, y)
            Next
        Next
        
        For n = 1 To 9
            If (tMap(n) <> 0) And (tMap(n) < MAX_MAPS + 1) Then
                For i = 1 To MAX_MAP_ITEMS
                    If MapItem(i, tMap(n)).Num > 0 Then
                        Call BltItem(i, tMap(n))
                    End If
                Next
            End If
        Next
        
        For y = -1 To MAX_MAPY + 1
            For i = 1 To High_Index
                If Player(i).y = y Then
                    Call BltPlayer(i)
                End If
            Next
            
            For n = 1 To 9
                If (tMap(n) <> 0) And (tMap(n) < MAX_MAPS + 1) Then
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(i, tMap(n)).y = y Then
                            Call BltNpc(i, tMap(n))
                        End If
                    Next
                End If
            Next
        Next
        
        For x = -1 To MAX_MAPX + 1
            For y = -1 To MAX_MAPY + 1
                Call BltMapFringeTile(x, y)
            Next
        Next
        
        If BFPS Then
            Call DX8.DrawText((MAX_MAPX + 1) * PIC_X - Len(Trim$("FPS: " & GameFPS)) * 10 - 3, 3, Trim$("FPS: " & GameFPS), DX8.ARGB(255, 255, 255, 255))
        End If

        
        For i = 1 To High_Index
            Call DrawPlayerName(i)
            Call DrawPlayerGuildName(i)
        Next
        
        ' Draw map name
        Call DX8.DrawText(DrawMapNameX, DrawMapNameY, map(5).Name, DX8.AcRGB(255, DrawMapNameColor))
    
        ' draw cursor and player location
        If BLoc Then
            Call DX8.DrawText(3, 3, Trim$("Cur X: " & CurX & " Y: " & CurY), DX8.ARGB(255, 255, 255, 255))
            Call DX8.DrawText(3, 14, Trim$("Loc X: " & GetPlayerX(MyIndex) & " Y: " & GetPlayerY(MyIndex)), DX8.ARGB(255, 255, 255, 255))
            Call DX8.DrawText(3, 25, Trim$("Map#:" & GetPlayerMap(MyIndex)), DX8.ARGB(255, 255, 255, 255))
        End If
                        
        ' draw map attributes
        If Editor = EDITOR_MAP Then
            If frmMainGame.optAttribs.Value Then
                Call DrawMapAttributes
            End If
        End If

    Call DX8.EndScene

End Sub

Public Sub BltMapTile(ByVal x As Long, ByVal y As Long)
    ' Blit out ground tile without transparency
    
    Call DX8.SetTexture(Tr_Tiles(MapTilePosition(x, y).Layer(0).ypos \ 256))
    Call DX8.DrawTexture(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, PIC_X, PIC_Y, MapTilePosition(x, y).Layer(0).xpos, MapTilePosition(x, y).Layer(0).ypos Mod 256, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
    
    'Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(0), DDBLTFAST_WAIT)

    If MapAnim = 0 Or ((MapTilePosition(x, y).Layer(2).ypos = 0) And (MapTilePosition(x, y).Layer(2).xpos = 0)) Then
        If ((MapTilePosition(x, y).Layer(1).ypos <> 0) Or (MapTilePosition(x, y).Layer(1).xpos <> 0)) Then
            If TempTile(x + NewPlayerX, y + NewPlayerY).DoorOpen = NO Then
                Call DX8.SetTexture(Tr_Tiles(MapTilePosition(x, y).Layer(1).ypos \ 256))
                Call DX8.DrawTexture(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, PIC_X, PIC_Y, MapTilePosition(x, y).Layer(1).xpos, MapTilePosition(x, y).Layer(1).ypos Mod 256, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
            End If
        End If
    Else
        Call DX8.SetTexture(Tr_Tiles(MapTilePosition(x, y).Layer(2).ypos \ 256))
        Call DX8.DrawTexture(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, PIC_X, PIC_Y, MapTilePosition(x, y).Layer(2).xpos, MapTilePosition(x, y).Layer(2).ypos Mod 256, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
    End If
   
    If MapAnim = 0 Or ((MapTilePosition(x, y).Layer(4).ypos = 0) And (MapTilePosition(x, y).Layer(4).xpos = 0)) Then
        If ((MapTilePosition(x, y).Layer(3).ypos <> 0) Or (MapTilePosition(x, y).Layer(3).xpos <> 0)) Then
            Call DX8.SetTexture(Tr_Tiles(MapTilePosition(x, y).Layer(3).ypos \ 256))
            Call DX8.DrawTexture(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, PIC_X, PIC_Y, MapTilePosition(x, y).Layer(3).xpos, MapTilePosition(x, y).Layer(3).ypos Mod 256, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
        End If
    Else
        Call DX8.SetTexture(Tr_Tiles(MapTilePosition(x, y).Layer(4).ypos \ 256))
        Call DX8.DrawTexture(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, PIC_X, PIC_Y, MapTilePosition(x, y).Layer(4).xpos, MapTilePosition(x, y).Layer(4).ypos Mod 256, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
    End If

    
End Sub

Public Sub BltMapFringeTile(ByVal x As Long, ByVal y As Long)
    If MapAnim = 0 Or ((MapTilePosition(x, y).Layer(6).ypos = 0) And (MapTilePosition(x, y).Layer(6).xpos = 0)) Then
        If ((MapTilePosition(x, y).Layer(5).ypos <> 0) Or (MapTilePosition(x, y).Layer(5).xpos <> 0)) Then
            Call DX8.SetTexture(Tr_Tiles(MapTilePosition(x, y).Layer(5).ypos \ 256))
            Call DX8.DrawTexture(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, PIC_X, PIC_Y, MapTilePosition(x, y).Layer(5).xpos, MapTilePosition(x, y).Layer(5).ypos Mod 256, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
        End If
    Else
        Call DX8.SetTexture(Tr_Tiles(MapTilePosition(x, y).Layer(6).ypos \ 256))
        Call DX8.DrawTexture(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, PIC_X, PIC_Y, MapTilePosition(x, y).Layer(6).xpos, MapTilePosition(x, y).Layer(6).ypos Mod 256, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
    End If
    
    If MapAnim = 0 Or ((MapTilePosition(x, y).Layer(8).ypos = 0) And (MapTilePosition(x, y).Layer(8).xpos = 0)) Then
        If ((MapTilePosition(x, y).Layer(7).ypos <> 0) Or (MapTilePosition(x, y).Layer(7).xpos <> 0)) Then
            Call DX8.SetTexture(Tr_Tiles(MapTilePosition(x, y).Layer(7).ypos \ 256))
            Call DX8.DrawTexture(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, PIC_X, PIC_Y, MapTilePosition(x, y).Layer(7).xpos, MapTilePosition(x, y).Layer(7).ypos Mod 256, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
        End If
    Else
        Call DX8.SetTexture(Tr_Tiles(MapTilePosition(x, y).Layer(8).ypos \ 256))
        Call DX8.DrawTexture(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, PIC_X, PIC_Y, MapTilePosition(x, y).Layer(8).xpos, MapTilePosition(x, y).Layer(8).ypos Mod 256, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
    End If
End Sub

Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim x2 As Long, y2 As Long
Dim i As Long
Dim insight As Boolean
Dim Sprite As Long, spriteleft As Long
Dim spellnum As Long


    Sprite = GetPlayerSprite(Index)
    If Sprite = 0 Then Sprite = 1
    DX8.SetTexture Tr_Sprites(Sprite)

            insight = False
                Select Case Player(Index).map
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
                
            If insight = True Then
                Anim = 0
                If Player(Index).Attacking = 0 Then
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            If Player(Index).YOffset < PIC_Y \ 2 Then Anim = 1
                        Case DIR_DOWN
                            If Player(Index).YOffset < PIC_Y \ -2 Then Anim = 1
                        Case DIR_LEFT
                            If Player(Index).XOffset < PIC_Y \ 2 Then Anim = 1
                        Case DIR_RIGHT
                            If Player(Index).XOffset < PIC_Y \ -2 Then Anim = 1
                    End Select
                Else
                    If Player(Index).AttackTimer + 500 > GetTickCount Then
                        Anim = 2
                    End If
                End If
                
                If Player(Index).AttackTimer + 1000 < GetTickCount Then
                    Player(Index).Attacking = 0
                    Player(Index).AttackTimer = 0
                End If
                
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spriteleft = DIR_UP
        Case DIR_RIGHT
            spriteleft = DIR_RIGHT
        Case DIR_DOWN
            spriteleft = DIR_DOWN
        Case DIR_LEFT
            spriteleft = DIR_LEFT
    End Select
                
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset

    Call DX8.DrawTexture(x2 * PIC_X + x + StaticX, y2 * PIC_Y + y + StaticY, PIC_X, PIC_Y, (spriteleft * 3 + Anim) * (32), 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 512, 32)
    
        For i = 1 To MAX_SPELLANIM
        spellnum = Player(Index).SpellAnimations(i).spellnum
        
        If spellnum > 0 Then
            If Spell(spellnum).Pic > 0 Then
        
                If Player(Index).SpellAnimations(i).Timer < GetTickCount Then
                    Player(Index).SpellAnimations(i).FramePointer = Player(Index).SpellAnimations(i).FramePointer + 1
                    Player(Index).SpellAnimations(i).Timer = GetTickCount + 120
                                                                            
                    If Player(Index).SpellAnimations(i).FramePointer >= 12 Then
                        Player(Index).SpellAnimations(i).spellnum = 0
                        Player(Index).SpellAnimations(i).Timer = 0
                        Player(Index).SpellAnimations(i).FramePointer = 0
                    End If
                End If
            
                If Player(Index).SpellAnimations(i).spellnum > 0 Then
            
                    Call DX8.SetTexture(Tr_Spells(spellnum))
                    Call DX8.DrawTexture(x, y, PIC_X, PIC_Y, Player(Index).SpellAnimations(i).FramePointer * SIZE_X, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 512, 32)

                End If
                
            End If
        End If
        
    Next
    End If
End Sub

Public Sub BltNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Anim As Long
Dim i As Long
Dim spellnum As Long
Dim x, x2 As Long
Dim y, y2 As Long
Dim Sprite As Long, spriteleft As Long
Dim insight As Boolean

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum, MapNum).Num <= 0 Then
        Exit Sub
    End If

    Sprite = Npc(MapNpc(MapNpcNum, MapNum).Num).Sprite
    DX8.SetTexture Tr_Sprites(Sprite)
    
            insight = False
                Select Case MapNum
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
                
If insight Then
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum, MapNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum, MapNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum, MapNum).YOffset < SIZE_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum, MapNum).YOffset < SIZE_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum, MapNum).XOffset < SIZE_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum, MapNum).XOffset < SIZE_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum, MapNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum, MapNum)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    Select Case MapNpc(MapNpcNum, MapNum).Dir
        Case DIR_UP
            spriteleft = DIR_UP
        Case DIR_RIGHT
            spriteleft = DIR_RIGHT
        Case DIR_DOWN
            spriteleft = DIR_DOWN
        Case DIR_LEFT
            spriteleft = DIR_LEFT
    End Select
   
    'With rec
    '    .Top = 0
    '    .Bottom = DDS_Sprite(Sprite).SurfDescription.lHeight
    '    .Left = (spriteleft * 3 + Anim) * (DDS_Sprite(Sprite).SurfDescription.lWidth / 12)
    '    .Right = .Left + (DDS_Sprite(Sprite).SurfDescription.lWidth / 12)
    'End With
    
    With MapNpc(MapNpcNum, MapNum)
        x = .x * PIC_X + .XOffset
        y = .y * PIC_Y + .YOffset
    End With

    Call DX8.DrawTexture(x2 * PIC_X + x + StaticX, y2 * PIC_Y + y + StaticY, PIC_X, PIC_Y, (spriteleft * 3 + Anim) * (32), 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 512, 32)
    
    
    ' ** Blit Spells Animations **
    For i = 1 To MAX_SPELLANIM
        spellnum = MapNpc(MapNpcNum, MapNum).SpellAnimations(i).spellnum
        
        If spellnum > 0 Then
            If Spell(spellnum).Pic > 0 Then
            
                If MapNpc(MapNpcNum, MapNum).SpellAnimations(i).Timer < GetTickCount Then
                    MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer = MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer + 1
                    MapNpc(MapNpcNum, MapNum).SpellAnimations(i).Timer = GetTickCount + 120
                    
                    If MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer >= 12 Then
                        MapNpc(MapNpcNum, MapNum).SpellAnimations(i).spellnum = 0
                        MapNpc(MapNpcNum, MapNum).SpellAnimations(i).Timer = 0
                        MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer = 0
                    End If
                End If
            
                If MapNpc(MapNpcNum, MapNum).SpellAnimations(i).spellnum > 0 Then
                    Call DX8.SetTexture(Tr_Spells(spellnum))
                    Call DX8.DrawTexture(x + x2 * PIC_X, y + y2 * PIC_Y, PIC_X, PIC_Y, MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer * SIZE_X, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 512, 32)
                End If
                
            End If
        End If
        
    Next
    End If
End Sub

Public Sub DrawNewChar()
Dim rec As D3DRECT
Dim ListIndexSprite As Long

    With rec
        .y1 = 0
        .y2 = 32
        .x1 = 0
        .x2 = 32
    End With
   
    ListIndexSprite = frmMainMenu.cmbClass.ListIndex
    If ListIndexSprite < 0 Then Exit Sub
    If Class(ListIndexSprite + 1).Sprite = 0 Then Exit Sub
    
    Call DX8.BeginScene
    
        Call DX8.SetTexture(Tr_Sprites(Class(ListIndexSprite + 1).Sprite))
        Call DX8.DrawTexture(0, 0, PIC_X, PIC_Y, 3 * PIC_X, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 512, 32)
    
    Call DX8.EndSceneSp(frmMainMenu.picPic.hWnd, rec)

End Sub

Public Sub DrawSelChar(ByVal i As Long)
Dim rec As D3DRECT

    With rec
        .y1 = 0
        .y2 = 32
        .x1 = 0
        .x2 = 32
    End With

    
    If CharSprites(i) = 0 Then
        Exit Sub
    End If
    
    Call DX8.BeginScene
    
        Call DX8.SetTexture(Tr_Sprites(CharSprites(i)))
        Call DX8.DrawTexture(0, 0, PIC_X, PIC_Y, 3 * PIC_X, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 512, 32)
    
    Call DX8.EndSceneSp(frmMainMenu.picSelChar.hWnd, rec)
End Sub


Public Sub NpcEditorBltSprite()
Dim SpritePic As Long
Dim sRECT As D3DRECT

    SpritePic = frmNpcEditor.scrlSprite.Value

    If SpritePic < 1 Or SpritePic > NumSprites Then
        frmNpcEditor.picSprite.Cls
        Exit Sub
    End If
    
    sRECT.y1 = 0
    sRECT.y2 = PIC_Y
    sRECT.x1 = 0
    sRECT.x2 = PIC_X

    Call DX8.BeginScene
    
        Call DX8.SetTexture(Tr_Sprites(SpritePic))
        Call DX8.DrawTexture(0, 0, PIC_X, PIC_Y, 3 * PIC_X, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 512, 32)
    
    Call DX8.EndSceneSp(frmNpcEditor.picSprite.hWnd, sRECT)
End Sub

Public Sub SpellEditorBltSpell()
Dim SpellPic As Long
Dim sRECT As D3DRECT

    SpellPic = frmSpellEditor.scrlPic.Value

    If SpellPic < 1 Or SpellPic > NumSpells Then
        frmSpellEditor.picPic.Cls
        Exit Sub
    End If
    
    sRECT.y1 = 0
    sRECT.y2 = PIC_Y
    sRECT.x1 = 0
    sRECT.x2 = PIC_X

    Call DX8.BeginScene
    
        Call DX8.SetTexture(Tr_Spells(SpellPic))
        Call DX8.DrawTexture(0, 0, PIC_X, PIC_Y, frmSpellEditor.scrlFrame.Value * SIZE_X, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 512, 32)
    
    Call DX8.EndSceneSp(frmSpellEditor.picPic.hWnd, sRECT)

End Sub

Public Sub BltInventory(ItemNum As Integer)
Dim PicNum As Integer
Dim sRECT As D3DRECT

    PicNum = Item(ItemNum).Pic

    If PicNum < 1 Or PicNum > NumItems Then
        frmMainGame.picInvSelected.Cls
        Exit Sub
    End If

    
    sRECT.y1 = 0
    sRECT.y2 = PIC_Y
    sRECT.x1 = 0
    sRECT.x2 = PIC_X

    
    Call DX8.BeginScene
    
        Call DX8.SetTexture(Tr_Items(PicNum))
        Call DX8.DrawTexture(0, 0, PIC_X, PIC_Y, 0, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 32, 32)
    
    Call DX8.EndSceneSp(frmMainGame.picInvSelected.hWnd, sRECT)
    
End Sub

Public Sub MapItemEditorBltItem()
Dim ItemNum As Integer
Dim sRECT As D3DRECT

    ItemNum = Item(frmMapItem.scrlItem.Value).Pic

    If ItemNum < 1 Or ItemNum > NumItems Then
        frmMapItem.picPreview.Cls
        Exit Sub
    End If

    sRECT.y1 = 0
    sRECT.y2 = PIC_Y
    sRECT.x1 = 0
    sRECT.x2 = PIC_X

    
    Call DX8.BeginScene
    
        Call DX8.SetTexture(Tr_Items(ItemNum))
        Call DX8.DrawTexture(0, 0, PIC_X, PIC_Y, 0, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 32, 32)
    
    Call DX8.EndSceneSp(frmMapItem.picPreview.hWnd, sRECT)
    
End Sub

Public Sub KeyItemEditorBltItem()
Dim ItemNum As Integer
Dim sRECT As D3DRECT

    ItemNum = Item(frmMapKey.scrlItem.Value).Pic
    
    If ItemNum < 1 Or ItemNum > NumItems Then
        frmMapKey.picPreview.Cls
        Exit Sub
    End If

    
    sRECT.y1 = 0
    sRECT.y2 = PIC_Y
    sRECT.x1 = 0
    sRECT.x2 = PIC_X

    
    Call DX8.BeginScene
    
        Call DX8.SetTexture(Tr_Items(ItemNum))
        Call DX8.DrawTexture(0, 0, PIC_X, PIC_Y, 0, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 32, 32)
    
    Call DX8.EndSceneSp(frmMapKey.picPreview.hWnd, sRECT)
    
End Sub

Public Sub ItemEditorBltItem()
Dim ItemPic As Integer
Dim sRECT As D3DRECT

    ItemPic = frmItemEditor.scrlPic.Value
    
    If ItemPic < 1 Or ItemPic > NumItems Then
        frmItemEditor.picPic.Cls
        Exit Sub
    End If

    sRECT.y1 = 0
    sRECT.y2 = PIC_Y
    sRECT.x1 = 0
    sRECT.x2 = PIC_X
    
    Call DX8.BeginScene
    
        Call DX8.SetTexture(Tr_Items(ItemPic))
        Call DX8.DrawTexture(0, 0, PIC_X, PIC_Y, 0, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 32, 32)
    
    Call DX8.EndSceneSp(frmItemEditor.picPic.hWnd, sRECT)
    
End Sub

Public Sub BltMapEditorTilePreview()
Dim sRECT As D3DRECT

    sRECT.y1 = 0
    sRECT.y2 = PIC_Y
    sRECT.x1 = 0
    sRECT.x2 = PIC_X
    
    Call DX8.BeginScene
    
        Call DX8.SetTexture(Tr_Tiles(frmMainGame.scrlTileSet))
        Call DX8.DrawTexture(0, 0, PIC_X, PIC_Y, EditorTileX * PIC_X, EditorTileY * PIC_Y, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 256, 256)
    
    Call DX8.EndSceneSp(frmMainGame.picSelect.hWnd, sRECT)
End Sub

Public Function DrawMapAttributes()
    Dim x As Long
    Dim y As Long
   
    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY

            With map(5).Tile(x, y)

                Select Case .Type
                   
                    Case TILE_TYPE_BLOCKED
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "B", &HFFFF0000
                       
                    Case TILE_TYPE_WARP
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "W", &HFF0000FF
                   
                    Case TILE_TYPE_ITEM
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "I", &HFFFFFFFF
                   
                    Case TILE_TYPE_NPCAVOID
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "N", &HFFFFFFFF
                   
                    Case TILE_TYPE_KEY
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "K", &HFFFFFFFF
                   
                    Case TILE_TYPE_KEYOPEN
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "O", &HFFFFFFFF
                   
                    Case TILE_TYPE_HEAL
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "H", &HFFFFFFFF
                                           
                    Case TILE_TYPE_KILL
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "K", &HFFFFFFFF
                   
                    Case TILE_TYPE_DOOR
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "D", &HFFFFFFFF
                   
                    Case TILE_TYPE_SIGN
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "S", &HFFFFFFFF

                    Case TILE_TYPE_MSG
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "M", &HFFFFFFFF
                   
                    Case TILE_TYPE_SPRITE
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "SP", &HFFFFFFFF
                   
                    Case TILE_TYPE_NPCSPAWN
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "NS", &HFFFFFFFF
                   
                    Case TILE_TYPE_NUDGE
                        DX8.DrawText ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "NU", &HFFFFFFFF

                End Select

            End With

        Next
    Next
   
End Function

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim x2, y2 As Long
Dim insight As Boolean
    
                insight = False
                Select Case Player(Index).map
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
                
    If insight Then
    
    ' Check access level to determine color
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = QBColor(Brown)
            Case 1
                Color = QBColor(Gray)
            Case 2
                Color = QBColor(Green)
            Case 3
                Color = QBColor(Blue)
            Case 4
                Color = QBColor(Yellow)
            Case 5
                Color = QBColor(White)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If

    ' Determine location for text
    TextX = (x2 + GetPlayerX(Index)) * PIC_X + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerName(Index)) / 2) * 10) + StaticX
    TextY = (y2 + GetPlayerY(Index)) * PIC_Y + Player(Index).YOffset - (32) + 16 + StaticY
    
    ' Draw name
    Call DX8.DrawText(TextX, TextY, GetPlayerName(Index), DX8.AcRGB(255, Color))

    End If
End Sub

Public Sub BltItem(ByVal ItemNum As Long, ByVal MapNum As Long)
Dim PicNum As Integer
Dim x As Long
Dim y As Long
Dim x2 As Long
Dim y2 As Long
Dim insight As Boolean
Dim rec As RECT

    PicNum = Item(MapItem(ItemNum, MapNum).Num).Pic
    
    If PicNum < 1 Or PicNum > NumItems Then Exit Sub

            insight = False
                Select Case MapNum
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
If insight Then

    With rec
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
    
    x = MapItem(ItemNum, MapNum).x
    y = MapItem(ItemNum, MapNum).y
    
    Call DX8.SetTexture(Tr_Items(PicNum))
    Call DX8.DrawTexture(MapTilePosition(x, y).PosX + StaticX + x2 * PIC_X, MapTilePosition(x, y).PosY + StaticY + y2 * PIC_Y, PIC_X, PIC_Y, 0, 0, PIC_X, PIC_Y, DX8.ARGB(255, 255, 255, 255), 32, 32)
    
    End If
End Sub

Public Sub DrawPlayerGuildName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim x2, y2 As Long
Dim insight As Boolean
     
If Player(Index).Guild = "" Then Exit Sub
    
                insight = False
                Select Case Player(Index).map
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
                
    If insight Then
    
    ' Check access level to determine color
        Select Case Player(Index).GuildAccess
            Case 0
                Color = QBColor(Brown)
            Case 1
                Color = QBColor(Gray)
            Case 2
                Color = QBColor(Green)
        End Select

    ' Determine location for text
    TextX = (x2 + GetPlayerX(Index)) * PIC_X + Player(Index).XOffset + (PIC_X \ 2) - ((Len(Player(Index).Guild) / 2) * 10) + StaticX
    TextY = (y2 + GetPlayerY(Index)) * PIC_Y - 14 + Player(Index).YOffset - 32 + 16 + StaticY
    
    ' Draw name
    Call DX8.DrawText(TextX, TextY, Player(Index).Guild, DX8.AcRGB(255, Color))
    End If
End Sub
