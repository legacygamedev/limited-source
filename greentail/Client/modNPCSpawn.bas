Attribute VB_Name = "modNPCSpawn"
Sub BltAttributeNPCName(ByVal index As Long, ByVal x As Long, ByVal y As Long)
Dim TextX As Long
Dim TextY As Long

If index > Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then Exit Sub
If MapAttributeNpc(index, x, y).Num <= 0 Then Exit Sub

If Npc(MapAttributeNpc(index, x, y).Num).Big = 0 Then
    With Npc(MapAttributeNpc(index, x, y).Num)
    'Draw name
        TextX = MapAttributeNpc(index, x, y).x * PIC_X + sx + MapAttributeNpc(index, x, y).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapAttributeNpc(index, x, y).y * PIC_Y + sx + MapAttributeNpc(index, x, y).YOffset - CLng(PIC_Y / 2) - 4
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
Else
    With Npc(MapAttributeNpc(index, x, y).Num)
    'Draw name
        TextX = MapAttributeNpc(index, x, y).x * PIC_X + sx + MapAttributeNpc(index, x, y).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapAttributeNpc(index, x, y).y * PIC_Y + sx + MapAttributeNpc(index, x, y).YOffset - CLng(PIC_Y / 2) - 32
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
End If
End Sub

Sub BltAttributeNpc(ByVal index As Long, ByVal x As Long, ByVal y As Long)
Dim Anim As Byte
Dim BX As Long, BY As Long

    If index > Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then Exit Sub
    If MapAttributeNpc(index, x, y).Num <= 0 Then Exit Sub

    ' Make sure that theres an npc there, and if not exit the sub
    If MapAttributeNpc(index, x, y).Num <= 0 Then
        Exit Sub
    End If
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapAttributeNpc(index, x, y).y * PIC_Y + MapAttributeNpc(index, x, y).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapAttributeNpc(index, x, y).x * PIC_X + MapAttributeNpc(index, x, y).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If MapAttributeNpc(index, x, y).Attacking = 0 Then
        Select Case MapAttributeNpc(index, x, y).Dir
            Case DIR_UP
                If (MapAttributeNpc(index, x, y).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapAttributeNpc(index, x, y).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapAttributeNpc(index, x, y).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapAttributeNpc(index, x, y).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapAttributeNpc(index, x, y).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If MapAttributeNpc(index, x, y).AttackTimer + 1000 < GetTickCount Then
        MapAttributeNpc(index, x, y).Attacking = 0
        MapAttributeNpc(index, x, y).AttackTimer = 0
    End If
    If Npc(MapAttributeNpc(index, x, y).Num).Big = 0 Then
        
        rec.Top = Npc(MapAttributeNpc(index, x, y).Num).Sprite * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapAttributeNpc(index, x, y).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        BX = MapAttributeNpc(index, x, y).x * PIC_X + sx + MapAttributeNpc(index, x, y).XOffset
        BY = MapAttributeNpc(index, x, y).y * PIC_Y + sx + MapAttributeNpc(index, x, y).YOffset
        
        ' Check if its out of bounds because of the offset
        If BY < 0 Then
            BY = 0
            rec.Top = rec.Top + (BY * -1)
        End If
            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(BX - (NewPlayerX * PIC_X) - NewXOffset, BY - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        rec.Top = Npc(MapAttributeNpc(index, x, y).Num).Sprite * 64 + 32
        rec.Bottom = rec.Top + 32
        rec.Left = (MapAttributeNpc(index, x, y).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64
    
        BX = MapAttributeNpc(index, x, y).x * 32 + sx - 16 + MapAttributeNpc(index, x, y).XOffset
        BY = MapAttributeNpc(index, x, y).y * 32 + sx + MapAttributeNpc(index, x, y).YOffset
   
        If BY < 0 Then
            rec.Top = Npc(MapAttributeNpc(index, x, y).Num).Sprite * 64 + 32
            rec.Bottom = rec.Top + 32
            BY = MapAttributeNpc(index, x, y).YOffset + sx
        End If
        
        If BX < 0 Then
            rec.Left = (MapAttributeNpc(index, x, y).Dir * 3 + Anim) * 64 + 16
            rec.Right = rec.Left + 48
            BX = MapAttributeNpc(index, x, y).XOffset + sx
        End If
        
        If BX > MAX_MAPX * 32 Then
            rec.Left = (MapAttributeNpc(index, x, y).Dir * 3 + Anim) * 64
            rec.Right = rec.Left + 48
            BX = MAX_MAPX * 32 + sx - 16 + MapAttributeNpc(index, x, y).XOffset
        End If

        Call DD_BackBuffer.BltFast(BX - (NewPlayerX * PIC_X) - NewXOffset, BY - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltAttributeNpcTop(ByVal index As Long, ByVal x As Long, ByVal y As Long)
Dim Anim As Byte

    If index > Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then Exit Sub
    If MapAttributeNpc(index, x, y).Num <= 0 Then Exit Sub
    
    ' Make sure that theres an npc there, and if not exit the sub
    If MapAttributeNpc(index, x, y).Num <= 0 Then
        Exit Sub
    End If
    
    If Npc(MapAttributeNpc(index, x, y).Num).Big = 0 Then Exit Sub
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapAttributeNpc(index, x, y).y * PIC_Y + MapAttributeNpc(index, x, y).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapAttributeNpc(index, x, y).x * PIC_X + MapAttributeNpc(index, x, y).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If MapAttributeNpc(index, x, y).Attacking = 0 Then
        Select Case MapAttributeNpc(index, x, y).Dir
            Case DIR_UP
                If (MapAttributeNpc(index, x, y).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapAttributeNpc(index, x, y).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapAttributeNpc(index, x, y).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapAttributeNpc(index, x, y).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapAttributeNpc(index, x, y).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If MapAttributeNpc(index, x, y).AttackTimer + 1000 < GetTickCount Then
        MapAttributeNpc(index, x, y).Attacking = 0
        MapAttributeNpc(index, x, y).AttackTimer = 0
    End If
    
    rec.Top = Npc(MapAttributeNpc(index, x, y).Num).Sprite * PIC_Y
        
     rec.Top = Npc(MapAttributeNpc(index, x, y).Num).Sprite * 64
     rec.Bottom = rec.Top + 32
     rec.Left = (MapAttributeNpc(index, x, y).Dir * 3 + Anim) * 64
     rec.Right = rec.Left + 64
 
     x = MapAttributeNpc(index, x, y).x * 32 + sx - 16 + MapAttributeNpc(index, x, y).XOffset
     y = MapAttributeNpc(index, x, y).y * 32 + sx - 32 + MapAttributeNpc(index, x, y).YOffset

     If y < 0 Then
         rec.Top = Npc(MapAttributeNpc(index, x, y).Num).Sprite * 64 + 32
         rec.Bottom = rec.Top
         y = MapAttributeNpc(index, x, y).YOffset + sx
     End If
     
     If x < 0 Then
         rec.Left = (MapAttributeNpc(index, x, y).Dir * 3 + Anim) * 64 + 16
         rec.Right = rec.Left + 48
         x = MapAttributeNpc(index, x, y).XOffset + sx
     End If
     
     If x > MAX_MAPX * 32 Then
         rec.Left = (MapAttributeNpc(index, x, y).Dir * 3 + Anim) * 64
         rec.Right = rec.Left + 48
         x = MAX_MAPX * 32 + sx - 16 + MapAttributeNpc(index, x, y).XOffset
     End If

     Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub ProcessAttributeNpcMovement(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    ' Check if npc is walking, and if so process moving them over
    If MapAttributeNpc(index, x, y).Moving = MOVING_WALKING Then
        Select Case MapAttributeNpc(index, x, y).Dir
            Case DIR_UP
                MapAttributeNpc(index, x, y).YOffset = MapAttributeNpc(index, x, y).YOffset - WALK_SPEED
            Case DIR_DOWN
                MapAttributeNpc(index, x, y).YOffset = MapAttributeNpc(index, x, y).YOffset + WALK_SPEED
            Case DIR_LEFT
                MapAttributeNpc(index, x, y).XOffset = MapAttributeNpc(index, x, y).XOffset - WALK_SPEED
            Case DIR_RIGHT
                MapAttributeNpc(index, x, y).XOffset = MapAttributeNpc(index, x, y).XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (MapAttributeNpc(index, x, y).XOffset = 0) And (MapAttributeNpc(index, x, y).YOffset = 0) Then
            MapAttributeNpc(index, x, y).Moving = 0
        End If
    End If
End Sub

Sub BltAttributeNpcBars(ByVal index As Long, ByVal x As Long, ByVal y As Long)
Dim BX As Long, BY As Long
    If MapAttributeNpc(index, x, y).HP <= 0 Then Exit Sub
    If MapAttributeNpc(index, x, y).Num < 1 Then Exit Sub

    If Npc(MapAttributeNpc(index, x, y).Num).Big = 1 Then
        BX = (MapAttributeNpc(index, x, y).x * PIC_X + sx - 9 + MapAttributeNpc(index, x, y).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        BY = (MapAttributeNpc(index, x, y).y * PIC_Y + sx + MapAttributeNpc(index, x, y).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(BX, BY + 32, BX + 50, BY + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(BX, BY + 32, BX + ((MapAttributeNpc(index, x, y).HP / 100) / (MapAttributeNpc(index, x, y).MaxHp / 100) * 50), BY + 36)
    Else
        BX = (MapAttributeNpc(index, x, y).x * PIC_X + sx + MapAttributeNpc(index, x, y).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        BY = (MapAttributeNpc(index, x, y).y * PIC_Y + sx + MapAttributeNpc(index, x, y).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(BX, BY + 32, BX + 32, BY + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
                If MapAttributeNpc(index, x, y).MaxHp < 1 Then
        Call DD_BackBuffer.DrawBox(BX, BY + 32, x + ((MapAttributeNpc(index, x, y).HP / 100) / ((MapAttributeNpc(index, x, y).MaxHp + 1) / 100) * 32), BY + 36)
        Else
        Call DD_BackBuffer.DrawBox(BX, BY + 32, BX + ((MapAttributeNpc(index, x, y).HP / 100) / (MapAttributeNpc(index, x, y).MaxHp / 100) * 32), BY + 36)
        End If
    End If
End Sub

Function CanAttributeNPCMove(ByVal Dir As Long) As Boolean
Dim x As Long, y As Long, index As Long

    CanAttributeNPCMove = True

    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                For i = 1 To MAX_ATTRIBUTE_NPCS
                    If i <= Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then
                        Select Case Dir
                            Case DIR_UP
                                If (MapAttributeNpc(i, x, y).x = GetPlayerX(MyIndex)) And (MapAttributeNpc(i, x, y).y = GetPlayerY(MyIndex) - 1) Then
                                    CanAttributeNPCMove = False
                                End If
                            Case DIR_DOWN
                                If (MapAttributeNpc(i, x, y).x = GetPlayerX(MyIndex)) And (MapAttributeNpc(i, x, y).y = GetPlayerY(MyIndex) + 1) Then
                                    CanAttributeNPCMove = False
                                End If
                            Case DIR_LEFT
                                If (MapAttributeNpc(i, x, y).x = GetPlayerX(MyIndex) - 1) And (MapAttributeNpc(i, x, y).y = GetPlayerY(MyIndex)) Then
                                    CanAttributeNPCMove = False
                                End If
                            Case DIR_RIGHT
                                If (MapAttributeNpc(i, x, y).x = GetPlayerX(MyIndex) + 1) And (MapAttributeNpc(i, x, y).y = GetPlayerY(MyIndex)) Then
                                    CanAttributeNPCMove = False
                                End If
                        End Select
                    End If
                Next i
            End If
        Next y
    Next x
End Function

