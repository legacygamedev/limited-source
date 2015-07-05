Attribute VB_Name = "modNPCSpawn"
Option Explicit

Sub BltAttributeNpc(ByVal Index As Long, ByVal x As Long, ByVal y As Long)

  Dim Anim As Byte
  Dim BX As Long
  Dim BY As Long

    If Index > Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then Exit Sub
    If MapAttributeNpc(Index, x, y).num <= 0 Then Exit Sub

    ' Make sure that theres an npc there, and if not exit the sub

    If MapAttributeNpc(Index, x, y).num <= 0 Then
        Exit Sub
    End If

    ' Only used if ever want to switch to blt rather then bltfast

    With rec_pos
        .Top = MapAttributeNpc(Index, x, y).y * PIC_Y + MapAttributeNpc(Index, x, y).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapAttributeNpc(Index, x, y).x * PIC_X + MapAttributeNpc(Index, x, y).XOffset
        .right = .Left + PIC_X
    End With

    ' Check for animation
    Anim = 0

    If MapAttributeNpc(Index, x, y).Attacking = 0 Then

        Select Case MapAttributeNpc(Index, x, y).Dir
         Case DIR_UP
            If (MapAttributeNpc(Index, x, y).YOffset < PIC_Y / 2) Then Anim = 1

         Case DIR_DOWN
            If (MapAttributeNpc(Index, x, y).YOffset < PIC_Y / 2 * -1) Then Anim = 1

         Case DIR_LEFT
            If (MapAttributeNpc(Index, x, y).XOffset < PIC_Y / 2) Then Anim = 1

         Case DIR_RIGHT
            If (MapAttributeNpc(Index, x, y).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select

     Else

        If MapAttributeNpc(Index, x, y).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If

    End If

    ' Check to see if we want to stop making him attack

    If MapAttributeNpc(Index, x, y).AttackTimer + 1000 < GetTickCount Then
        MapAttributeNpc(Index, x, y).Attacking = 0
        MapAttributeNpc(Index, x, y).AttackTimer = 0
    End If

    If Npc(MapAttributeNpc(Index, x, y).num).Big = 0 Then

        rec.Top = Npc(MapAttributeNpc(Index, x, y).num).Sprite * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapAttributeNpc(Index, x, y).Dir * 3 + Anim) * PIC_X
        rec.right = rec.Left + PIC_X

        BX = MapAttributeNpc(Index, x, y).x * PIC_X + sx + MapAttributeNpc(Index, x, y).XOffset
        BY = MapAttributeNpc(Index, x, y).y * PIC_Y + sx + MapAttributeNpc(Index, x, y).YOffset

        ' Check if its out of bounds because of the offset

        If BY < 0 Then
            BY = 0
            rec.Top = rec.Top + (BY * -1)
        End If

        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(BX - (NewPlayerX * PIC_X) - NewXOffset, BY - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
     Else
        rec.Top = Npc(MapAttributeNpc(Index, x, y).num).Sprite * 64 + 32
        rec.Bottom = rec.Top + 32
        rec.Left = (MapAttributeNpc(Index, x, y).Dir * 3 + Anim) * 64
        rec.right = rec.Left + 64

        BX = MapAttributeNpc(Index, x, y).x * 32 + sx - 16 + MapAttributeNpc(Index, x, y).XOffset
        BY = MapAttributeNpc(Index, x, y).y * 32 + sx + MapAttributeNpc(Index, x, y).YOffset

        If BY < 0 Then
            rec.Top = Npc(MapAttributeNpc(Index, x, y).num).Sprite * 64 + 32
            rec.Bottom = rec.Top + 32
            BY = MapAttributeNpc(Index, x, y).YOffset + sx
        End If

        If BX < 0 Then
            rec.Left = (MapAttributeNpc(Index, x, y).Dir * 3 + Anim) * 64 + 16
            rec.right = rec.Left + 48
            BX = MapAttributeNpc(Index, x, y).XOffset + sx
        End If

        If BX > MAX_MAPX * 32 Then
            rec.Left = (MapAttributeNpc(Index, x, y).Dir * 3 + Anim) * 64
            rec.right = rec.Left + 48
            BX = MAX_MAPX * 32 + sx - 16 + MapAttributeNpc(Index, x, y).XOffset
        End If

        Call DD_BackBuffer.BltFast(BX - (NewPlayerX * PIC_X) - NewXOffset, BY - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

End Sub

Sub BltAttributeNpcBars(ByVal Index As Long, ByVal x As Long, ByVal y As Long)

  Dim BX As Long
  Dim BY As Long

    If MapAttributeNpc(Index, x, y).HP <= 0 Then Exit Sub
    If MapAttributeNpc(Index, x, y).num < 1 Then Exit Sub

    If Npc(MapAttributeNpc(Index, x, y).num).Big = 1 Then
        BX = (MapAttributeNpc(Index, x, y).x * PIC_X + sx - 9 + MapAttributeNpc(Index, x, y).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        BY = (MapAttributeNpc(Index, x, y).y * PIC_Y + sx + MapAttributeNpc(Index, x, y).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset

        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(BX, BY + 32, BX + 50, BY + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(BX, BY + 32, BX + ((MapAttributeNpc(Index, x, y).HP / 100) / (MapAttributeNpc(Index, x, y).MaxHp / 100) * 50), BY + 36)
     Else
        BX = (MapAttributeNpc(Index, x, y).x * PIC_X + sx + MapAttributeNpc(Index, x, y).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        BY = (MapAttributeNpc(Index, x, y).y * PIC_Y + sx + MapAttributeNpc(Index, x, y).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset

        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(BX, BY + 32, BX + 32, BY + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(BX, BY + 32, BX + ((MapAttributeNpc(Index, x, y).HP / 100) / (MapAttributeNpc(Index, x, y).MaxHp / 100) * 32), BY + 36)
    End If

End Sub

Sub BltAttributeNPCName(ByVal Index As Long, ByVal x As Long, ByVal y As Long)

  Dim TextX As Long
  Dim TextY As Long

    If Index > Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then Exit Sub
    If MapAttributeNpc(Index, x, y).num <= 0 Then Exit Sub

    If Npc(MapAttributeNpc(Index, x, y).num).Big = 0 Then

        With Npc(MapAttributeNpc(Index, x, y).num)
            'Draw name
            TextX = MapAttributeNpc(Index, x, y).x * PIC_X + sx + MapAttributeNpc(Index, x, y).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
            TextY = MapAttributeNpc(Index, x, y).y * PIC_Y + sx + MapAttributeNpc(Index, x, y).YOffset - CLng(PIC_Y / 2) - 4
            DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
        End With

     Else

        With Npc(MapAttributeNpc(Index, x, y).num)
            'Draw name
            TextX = MapAttributeNpc(Index, x, y).x * PIC_X + sx + MapAttributeNpc(Index, x, y).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
            TextY = MapAttributeNpc(Index, x, y).y * PIC_Y + sx + MapAttributeNpc(Index, x, y).YOffset - CLng(PIC_Y / 2) - 32
            DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
        End With

    End If

End Sub

Sub BltAttributeNpcTop(ByVal Index As Long, ByVal x As Long, ByVal y As Long)

  Dim Anim As Byte

    If Index > Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then Exit Sub
    If MapAttributeNpc(Index, x, y).num <= 0 Then Exit Sub

    ' Make sure that theres an npc there, and if not exit the sub

    If MapAttributeNpc(Index, x, y).num <= 0 Then
        Exit Sub
    End If

    If Npc(MapAttributeNpc(Index, x, y).num).Big = 0 Then Exit Sub

    ' Only used if ever want to switch to blt rather then bltfast

    With rec_pos
        .Top = MapAttributeNpc(Index, x, y).y * PIC_Y + MapAttributeNpc(Index, x, y).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapAttributeNpc(Index, x, y).x * PIC_X + MapAttributeNpc(Index, x, y).XOffset
        .right = .Left + PIC_X
    End With

    ' Check for animation
    Anim = 0

    If MapAttributeNpc(Index, x, y).Attacking = 0 Then

        Select Case MapAttributeNpc(Index, x, y).Dir
         Case DIR_UP
            If (MapAttributeNpc(Index, x, y).YOffset < PIC_Y / 2) Then Anim = 1

         Case DIR_DOWN
            If (MapAttributeNpc(Index, x, y).YOffset < PIC_Y / 2 * -1) Then Anim = 1

         Case DIR_LEFT
            If (MapAttributeNpc(Index, x, y).XOffset < PIC_Y / 2) Then Anim = 1

         Case DIR_RIGHT
            If (MapAttributeNpc(Index, x, y).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select

     Else

        If MapAttributeNpc(Index, x, y).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If

    End If

    ' Check to see if we want to stop making him attack

    If MapAttributeNpc(Index, x, y).AttackTimer + 1000 < GetTickCount Then
        MapAttributeNpc(Index, x, y).Attacking = 0
        MapAttributeNpc(Index, x, y).AttackTimer = 0
    End If

    rec.Top = Npc(MapAttributeNpc(Index, x, y).num).Sprite * PIC_Y

    rec.Top = Npc(MapAttributeNpc(Index, x, y).num).Sprite * 64
    rec.Bottom = rec.Top + 32
    rec.Left = (MapAttributeNpc(Index, x, y).Dir * 3 + Anim) * 64
    rec.right = rec.Left + 64

    x = MapAttributeNpc(Index, x, y).x * 32 + sx - 16 + MapAttributeNpc(Index, x, y).XOffset
    y = MapAttributeNpc(Index, x, y).y * 32 + sx - 32 + MapAttributeNpc(Index, x, y).YOffset

    If y < 0 Then
        rec.Top = Npc(MapAttributeNpc(Index, x, y).num).Sprite * 64 + 32
        rec.Bottom = rec.Top
        y = MapAttributeNpc(Index, x, y).YOffset + sx
    End If

    If x < 0 Then
        rec.Left = (MapAttributeNpc(Index, x, y).Dir * 3 + Anim) * 64 + 16
        rec.right = rec.Left + 48
        x = MapAttributeNpc(Index, x, y).XOffset + sx
    End If

    If x > MAX_MAPX * 32 Then
        rec.Left = (MapAttributeNpc(Index, x, y).Dir * 3 + Anim) * 64
        rec.right = rec.Left + 48
        x = MAX_MAPX * 32 + sx - 16 + MapAttributeNpc(Index, x, y).XOffset
    End If

    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub

Function CanAttributeNPCMove(ByVal Dir As Long) As Boolean
  Dim i As Long
  Dim x As Long
  Dim y As Long

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

Sub ProcessAttributeNpcMovement(ByVal Index As Long, ByVal x As Long, ByVal y As Long)

    ' Check if npc is walking, and if so process moving them over

    If MapAttributeNpc(Index, x, y).Moving = MOVING_WALKING Then

        Select Case MapAttributeNpc(Index, x, y).Dir
         Case DIR_UP
            MapAttributeNpc(Index, x, y).YOffset = MapAttributeNpc(Index, x, y).YOffset - WALK_SPEED

         Case DIR_DOWN
            MapAttributeNpc(Index, x, y).YOffset = MapAttributeNpc(Index, x, y).YOffset + WALK_SPEED

         Case DIR_LEFT
            MapAttributeNpc(Index, x, y).XOffset = MapAttributeNpc(Index, x, y).XOffset - WALK_SPEED

         Case DIR_RIGHT
            MapAttributeNpc(Index, x, y).XOffset = MapAttributeNpc(Index, x, y).XOffset + WALK_SPEED
        End Select

        ' Check if completed walking over to the next tile

        If (MapAttributeNpc(Index, x, y).XOffset = 0) And (MapAttributeNpc(Index, x, y).YOffset = 0) Then
            MapAttributeNpc(Index, x, y).Moving = 0
        End If

    End If

End Sub

