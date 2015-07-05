Attribute VB_Name = "modDirectX"
Option Explicit

' Direct X
Public DX As New DirectX7
Public DD As DirectDraw7
Public DD_Clip As DirectDrawClipper
Public DD_PrimarySurf As DirectDrawSurface7
Public DD_SpriteSurf As DirectDrawSurface7
Public DD_TileSurf As DirectDrawSurface7
Public DD_ItemSurf As DirectDrawSurface7
Public DD_EmoticonSurf As DirectDrawSurface7
Public DD_BackBuffer As DirectDrawSurface7
Public DD_AnimationSurf As DirectDrawSurface7
Public DD_AnimationSurf2 As DirectDrawSurface7

Public DDSD_Primary As DDSURFACEDESC2
Public DDSD_Sprite As DDSURFACEDESC2
Public DDSD_Tile As DDSURFACEDESC2
Public DDSD_Item As DDSURFACEDESC2
Public DDSD_Emoticon As DDSURFACEDESC2
Public DDSD_BackBuffer As DDSURFACEDESC2
Public DDSD_Animation As DDSURFACEDESC2
Public DDSD_Animation2 As DDSURFACEDESC2

Public LookUpTileRec(MAX_INTEGER) As RECT

Sub InitDirectX()

    frmSendGetData.Visible = True
    SetStatus "Initializing DirectX..."
    
    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate(vbNullString)
    
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMainGame.hwnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMainGame.picScreen.hwnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
          
    ' Initialize all surfaces
    InitSurfaces
    
    ' Set the TILE_WIDTH
    TILE_WIDTH = (DDSD_Tile.lWidth / PIC_X) + 1
    
    ' Init look up math
    ' Has to be after the TILE_WIDTH
    InitLookUps
End Sub

Sub InitLookUps()
Dim i As Long

    For i = 0 To MAX_INTEGER
        With LookUpTileRec(i)
            .Top = (i \ TILE_WIDTH) * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = (i Mod TILE_WIDTH) * PIC_X
            .Right = .Left + PIC_X
        End With
    Next
End Sub

Sub InitSurfaces()
Dim Key As DDCOLORKEY
    
    ' Set the key for masks
    With Key
        .Low = 0
        .High = 0
    End With
    
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lWidth = (MAX_MAPX + 3) * PIC_X
        .lHeight = (MAX_MAPY + 3) * PIC_Y
    End With
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
        
    InitSurface "Sprites", DDSD_Sprite, DD_SpriteSurf
    InitSurface "Emoticons", DDSD_Emoticon, DD_EmoticonSurf
    InitSurface "Tiles", DDSD_Tile, DD_TileSurf
    InitSurface "Items", DDSD_Item, DD_ItemSurf
    InitSurface "animations1", DDSD_Animation, DD_AnimationSurf
    InitSurface "animations2", DDSD_Animation2, DD_AnimationSurf2
End Sub

Sub DestroyDirectX()
    Set DX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_TileSurf = Nothing
    Set DD_ItemSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    Set DD_AnimationSurf = Nothing
    Set DD_AnimationSurf2 = Nothing
End Sub

Public Sub InitSurface(ByVal File As String, ByRef ddsd_temp As DDSURFACEDESC2, ByRef DD_Temp As DirectDrawSurface7, Optional ByVal compressed As Boolean = False)
Dim SourceFile As String
Dim DestFile As String
On Error GoTo handler

    If compressed Then
        SourceFile = App.Path & "\core files\graphics\" & File & ".mrg"
        DestFile = App.Path & "\core files\graphics\" & File & ".bmp"
        
        If Not FileExist(DestFile, True) Then
            DecompressFile SourceFile, DestFile
        End If
        
        With ddsd_temp
            .lFlags = DDSD_CAPS
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End With
        Set DD_Temp = DD.CreateSurfaceFromFile(DestFile, ddsd_temp)
        SetMaskColorFromPixel DD_Temp, 0, 0
        
        Kill DestFile
    Else
        SourceFile = App.Path & "\core files\graphics\" & File & ".bmp"
        With ddsd_temp
            .lFlags = DDSD_CAPS
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End With
        Set DD_Temp = DD.CreateSurfaceFromFile(SourceFile, ddsd_temp)
        SetMaskColorFromPixel DD_Temp, 0, 0
    End If
    
    Exit Sub
handler:
    Call MsgBox("Error #001 - Graphics file(s) missing! " & File, vbOKOnly, GAME_NAME)
    Call GameDestroy
End Sub

Function NeedToRestoreSurfaces() As Boolean
    NeedToRestoreSurfaces = True
    If DD.TestCooperativeLevel = DD_OK Then NeedToRestoreSurfaces = False
End Function

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = X
        .Top = Y
        .Right = X
        .Bottom = Y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0
    
    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(X, Y)
        .High = .Low
    End With
    
    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    
    TheSurface.Unlock TmpR
End Sub


Public Sub BltTargets()
Dim i As Long

    For i = 1 To MapPlayersCount
        If Player(MapPlayers(i)).X = CurX Then
            If Player(MapPlayers(i)).Y = CurY Then
                ' Draw Targets over players
                BltTarget True, MapPlayers(i), TARGET_TYPE_PLAYER
                Exit For
            End If
        End If
    Next
'    For i = 1 To MAX_PLAYERS
'        If IsPlaying(i) Then
'            If Current_Map(i) = Current_Map(MyIndex) Then
'                If Player(i).X = CurX Then
'                    If Player(i).Y = CurY Then
'                        ' Draw Targets over players
'                        BltTarget True, i, TARGET_TYPE_PLAYER
'                        Exit For
'                    End If
'                End If
'            End If
'        End If
'    Next
    For i = 1 To MapNpcCount
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).X = CurX Then
                If MapNpc(i).Y = CurY Then
                    ' If it's a shopkeeper we won't draw the target but show an icon
                    If Npc(MapNpc(i).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        BltTarget True, i, TARGET_TYPE_NPC
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    If MyTarget Then BltTarget False, MyTarget, MyTargetType
End Sub

Public Sub BltTarget(ByVal MouseOver As Boolean, ByVal Target As Long, ByVal TargetType As Long)
Dim X As Long
Dim Y As Long
Dim rec As RECT

    If MouseOver Then
        With rec
            .Top = 0
            .Left = 96
            .Bottom = 32
            .Right = .Left + 64
        End With
    Else
        With rec
            .Top = 0
            .Left = 32
            .Bottom = 32
            .Right = .Left + 64
        End With
    End If

    Select Case TargetType
        Case TARGET_TYPE_PLAYER
            X = ConvertMapX(Player(Target).X) + Player(Target).XOffset - 16
            Y = ConvertMapY(Player(Target).Y) - 4 + Player(Target).YOffset

        Case TARGET_TYPE_NPC
            X = ConvertMapX(MapNpc(Target).X) + MapNpc(Target).XOffset - 16
            Y = ConvertMapY(MapNpc(Target).Y) - 4 + MapNpc(Target).YOffset
    End Select

    Render_Sprite X, Y, rec, DD_TileSurf
End Sub

Sub BltTiles()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Ground As Long
Dim Mask As Long
Dim Anim As Long
Dim Mask2 As Long
Dim M2Anim As Long

    For X = TileView.Left To TileView.Right
        For Y = TileView.Top To TileView.Bottom
            If IsValidMapPoint(X, Y) Then
                Ground = Map.Tile(X, Y).Ground
                Mask = Map.Tile(X, Y).Mask
                Anim = Map.Tile(X, Y).Anim
                Mask2 = Map.Tile(X, Y).Mask2
                M2Anim = Map.Tile(X, Y).M2Anim
    
                If Ground Then
                    DD_BackBuffer.BltFast ConvertMapX(X), ConvertMapY(Y), DD_TileSurf, LookUpTileRec(Ground), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
    
                If (Not MapAnim) Or (Anim <= 0) Then
                    If Mask Then
                        If Not TempTile(X, Y).Open Then
                            DD_BackBuffer.BltFast ConvertMapX(X), ConvertMapY(Y), DD_TileSurf, LookUpTileRec(Mask), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                        End If
                    End If
                Else
                    If Anim Then
                        DD_BackBuffer.BltFast ConvertMapX(X), ConvertMapY(Y), DD_TileSurf, LookUpTileRec(Anim), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                End If
    
                If (Not MapAnim) Or (M2Anim <= 0) Then
                    If Mask2 Then
                        DD_BackBuffer.BltFast ConvertMapX(X), ConvertMapY(Y), DD_TileSurf, LookUpTileRec(Mask2), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Else
                    If M2Anim Then
                        DD_BackBuffer.BltFast ConvertMapX(X), ConvertMapY(Y), DD_TileSurf, LookUpTileRec(M2Anim), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                End If
            End If
        Next
    Next
End Sub

Sub BltFringeTiles()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Fringe As Long
Dim FAnim As Long
Dim Fringe2 As Long
Dim F2Anim As Long

    For X = TileView.Left To TileView.Right
        For Y = TileView.Top To TileView.Bottom
            If IsValidMapPoint(X, Y) Then
                Fringe = Map.Tile(X, Y).Fringe
                FAnim = Map.Tile(X, Y).FAnim
                Fringe2 = Map.Tile(X, Y).Fringe2
                F2Anim = Map.Tile(X, Y).F2Anim
    
                If (Not MapAnim) Or (FAnim <= 0) Then
                    If Fringe Then
                        DD_BackBuffer.BltFast ConvertMapX(X), ConvertMapY(Y), DD_TileSurf, LookUpTileRec(Fringe), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Else
                    If FAnim Then
                        DD_BackBuffer.BltFast ConvertMapX(X), ConvertMapY(Y), DD_TileSurf, LookUpTileRec(FAnim), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                End If
    
                If (Not MapAnim) Or (F2Anim <= 0) Then
                    If Fringe2 Then
                        DD_BackBuffer.BltFast ConvertMapX(X), ConvertMapY(Y), DD_TileSurf, LookUpTileRec(Fringe2), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Else
                    If F2Anim Then
                        DD_BackBuffer.BltFast ConvertMapX(X), ConvertMapY(Y), DD_TileSurf, LookUpTileRec(F2Anim), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                End If
            End If
        Next
    Next
End Sub

Public Sub BltItems()
Dim i As Long
Dim rec As RECT
    
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(i).Num Then
            If InViewPort(MapItem(i).X, MapItem(i).Y) Then
                With rec
                    .Top = Item(MapItem(i).Num).Pic * PIC_Y
                    .Bottom = .Top + PIC_Y
                    .Left = 0
                    .Right = PIC_X
                End With
            
                Render_Sprite ConvertMapX(MapItem(i).X), ConvertMapY(MapItem(i).Y), rec, DD_ItemSurf
            End If
        End If
    Next
End Sub

Public Sub BltPlayers()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Anim As Byte
Dim rec As RECT

    For i = 1 To MapPlayersCount
        If InViewPort(Current_X(MapPlayers(i)), Current_Y(MapPlayers(i))) Then
            If Not Current_IsDead(MapPlayers(i)) Then
                Anim = 0
                
                ' Check to see if we want to stop making him attack
                If GetTickCount > Player(MapPlayers(i)).AttackTimer + 1000 Then
                    Player(MapPlayers(i)).Attacking = 0
                    Player(MapPlayers(i)).AttackTimer = 0
                End If
                
                If Player(MapPlayers(i)).Moving = MOVING_WALKING Then
                    Select Case Current_Dir(MapPlayers(i))
                        Case DIR_UP
                            If Player(MapPlayers(i)).YOffset < Half_PIC_Y Then Anim = 1
                            If Player(MapPlayers(i)).YOffset > Half_PIC_Y Then Anim = 2
                        Case DIR_DOWN
                            If Player(MapPlayers(i)).YOffset < -Half_PIC_Y Then Anim = 1
                            If Player(MapPlayers(i)).YOffset > -Half_PIC_Y Then Anim = 2
                        Case DIR_LEFT
                            If Player(MapPlayers(i)).XOffset < Half_PIC_Y Then Anim = 1
                            'If (Player(Index).XOffset > PIC_Y / 2) Then Anim = 2
                        Case DIR_RIGHT
                            If Player(MapPlayers(i)).XOffset < -Half_PIC_Y Then Anim = 1
                            'If (Player(Index).XOffset > PIC_Y / 2 * -1) Then Anim = 2
                    End Select
                End If
                
                If Player(MapPlayers(i)).Attacking Then
                    If GetTickCount < Player(MapPlayers(i)).AttackTimer + 500 Then Anim = 3
                End If
                
                With rec
                    .Top = Current_Sprite(MapPlayers(i)) * PIC_Y
                    .Bottom = .Top + PIC_Y
                    .Left = ((Current_Dir(MapPlayers(i)) * 4) + Anim) * PIC_X
                    .Right = .Left + PIC_X
                End With
            Else
                With rec
                    .Top = Current_Sprite(MapPlayers(i)) * PIC_Y
                    .Bottom = .Top + PIC_Y
                    .Left = 512
                    .Right = .Left + PIC_X
                End With
            End If
            
            X = ConvertMapX(Current_X(MapPlayers(i))) + Player(MapPlayers(i)).XOffset
            Y = ConvertMapY(Current_Y(MapPlayers(i))) + Player(MapPlayers(i)).YOffset - 4
        
            Render_Sprite X, Y, rec, DD_SpriteSurf
        End If
    Next
'    For i = 1 To MAX_PLAYERS
'        If IsPlaying(i) Then
'            If Current_Map(i) = Current_Map(MyIndex) Then
'                If InViewPort(Current_X(i), Current_Y(i)) Then
'                    If Not Current_IsDead(i) Then
'                        Anim = 0
'
'                        ' Check to see if we want to stop making him attack
'                        If GetTickCount > Player(i).AttackTimer + 1000 Then
'                            Player(i).Attacking = 0
'                            Player(i).AttackTimer = 0
'                        End If
'
'                        If Player(i).Moving = MOVING_WALKING Then
'                            Select Case Current_Dir(i)
'                                Case DIR_UP
'                                    If Player(i).YOffset < Half_PIC_Y Then Anim = 1
'                                    If Player(i).YOffset > Half_PIC_Y Then Anim = 2
'                                Case DIR_DOWN
'                                    If Player(i).YOffset < -Half_PIC_Y Then Anim = 1
'                                    If Player(i).YOffset > -Half_PIC_Y Then Anim = 2
'                                Case DIR_LEFT
'                                    If Player(i).XOffset < Half_PIC_Y Then Anim = 1
'                                    'If (Player(Index).XOffset > PIC_Y / 2) Then Anim = 2
'                                Case DIR_RIGHT
'                                    If Player(i).XOffset < -Half_PIC_Y Then Anim = 1
'                                    'If (Player(Index).XOffset > PIC_Y / 2 * -1) Then Anim = 2
'                            End Select
'                        End If
'
'                        If Player(i).Attacking Then
'                            If GetTickCount < Player(i).AttackTimer + 500 Then Anim = 3
'                        End If
'
'                        With rec
'                            .Top = Current_Sprite(i) * PIC_Y
'                            .Bottom = .Top + PIC_Y
'                            .Left = ((Current_Dir(i) * 4) + Anim) * PIC_X
'                            .Right = .Left + PIC_X
'                        End With
'                    Else
'                        With rec
'                            .Top = Current_Sprite(i) * PIC_Y
'                            .Bottom = .Top + PIC_Y
'                            .Left = 16 * PIC_X
'                            .Right = .Left + PIC_X
'                        End With
'                    End If
'
'                    X = ConvertMapX(Current_X(i)) + Player(i).XOffset
'                    Y = ConvertMapY(Current_Y(i)) + Player(i).YOffset - 4
'
'                    Render_Sprite X, Y, rec, DD_SpriteSurf
'                End If
'            End If
'        End If
'    Next
End Sub

Sub BltPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim Text As String
    
    ' Check if it's in range
    If Not InViewPort(Player(Index).X, Player(Index).Y) Then Exit Sub
    
    ' Check access level
    If Current_PK(Index) = NO Then
        Select Case Current_Access(Index)
            Case 0
                Color = QBColor(BrightGreen)
            Case 1
                Color = QBColor(BrightGreen)
            Case 2
                Color = QBColor(BrightGreen)
            Case 3
                Color = QBColor(BrightGreen)
            Case 4
                Color = QBColor(BrightGreen)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If
    
    If Len(Current_GuildAbbreviation(Index)) > 0 Then
        Text = Current_GuildAbbreviation(Index) & " "
        TextX = ConvertMapX(Current_X(Index)) + Player(Index).XOffset + Half_PIC_X - (getTextWidth(Text & Current_Name(Index)) / 2)
        TextY = ConvertMapY(Current_Y(Index)) + Player(Index).YOffset - Half_PIC_Y - 4

        DrawText TexthDC, TextX, TextY, Current_GuildAbbreviation(Index), QBColor(Yellow)
        
        TextX = ConvertMapX(Current_X(Index)) + Player(Index).XOffset + Half_PIC_X - (getTextWidth(Text & Current_Name(Index)) / 2) + getTextWidth(Text)
        
        DrawText TexthDC, TextX, TextY, Current_Name(Index), Color
    Else
        TextX = ConvertMapX(Current_X(Index)) + Player(Index).XOffset + Half_PIC_X - (getTextWidth(Current_Name(Index)) / 2)
        TextY = ConvertMapY(Current_Y(Index)) + Player(Index).YOffset - Half_PIC_Y - 4

        DrawText TexthDC, TextX, TextY, Current_Name(Index), Color
    End If
End Sub

Public Sub BltInventory()
Dim i As Long, X As Long, Y As Long, ItemNum As Long
Dim Amount As String
Dim rec As RECT, rec_pos As RECT

    If frmMainGame.picPouch.Visible Then
        frmMainGame.picPouch.Cls
                
        For i = 1 To MAX_INV
            ItemNum = Current_InvItemNum(MyIndex, i)
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                With rec
                    .Top = Item(ItemNum).Pic * PIC_Y
                    .Bottom = .Top + PIC_Y
                    .Left = 0
                    .Right = .Left + PIC_X
                End With
                
                With rec_pos
                    .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With
                
                DD_ItemSurf.BltToDC frmMainGame.picPouch.hdc, rec, rec_pos

                ' If item is a stack - draw the amount you have
                If Item(ItemNum).Stack Then
                    Amount = CStr(Current_InvItemValue(MyIndex, i))
                    DrawText frmMainGame.picPouch.hdc, rec_pos.Left, rec_pos.Top - 10, Amount, QBColor(White)
                End If
            End If
        Next
    
        frmMainGame.picPouch.Refresh
    End If
    
End Sub

Public Sub BltInventoryItem(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT
Dim ItemNum As Long
    
    ItemNum = Current_InvItemNum(MyIndex, DragInvSlotNum)
    
    If ItemNum <= 0 Then Exit Sub
    If ItemNum > MAX_ITEMS Then Exit Sub
        
    With rec
        .Top = Item(ItemNum).Pic * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
            
    With rec_pos
        .Top = 4
        .Bottom = .Top + PIC_Y
        .Left = 4
        .Right = .Left + PIC_X
    End With
    
    DD_ItemSurf.BltToDC frmMainGame.picTempInv.hdc, rec, rec_pos
    
    With frmMainGame.picTempInv
        .Top = Y
        .Left = X
        .Visible = True
        .ZOrder (0)
    End With
    
End Sub

Public Sub BltActionMsg(ByVal Index As Long)
Dim X As Long
Dim Y As Long
Dim i As Long
    
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            If Not InViewPort(ActionMsg(Index).X, ActionMsg(Index).Y) Then Exit Sub
            
            If ActionMsg(Index).Y > 0 Then
                'X = ConvertMapX(ActionMsg(Index).X) + Half_PIC_X - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                X = ConvertMapX(ActionMsg(Index).X) + Half_PIC_X - (getTextWidth(ActionMsg(Index).Message) / 2)
                Y = ConvertMapY(ActionMsg(Index).Y) - Half_PIC_Y - 2
            Else
                'X = ConvertMapX(ActionMsg(Index).X) + Half_PIC_X - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                X = ConvertMapX(ActionMsg(Index).X) + Half_PIC_X - (getTextWidth(ActionMsg(Index).Message) / 2)
                Y = ConvertMapY(ActionMsg(Index).Y) - Half_PIC_Y + 18
            End If
            
        Case ACTIONMSG_SCROLL
            If Not InViewPort(ActionMsg(Index).X, ActionMsg(Index).Y) Then Exit Sub
            
            If ActionMsg(Index).Y > 0 Then
                'X = ConvertMapX(ActionMsg(Index).X) + Half_PIC_X - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                X = ConvertMapX(ActionMsg(Index).X) + Half_PIC_X - (getTextWidth(ActionMsg(Index).Message) / 2)
                Y = ConvertMapY(ActionMsg(Index).Y) - Half_PIC_Y - 2 - (ActionMsg(Index).Scroll * 0.35)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                'X = ConvertMapX(ActionMsg(Index).X) + Half_PIC_X - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                X = ConvertMapX(ActionMsg(Index).X) + Half_PIC_X - (getTextWidth(ActionMsg(Index).Message) / 2)
                Y = ConvertMapY(ActionMsg(Index).Y) - Half_PIC_Y + 18 + (ActionMsg(Index).Scroll * 0.35)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If
            
        Case ACTIONMSG_SCREEN
            'X = Camera.Left + HalfX - (Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8\
            X = Camera.Left + HalfX - (getTextWidth(ActionMsg(Index).Message) / 2)
            Y = Camera.Top + 365
            
    End Select
    
    If GetTickCount < ActionMsg(Index).Created Then
        DrawText TexthDC, X, Y, ActionMsg(Index).Message, QBColor(ActionMsg(Index).Color)
    Else
        ClearActionMsg Index
    End If
End Sub

Public Sub BltAnimations(ByVal Layer As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim rec As RECT
    
    For i = 1 To MAX_BYTE
        If Anim(i).Anim Then
            If InViewPort(Anim(i).X, Anim(i).Y) Then
                If Anim(i).Layer = Layer Then
                    Select Case Anim(i).Size
                        Case 1
                            With rec
                                .Top = Animation(Anim(i).AnimNum).Animation * PIC_Y
                                .Bottom = .Top + PIC_Y
                                .Left = Anim(i).CurrFrame * PIC_X
                                .Right = .Left + PIC_X
                            End With
                            
                            X = ConvertMapX(Anim(i).X)
                            Y = ConvertMapY(Anim(i).Y) - 4
                            
                            Render_Sprite X, Y, rec, DD_AnimationSurf
                        
                        Case 2
                            With rec
                                .Top = Animation(Anim(i).AnimNum).Animation * (PIC_Y * 2)
                                .Bottom = .Top + (PIC_Y * 2)
                                .Left = Anim(i).CurrFrame * (PIC_X * 2)
                                .Right = .Left + (PIC_X * 2)
                            End With
                            
                            X = ConvertMapX(Anim(i).X) - Half_PIC_X
                            Y = ConvertMapY(Anim(i).Y) - Half_PIC_Y - 4

                            Render_Sprite X, Y, rec, DD_AnimationSurf2
                    End Select
                End If
            End If
            
            ' Even if it's not in the viewport, still have to handle it
            ' Slows down animation so we can actually see it..
            If GetTickCount > Anim(i).Created Then
            
                ' Sets the next frame of the animation
                Anim(i).CurrFrame = Anim(i).CurrFrame + 1
                
                Anim(i).Created = GetTickCount + Anim(i).Speed
                
                ' Checks to see if it's above the Max amount of frames for spells
                If Anim(i).CurrFrame > Anim(i).MaxFrames Then ClearAnim i
            End If
                
        End If
    Next
End Sub

Sub BltMapNPCName(ByVal Index As Long)
On Local Error Resume Next
Dim TextX As Long
Dim TextY As Long

    If MapNpc(Index).Num <= 0 Then Exit Sub
    
    ' Check if it's in range
    If Not InViewPort(MapNpc(Index).X, MapNpc(Index).Y) Then Exit Sub
    
    'TextX = ConvertMapX(MapNpc(Index).X) + MapNpc(Index).XOffset + Half_PIC_X - ((Len(Trim$(Npc(MapNpc(Index).Num).Name)) \ 2) * 8)
    TextX = ConvertMapX(MapNpc(Index).X) + MapNpc(Index).XOffset + Half_PIC_X - (getTextWidth(Trim$(Npc(MapNpc(Index).Num).Name)) / 2)
    TextY = ConvertMapY(MapNpc(Index).Y) + MapNpc(Index).YOffset - Half_PIC_Y - 4
    
    DrawText TexthDC, TextX, TextY, Trim$(Npc(MapNpc(Index).Num).Name), vbWhite
End Sub

Sub BltMapItemName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Text As String

    With Item(MapItem(Index).Num)

        ' If it's stack or currency draw out amount
        If .Stack Then
            Text = Trim$(.Name) & " (" & MapItem(Index).Value & ")"
        Else
            Text = Trim$(.Name)
        End If
        
        'TextX = ConvertMapX(MapItem(Index).X) + Half_PIC_X - ((Len(Text) \ 2) * 8)
        TextX = ConvertMapX(MapItem(Index).X) + Half_PIC_X - (getTextWidth(Text) / 2)
        TextY = ConvertMapY(MapItem(Index).Y) - Half_PIC_Y - 4
        
        DrawText TexthDC, TextX, TextY, Text, QBColor(BrightCyan)
    End With
End Sub

Public Sub BltNpcs()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Anim As Byte
Dim rec As RECT

    For i = 1 To MapNpcCount
        If MapNpc(i).Num Then
            If InViewPort(MapNpc(i).X, MapNpc(i).Y) Then
                Anim = 0
                
                 ' Check to see if we want to stop making him attack
                If GetTickCount > MapNpc(i).AttackTimer + 1000 Then
                    MapNpc(i).Attacking = 0
                    MapNpc(i).AttackTimer = 0
                End If
               
                If MapNpc(i).Attacking = 0 Then
                    If MapNpc(i).Moving = MOVING_WALKING Then
                        Select Case MapNpc(i).Dir
                            Case DIR_UP
                                If MapNpc(i).YOffset < Half_PIC_Y Then Anim = 1
                                If MapNpc(i).YOffset > Half_PIC_Y Then Anim = 2
                            Case DIR_DOWN
                                If MapNpc(i).YOffset < -Half_PIC_Y Then Anim = 1
                                If MapNpc(i).YOffset > -Half_PIC_Y Then Anim = 2
                            Case DIR_LEFT
                                If MapNpc(i).XOffset < Half_PIC_Y Then Anim = 1
                                'If (Player(Index).XOffset > PIC_Y / 2) Then Anim = 2
                            Case DIR_RIGHT
                                If MapNpc(i).XOffset < -Half_PIC_Y Then Anim = 1
                                'If (Player(Index).XOffset > PIC_Y / 2 * -1) Then Anim = 2
                        End Select
                    End If
                Else
                    If GetTickCount < MapNpc(i).AttackTimer + 500 Then
                        Anim = 3
                    End If
                End If
                
                With rec
                    .Top = Npc(MapNpc(i).Num).Sprite * PIC_Y
                    .Bottom = .Top + PIC_Y
                    .Left = ((MapNpc(i).Dir * 4) + Anim) * PIC_X
                    .Right = .Left + PIC_X
                End With
                
                X = ConvertMapX(MapNpc(i).X) + MapNpc(i).XOffset
                Y = ConvertMapY(MapNpc(i).Y) + MapNpc(i).YOffset - 4
                
                Render_Sprite X, Y, rec, DD_SpriteSurf
            End If
        End If
    Next
End Sub

Sub BltNpcIcons()
Dim i As Long
Dim n As Long, Q As Long
Dim X As Long, Y As Long
Dim rec As RECT
Dim NpcNum As Long
Dim QuestNum As Long

    For i = 1 To MapNpcCount
        NpcNum = MapNpc(i).Num
        If NpcNum Then
            If InViewPort(MapNpc(i).X, MapNpc(i).Y) Then
                If MapNpc(i).X = CurX Then
                    If MapNpc(i).Y = CurY Then
                        Select Case Npc(NpcNum).Behavior
                            Case NPC_BEHAVIOR_SHOPKEEPER
                                ' Draw icon if inrange
                                If NpcInRange(i, 2) Then
                                    With rec
                                        .Top = 8 * PIC_Y
                                        .Bottom = .Top + PIC_Y
                                        .Left = 0
                                        .Right = PIC_X
                                    End With
                                Else
                                    With rec
                                        .Top = 7 * PIC_Y
                                        .Bottom = .Top + PIC_Y
                                        .Left = 0
                                        .Right = PIC_X
                                    End With
                                End If
                                
                                X = ConvertMapX(MapNpc(i).X) + MapNpc(i).XOffset
                                Y = ConvertMapY(MapNpc(i).Y) + MapNpc(i).YOffset - 50
                                
                                Render_Sprite X, Y, rec, DD_EmoticonSurf
                                
                            Case NPC_BEHAVIOR_QUEST
                                ' TODO: Optimize the shit out of this
                                'For n = 1 To MAX_QUESTS
                                For n = 1 To NpcQuests(NpcNum).QuestCount
                                    QuestNum = NpcQuests(NpcNum).QuestList(n)
                                    If LenB(Quest(QuestNum).Name) > 0 Then
                                        ' Check if a starter NPC
                                        If Quest(QuestNum).StartNPC = NpcNum Then
                                            ' Now we check if the status of the quest
                                            Select Case Player(MyIndex).CompletedQuests(QuestNum)
                                                Case QUEST_STATUS_COMPLETE
                                                    ' DRAW INCOMPLETE STATUS ICON ABOVE NPC
                                                    With rec
                                                        .Top = 6 * PIC_Y
                                                        .Bottom = .Top + PIC_Y
                                                        .Left = 0
                                                        .Right = PIC_X
                                                    End With
                                                    
                                                    X = ConvertMapX(MapNpc(i).X) + MapNpc(i).XOffset
                                                    Y = ConvertMapY(MapNpc(i).Y) + MapNpc(i).YOffset - 50
                                                    
                                                    Render_Sprite X, Y, rec, DD_EmoticonSurf
                                                    
                                                Case QUEST_STATUS_INCOMPLETE
                                                    ' Since this is incomplete, we now check if it can be accepted
                                                    If CanAcceptQuest(MyIndex, QuestNum) Then
                                                        ' DRAW NEW QUEST ICON ABOVE NPC
                                                        With rec
                                                            .Top = 5 * PIC_Y
                                                            .Bottom = .Top + PIC_Y
                                                            .Left = 0
                                                            .Right = PIC_X
                                                        End With
                                                        
                                                        X = ConvertMapX(MapNpc(i).X) + MapNpc(i).XOffset
                                                        Y = ConvertMapY(MapNpc(i).Y) + MapNpc(i).YOffset - 50
                                                        
                                                        Render_Sprite X, Y, rec, DD_EmoticonSurf
                                                    End If
                                                    
                                            End Select
                                            Exit For
                                        ElseIf Quest(QuestNum).EndNPC = NpcNum Then
                                            ' TODO : Finish
                                        End If
                                    End If
                                Next
                                
                        End Select
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub BltEmoticons()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim rec As RECT
    
    For i = 1 To MapPlayersCount
        If Player(MapPlayers(i)).EmoticonNum >= 0 Then
            If InViewPort(Current_X(MapPlayers(i)), Current_Y(MapPlayers(i))) Then
                With rec
                    .Top = Player(MapPlayers(i)).EmoticonNum * PIC_Y
                    .Bottom = .Top + PIC_Y
                    .Left = 0
                    .Right = .Left + PIC_X
                End With
        
                X = ConvertMapX(Current_X(MapPlayers(i))) + Player(MapPlayers(i)).XOffset
                Y = ConvertMapY(Current_Y(MapPlayers(i))) + Player(MapPlayers(i)).YOffset - 50
                
                Render_Sprite X, Y, rec, DD_EmoticonSurf
                
                If GetTickCount > Player(MapPlayers(i)).EmoticonTime Then
                    Player(MapPlayers(i)).EmoticonNum = -1
                End If
            End If
        End If
    Next
End Sub

Public Sub BltTrade(ByVal ShopNum As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Amount As String
Dim rec As RECT, rec_pos As RECT

    frmTrade.picTrade.Cls
    
    For i = 1 To MAX_TRADES
        If Shop(ShopNum).TradeItem(i).GetItem > 0 Then
            With rec
                .Top = Item(Shop(ShopNum).TradeItem(i).GetItem).Pic * PIC_Y
                .Bottom = .Top + PIC_Y
                .Left = 0
                .Right = .Left + PIC_X
            End With
            
            With rec_pos
                .Top = 150 + ((InvOffsetY + 32) * ((i - 1) \ 4))
                .Bottom = .Top + PIC_Y
                .Left = 31 + ((InvOffsetX + 32) * (((i - 1) Mod 4)))
                .Right = .Left + PIC_X
            End With
            
            DD_ItemSurf.BltToDC frmTrade.picTrade.hdc, rec, rec_pos
            
            If Shop(ShopNum).TradeItem(i).GetValue > 0 Then
                Y = 150 + ((InvOffsetY + 32) * ((i - 1) \ 4)) - 10
                X = 31 + ((InvOffsetX + 32) * (((i - 1) Mod 4)))
                
                Amount = Shop(ShopNum).TradeItem(i).GetValue
                Call DrawText(frmTrade.picTrade.hdc, X, Y, Amount, QBColor(White))
            End If
        End If
    Next

    frmTrade.picTrade.Refresh
End Sub

Public Sub BltCastBar()
Dim rec As RECT, rec2 As RECT
Dim SpellCastTime As Long

    If CastingSpell Then
        SpellCastTime = Spell(Player(MyIndex).Spell(CastingSpell).SpellNum).CastTime * 1000
        
        If SpellCastTime > 0 Then
            With rec2
                .Top = Camera.Top + 325
                .Bottom = .Top + 18
                .Left = Camera.Left + 177
                .Right = .Left + 158
            End With
            DD_BackBuffer.BltColorFill rec2, RGB(255, 255, 255)
            
            With rec
                .Top = Camera.Top + 325
                .Bottom = .Top + 18
                .Left = Camera.Left + 177
                .Right = .Left + (((SpellCastTime - (CastTime - GetTickCount)) / SpellCastTime) * 158)
            End With
            DD_BackBuffer.BltColorFill rec, RGB(0, 0, 255)
            
            DD_BackBuffer.DrawBox rec2.Left, rec2.Top, rec2.Right, rec2.Bottom
        End If
        
        ' This is where we check if the player is done casting
        If GetTickCount >= CastTime Then
            CastingSpell = 0
            CastTime = 0
        End If
    End If
End Sub

Public Sub Render_Sprite(ByRef X As Long, ByRef Y As Long, ByRef SrcRec As RECT, ByRef DD_Temp As DirectDrawSurface7)
    
    If Y < Camera.Top Then
        SrcRec.Top = SrcRec.Top + (Camera.Top - Y)
        Y = Camera.Top
    End If
    
    If Y > Camera.Bottom Then
        SrcRec.Bottom = SrcRec.Bottom - (Camera.Bottom - Y)
        Y = Camera.Bottom
    End If
    
    If X < Camera.Left Then
        SrcRec.Left = SrcRec.Left + (Camera.Left - X)
        X = Camera.Left
    End If
    
    If X > Camera.Right Then
        SrcRec.Right = SrcRec.Right - (Camera.Right - X)
        X = Camera.Right
    End If
    
    DD_BackBuffer.BltFast X, Y, DD_Temp, SrcRec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub

Public Sub UpdateCamera()
Dim OffsetX As Long
Dim OffsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    OffsetX = Player(MyIndex).XOffset + PIC_X
    OffsetY = Player(MyIndex).YOffset + PIC_Y

    StartX = Current_X(MyIndex) - StartXValue
    StartY = Current_Y(MyIndex) - StartYValue
    If StartX < 0 Then
        OffsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).XOffset > 0 Then
                OffsetX = Player(MyIndex).XOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        OffsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).YOffset > 0 Then
                OffsetY = Player(MyIndex).YOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        OffsetX = 32
        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).XOffset < 0 Then
                OffsetX = Player(MyIndex).XOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        OffsetY = 32
        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).YOffset < 0 Then
                OffsetY = Player(MyIndex).YOffset + PIC_Y
            End If
        End If
        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = OffsetY
        .Bottom = .Top + ScreenY
        .Left = OffsetX
        .Right = .Left + ScreenX
    End With
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ConvertMapX = (X - TileView.Left) * PIC_X
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ConvertMapY = (Y - TileView.Top) * PIC_Y
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    InViewPort = False
    If X < TileView.Left Then Exit Function
    If Y < TileView.Top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    IsValidMapPoint = False
    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
End Function

