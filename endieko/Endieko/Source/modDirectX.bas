Attribute VB_Name = "modDirectX"
Option Explicit

Public DX As New DirectX7
Public DD As DirectDraw7
Public DD_PrimarySurf As DirectDrawSurface7
Public DD_SpriteSurf As DirectDrawSurface7
Public DD_MonsterSurf As DirectDrawSurface7
Public DD_TileSurf As DirectDrawSurface7
Public DD_ItemSurf As DirectDrawSurface7
Public DD_EmoticonSurf As DirectDrawSurface7
Public DD_BackBuffer As DirectDrawSurface7
Public DD_ArrowAnim As DirectDrawSurface7
Public DD_BigSpriteSurf As DirectDrawSurface7
Public DD_NightSurf As DirectDrawSurface7
Public DD_Clip As DirectDrawClipper

Public DDSD_Primary As DDSURFACEDESC2
Public DDSD_Sprite As DDSURFACEDESC2
Public DDSD_Tile As DDSURFACEDESC2
Public DDSD_Item As DDSURFACEDESC2
Public DDSD_Emoticon As DDSURFACEDESC2
Public DDSD_BackBuffer As DDSURFACEDESC2
Public DDSD_BigSprite As DDSURFACEDESC2
Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DDSD_Night As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX()

    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
    frmEndieko.Show
    
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmEndieko.hWnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmEndieko.picScreen.hWnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
        
    ' Initialize all surfaces
    Call InitSurfaces
End Sub

Sub InitSurfaces()
Dim key As DDCOLORKEY

    ' Check for files existing
    If FileExist("\Graphics\sprites.bmp") = False Or FileExist("\Graphics\tiles.bmp") = False Or FileExist("\Graphics\items.bmp") = False Or FileExist("\Graphics\bigsprites.bmp") = False Or FileExist("\Graphics\emoticons.bmp") = False Then
        Call MsgBox("Your missing some graphic files!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If
    
    ' Set the key for masks
    With key
        .low = 0
        .high = 0
    End With
    
    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' Init sprite ddsd type and load the bitmap
    DDSD_Sprite.lFlags = DDSD_CAPS
    DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\Graphics\sprites.bmp", DDSD_Sprite)
    'Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\monster.bmp", DDSD_Sprite)
    
    DD_SpriteSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init tiles ddsd type and load the bitmap
    DDSD_Tile.lFlags = DDSD_CAPS
    DDSD_Tile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_TileSurf = DD.CreateSurfaceFromFile(App.Path & "\Graphics\tiles.bmp", DDSD_Tile)
    DD_TileSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(App.Path & "\Graphics\items.bmp", DDSD_Item)
    DD_ItemSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init big sprites ddsd type and load the bitmap
    DDSD_BigSprite.lFlags = DDSD_CAPS
    DDSD_BigSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_BigSpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\Graphics\bigsprites.bmp", DDSD_BigSprite)
    DD_BigSpriteSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init emoticons ddsd type and load the bitmap
    DDSD_Emoticon.lFlags = DDSD_CAPS
    DDSD_Emoticon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_EmoticonSurf = DD.CreateSurfaceFromFile(App.Path & "\Graphics\emoticons.bmp", DDSD_Emoticon)
    DD_EmoticonSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init arrows ddsd type and load the bitmap
    DDSD_ArrowAnim.lFlags = DDSD_CAPS
    DDSD_ArrowAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ArrowAnim = DD.CreateSurfaceFromFile(App.Path & "\Graphics\arrows.bmp", DDSD_ArrowAnim)
    DD_ArrowAnim.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init nights ddsd type and load the bitmap
    DDSD_Night.lFlags = DDSD_CAPS
    DDSD_Night.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_NightSurf = DD.CreateSurfaceFromFile(App.Path & "\Graphics\night.bmp", DDSD_Night)
     DD_NightSurf.SetColorKey DDCKEY_SRCBLT, key
End Sub

Sub DestroyDirectX()
    Set DX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_TileSurf = Nothing
    Set DD_ItemSurf = Nothing
    Set DD_BigSpriteSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    Set DD_ArrowAnim = Nothing
    Set DD_NightSurf = Nothing
End Sub

Function NeedToRestoreSurfaces() As Boolean
    Dim TestCoopRes As Long
    
    TestCoopRes = DD.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then
        NeedToRestoreSurfaces = False
    Else
        NeedToRestoreSurfaces = True
    End If
End Function

Sub BltArrow(ByVal Index As Long)
Dim X As Long, Y As Long, i As Long, z As Long
Dim BX As Long, BY As Long

For z = 1 To MAX_PLAYER_ARROWS
    If Player(Index).Arrow(z).Arrow > 0 Then
    
        rec.Top = Player(Index).Arrow(z).ArrowAnim * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(Index).Arrow(z).ArrowPosition * PIC_X
        rec.Right = rec.Left + PIC_X
       
        If GetTickCount > Player(Index).Arrow(z).ArrowTime + 30 Then
             Player(Index).Arrow(z).ArrowTime = GetTickCount
             Player(Index).Arrow(z).ArrowVarX = Player(Index).Arrow(z).ArrowVarX + 10
             Player(Index).Arrow(z).ArrowVarY = Player(Index).Arrow(z).ArrowVarY + 10
        End If
       
        If Player(Index).Arrow(z).ArrowPosition = 0 Then
             X = Player(Index).Arrow(z).ArrowX
             Y = Player(Index).Arrow(z).ArrowY + Int(Player(Index).Arrow(z).ArrowVarY / 32)
             If Y > Player(Index).Arrow(z).ArrowY + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                 Player(Index).Arrow(z).Arrow = 0
             End If
            
             If Y <= MAX_MAPY Then
                 Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             End If
        End If
       
        If Player(Index).Arrow(z).ArrowPosition = 1 Then
             X = Player(Index).Arrow(z).ArrowX
             Y = Player(Index).Arrow(z).ArrowY - Int(Player(Index).Arrow(z).ArrowVarY / 32)
             If Y < Player(Index).Arrow(z).ArrowY - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                 Player(Index).Arrow(z).Arrow = 0
             End If
            
             If Y >= 0 Then
                 Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             End If
        End If
       
        If Player(Index).Arrow(z).ArrowPosition = 2 Then
             X = Player(Index).Arrow(z).ArrowX + Int(Player(Index).Arrow(z).ArrowVarX / 32)
             Y = Player(Index).Arrow(z).ArrowY
             If X > Player(Index).Arrow(z).ArrowX + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                 Player(Index).Arrow(z).Arrow = 0
             End If
            
             If X <= MAX_MAPX Then
                 Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             End If
        End If
       
        If Player(Index).Arrow(z).ArrowPosition = 3 Then
             X = Player(Index).Arrow(z).ArrowX - Int(Player(Index).Arrow(z).ArrowVarX / 32)
             Y = Player(Index).Arrow(z).ArrowY
             If X < Player(Index).Arrow(z).ArrowX - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                 Player(Index).Arrow(z).Arrow = 0
             End If
            
             If X >= 0 Then
              Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             End If
        End If
       
        If X >= 0 And X <= MAX_MAPX Then
             If Y >= 0 And Y <= MAX_MAPY Then
                 If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                     Player(Index).Arrow(z).Arrow = 0
                 End If
             End If
        End If
       
        For i = 1 To MAX_PLAYERS
           If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                 If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                     If Index = MyIndex Then
                           Call SendData("arrowhit" & SEP_CHAR & 0 & SEP_CHAR & i & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR)
                     End If
                     If Index <> i Then Player(Index).Arrow(z).Arrow = 0
                     Exit Sub
                 End If
             End If
        Next i
       
        For i = 1 To MAX_MAP_NPCS
             If MapNpc(i).Num > 0 Then
                 If MapNpc(i).X = X And MapNpc(i).Y = Y Then
                     If Index = MyIndex Then
                           Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & i & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR)
                     End If
                     Player(Index).Arrow(z).Arrow = 0
                     Exit Sub
                 End If
             End If
        Next i
    End If
Next z
End Sub

Sub DisplayFx(ByRef surfDisplay As DirectDrawSurface7, intX As Integer, intY As Integer, intWidth As Integer, intHeight As Integer, lngROP As Long, blnFxCap As Boolean)

Dim rectSource As RECT
Dim lngSrcDC As Long
Dim lngDestDC As Long


    With rectSource
        'Set and clip
        .Top = 200
        .Bottom = 0
        .Left = 200
        .Right = 0
    End With
    
        lngDestDC = DD_BackBuffer.GetDC
        lngSrcDC = surfDisplay.GetDC
        'Do the fancy old-fashioned blit
        BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, 0, 0, lngROP
        'Release our DCs
        surfDisplay.ReleaseDC lngSrcDC
        DD_BackBuffer.ReleaseDC lngDestDC
End Sub

Sub Night()
    Call DisplayFx(DD_NightSurf, 0, 0, frmEndieko.picScreen.Width, frmEndieko.picScreen.Height, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT)
End Sub
