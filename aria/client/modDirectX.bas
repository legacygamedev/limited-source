Attribute VB_Name = "modDirectX"
Option Explicit

Public Const TilesInSheets = 14
Public ExtraSheets As Long

Public DX As New DirectX7
Public DD As DirectDraw7

Public DD_Clip As DirectDrawClipper
Public DD_MiniClip As DirectDrawClipper

Public DD_PrimarySurf As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

Public DD_SpriteSurf As DirectDrawSurface7
Public DDSD_Sprite As DDSURFACEDESC2

Public DD_ItemSurf As DirectDrawSurface7
Public DDSD_Item As DDSURFACEDESC2

Public DD_EmoticonSurf As DirectDrawSurface7
Public DDSD_Emoticon As DDSURFACEDESC2

Public DD_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

Public DD_BigSpriteSurf As DirectDrawSurface7
Public DDSD_BigSprite As DDSURFACEDESC2

Public DD_SpellAnim As DirectDrawSurface7
Public DDSD_SpellAnim As DDSURFACEDESC2

Public DD_TileSurf() As DirectDrawSurface7
Public DDSD_Tile() As DDSURFACEDESC2
Public TileFile() As Byte

Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DD_ArrowAnim As DirectDrawSurface7

Public DDSD_MiniMap As DDSURFACEDESC2
Public DD_MiniMap As DirectDrawSurface7

Public DDSD_MiniBuff As DDSURFACEDESC2
Public DD_MiniBuff As DirectDrawSurface7

Public DDSD_MiniOver As DDSURFACEDESC2
Public DD_MiniOver As DirectDrawSurface7

Public Rec As RECT
Public rec_pos As RECT

Sub InitDirectX()

    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
    
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hwnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMirage.picScreen.hwnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
    
    ' Initialize all surfaces
    Call InitSurfaces
End Sub

Sub InitSurfaces()
Dim key As DDCOLORKEY
Dim i As Long

    ' Check for files existing
    If FileExist("\GFX\sprites.bmp") = False Or FileExist("\GFX\items.bmp") = False Or FileExist("\GFX\bigsprites.bmp") = False Or FileExist("\GFX\emoticons.bmp") = False Or FileExist("\GFX\arrows.bmp") = False Or FileExist("\GFX\minimap.bmp") = False Then
        Call MsgBox("Your missing some graphic files!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If
    
    ' Set the key for masks
    key.low = 0
    key.high = 0
    
    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' Initialize back buffer
    DDSD_MiniBuff.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_MiniBuff.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_MiniBuff.lWidth = (MAX_MAPX + 1) * 6
    DDSD_MiniBuff.lHeight = (MAX_MAPY + 1) * 6
    Set DD_MiniBuff = DD.CreateSurface(DDSD_MiniBuff)
    
    ' Init sprite ddsd type and load the bitmap
    DDSD_Sprite.lFlags = DDSD_CAPS
    DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\sprites.bmp", DDSD_Sprite)
    SetMaskColorFromPixel DD_SpriteSurf, 0, 0
    
    ' Init tiles ddsd type and load the bitmap
    For i = 0 To ExtraSheets
        If Dir(App.Path & "\GFX\tiles" & i & ".bmp") <> "" Then
            DDSD_Tile(i).lFlags = DDSD_CAPS
            DDSD_Tile(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_TileSurf(i) = DD.CreateSurfaceFromFile(App.Path & "\GFX\tiles" & i & ".bmp", DDSD_Tile(i))
            SetMaskColorFromPixel DD_TileSurf(i), 0, 0
            TileFile(i) = 1
        Else
            TileFile(i) = 0
        End If
    Next i
    
    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\items.bmp", DDSD_Item)
    SetMaskColorFromPixel DD_ItemSurf, 0, 0
    
    ' Init big sprites ddsd type and load the bitmap
    DDSD_BigSprite.lFlags = DDSD_CAPS
    DDSD_BigSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_BigSpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\bigsprites.bmp", DDSD_BigSprite)
    SetMaskColorFromPixel DD_BigSpriteSurf, 0, 0
    
    ' Init emoticons ddsd type and load the bitmap
    DDSD_Emoticon.lFlags = DDSD_CAPS
    DDSD_Emoticon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_EmoticonSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\emoticons.bmp", DDSD_Emoticon)
    SetMaskColorFromPixel DD_EmoticonSurf, 0, 0
    
    ' Init spells ddsd type and load the bitmap
    DDSD_SpellAnim.lFlags = DDSD_CAPS
    DDSD_SpellAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpellAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\spells.bmp", DDSD_SpellAnim)
    SetMaskColorFromPixel DD_SpellAnim, 0, 0
    
    ' Init arrows ddsd type and load the bitmap
    DDSD_ArrowAnim.lFlags = DDSD_CAPS
    DDSD_ArrowAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ArrowAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\arrows.bmp", DDSD_ArrowAnim)
    SetMaskColorFromPixel DD_ArrowAnim, 0, 0
    
    ' Init Mini Map ddsd type and load the bitmap
    DDSD_MiniMap.lFlags = DDSD_CAPS
    DDSD_MiniMap.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_MiniMap = DD.CreateSurfaceFromFile(App.Path & "\GFX\minimap.bmp", DDSD_MiniMap)
    SetMaskColor DD_MiniMap, 0
    
    ' Init Mini Map ddsd type and load the bitmap
    DDSD_MiniOver.lFlags = DDSD_CAPS
    DDSD_MiniOver.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_MiniOver = DD.CreateSurfaceFromFile(App.Path & "\GFX\minimapoverlay.bmp", DDSD_MiniOver)
    SetMaskColor DD_MiniOver, 0
End Sub

Sub DestroyDirectX()
Dim i As Long

    Set DX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    For i = 0 To ExtraSheets
        If TileFile(i) = 1 Then
            Set DD_TileSurf(i) = Nothing
        End If
    Next i
    Set DD_ItemSurf = Nothing
    Set DD_BigSpriteSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    Set DD_SpellAnim = Nothing
    Set DD_ArrowAnim = Nothing
    Set DD_MiniMap = Nothing
    Set DD_MiniBuff = Nothing
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
.low = TheSurface.GetLockedPixel(X, Y)
.high = .low
End With

TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

TheSurface.Unlock TmpR
End Sub

Public Sub SetMaskColor(ByRef TheSurface As DirectDrawSurface7, ByVal Color As Long)
Dim TmpColorKey As DDCOLORKEY

    With TmpColorKey
        .low = Color
        .high = Color
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
End Sub

Sub DisplayFx(ByRef surfDisplay As DirectDrawSurface7, intX As Long, intY As Long, intWidth As Long, intHeight As Long, lngROP As Long, blnFxCap As Boolean, Tile As Long)
Dim lngSrcDC As Long
Dim lngDestDC As Long
'Dim srcRect As RECT

    lngDestDC = DD_BackBuffer.GetDC
    lngSrcDC = surfDisplay.GetDC
    BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X, Int(Tile / TilesInSheets) * PIC_Y, lngROP
    surfDisplay.ReleaseDC lngSrcDC
    DD_BackBuffer.ReleaseDC lngDestDC
    
    'srcRect.Left = (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X
    'srcRect.Top = Int(Tile / TilesInSheets) * PIC_Y
    'srcRect.Bottom = srcRect.Top + intHeight
    'srcRect.Right = srcRect.Left + intWidth
    
    'AlphaBlend surfDisplay, srcRect, DD_BackBuffer, intX, intY, 155
End Sub

Sub Night()
Dim X As Long, Y As Long
Dim NewX As Long, NewY As Long
Dim NewX2 As Long, NewY2 As Long
Dim Tile As Long, Found As Boolean
Dim i As Long, J As Long
Dim NoDraw(0 To 30, 0 To 30) As Boolean

If TileFile(ExtraSheets) = 0 Then Exit Sub
    
    NewX = GetPlayerX(MyIndex) - 11
    NewY = GetPlayerY(MyIndex) - 8
    
    NewX2 = GetPlayerX(MyIndex) + 10
    NewY2 = GetPlayerY(MyIndex) + 8
    
    If NewX < 0 Then
        NewX = 0
        NewX2 = 20
    ElseIf NewX2 > MAX_MAPX Then
        NewX2 = MAX_MAPX
        NewX = MAX_MAPX - 20
    End If
    
    If NewY < 0 Then
        NewY = 0
        NewY2 = 15
    ElseIf NewY2 > MAX_MAPY Then
        NewY2 = MAX_MAPY
        NewY = MAX_MAPY - 15
    End If

    If MAX_MAPX = 19 Then
        NewY = 0
        NewY2 = MAX_MAPY
        NewX = 0
        NewX2 = MAX_MAPX
    End If
        
    For i = 1 To MAX_PLAYERS
        If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If GetPlayerWeaponSlot(i) > 0 Then
                If Item(GetPlayerInvItemNum(i, GetPlayerWeaponSlot(i))).Type = ITEM_TYPE_LAMP Then
                    If GetPlayerX(i) > 0 And GetPlayerY(i) > 0 Then NoDraw(GetPlayerX(i) - 1, GetPlayerY(i) - 1) = True
                    If GetPlayerY(i) > 0 Then NoDraw(GetPlayerX(i), GetPlayerY(i) - 1) = True
                    If GetPlayerX(i) < MAX_MAPX And GetPlayerY(i) > 0 Then NoDraw(GetPlayerX(i) + 1, GetPlayerY(i) - 1) = True
                    If GetPlayerX(i) > 0 Then NoDraw(GetPlayerX(i) - 1, GetPlayerY(i)) = True
                    NoDraw(GetPlayerX(i), GetPlayerY(i)) = True
                    If GetPlayerX(i) < MAX_MAPX Then NoDraw(GetPlayerX(i) + 1, GetPlayerY(i)) = True
                    If GetPlayerX(i) > 0 And GetPlayerY(i) < MAX_MAPY Then NoDraw(GetPlayerX(i) - 1, GetPlayerY(i) + 1) = True
                    If GetPlayerY(i) < MAX_MAPY Then NoDraw(GetPlayerX(i), GetPlayerY(i) + 1) = True
                    If GetPlayerX(i) < MAX_MAPX And GetPlayerY(i) < MAX_MAPY Then NoDraw(GetPlayerX(i) + 1, GetPlayerY(i) + 1) = True
                    
                    DisplayFx DD_TileSurf(ExtraSheets), (GetPlayerX(i) - 1 - NewPlayerX) * PIC_X + sx - NewXOffset, (GetPlayerY(i) - 1 - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 28
                    DisplayFx DD_TileSurf(ExtraSheets), (GetPlayerX(i) - NewPlayerX) * PIC_X + sx - NewXOffset, (GetPlayerY(i) - 1 - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 29
                    DisplayFx DD_TileSurf(ExtraSheets), (GetPlayerX(i) + 1 - NewPlayerX) * PIC_X + sx - NewXOffset, (GetPlayerY(i) - 1 - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 30
                    DisplayFx DD_TileSurf(ExtraSheets), (GetPlayerX(i) - 1 - NewPlayerX) * PIC_X + sx - NewXOffset, (GetPlayerY(i) - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 42
                    DisplayFx DD_TileSurf(ExtraSheets), (GetPlayerX(i) - NewPlayerX) * PIC_X + sx - NewXOffset, (GetPlayerY(i) - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 43
                    DisplayFx DD_TileSurf(ExtraSheets), (GetPlayerX(i) + 1 - NewPlayerX) * PIC_X + sx - NewXOffset, (GetPlayerY(i) - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 44
                    DisplayFx DD_TileSurf(ExtraSheets), (GetPlayerX(i) - 1 - NewPlayerX) * PIC_X + sx - NewXOffset, (GetPlayerY(i) + 1 - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 56
                    DisplayFx DD_TileSurf(ExtraSheets), (GetPlayerX(i) - NewPlayerX) * PIC_X + sx - NewXOffset, (GetPlayerY(i) + 1 - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 57
                    DisplayFx DD_TileSurf(ExtraSheets), (GetPlayerX(i) + 1 - NewPlayerX) * PIC_X + sx - NewXOffset, (GetPlayerY(i) + 1 - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 58
                End If
            End If
        End If
    Next i
        
    For Y = NewY To NewY2
        For X = NewX To NewX2
            If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light <= 0 Then
                If Not NoDraw(X, Y) Then
                    DisplayFx DD_TileSurf(ExtraSheets), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31
                End If
            Else
                DisplayFx DD_TileSurf(ExtraSheets), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light
            End If
        Next X
    Next Y
End Sub

Sub BltWeather()
Dim i As Long
Dim X As Long
Dim Y As Long
    
    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For i = 1 To MAX_RAINDROPS
            If DropRain(i).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = "" Then
                        frmMirage.tmrRainDrop.Interval = 200
                        frmMirage.tmrRainDrop.Tag = "123"
                    End If
                End If
            End If
        Next i
        
        Rec.Top = 32
        Rec.Bottom = 64
        Rec.Left = 6 * PIC_X
        Rec.Right = Rec.Left + PIC_X
        Y = Int(Rnd * (5 * (RainIntensity / 50)))
        For X = 0 To Y
            Call DD_BackBuffer.BltFast(Int(Rnd * (MAX_MAPX * PIC_X)), Int(Rnd * (MAX_MAPY * PIC_Y)), DD_TileSurf(ExtraSheets), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Next X
        
    ElseIf GameWeather = WEATHER_SNOWING Then
        For i = 1 To MAX_RAINDROPS
            If DropSnow(i).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                    If frmMirage.tmrSnowDrop.Tag = "" Then
                        frmMirage.tmrSnowDrop.Interval = 200
                        frmMirage.tmrSnowDrop.Tag = "123"
                    End If
                End If
            End If
        Next i
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= RainIntensity Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = ""
    End If
    
    If TileFile(ExtraSheets) = 1 Then
        Rec.Top = 0
        Rec.Bottom = 32
        Rec.Left = 6 * PIC_X
        Rec.Right = Rec.Left + PIC_X
        
        For i = 1 To MAX_RAINDROPS
            If Not ((DropRain(i).X = 0) Or (DropRain(i).Y = 0)) Then
                DropRain(i).X = DropRain(i).X + DropRain(i).Speed
                DropRain(i).Y = DropRain(i).Y + DropRain(i).Speed
                Call DD_BackBuffer.BltFast(DropRain(i).X + DropRain(i).Speed, DropRain(i).Y + DropRain(i).Speed, DD_TileSurf(ExtraSheets), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropRain(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).Y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropRain(i).Randomized = False
                End If
            End If
        Next i
    End If
    
    If TileFile(ExtraSheets) = 1 Then
        Rec.Top = 32
        Rec.Bottom = 64
        Rec.Left = 0
        Rec.Right = 32
            
        For i = 1 To MAX_RAINDROPS
            If Not ((DropSnow(i).X = 0) Or (DropSnow(i).Y = 0)) Then
                DropSnow(i).X = DropSnow(i).X + DropSnow(i).Speed * (-1 * (DropSnow(i).LeftRight * 2 - 1))
                DropSnow(i).Y = DropSnow(i).Y + DropSnow(i).Speed
                
                If GetTickCount > DropSnow(i).SwitchTime + 250 Then
                    DropSnow(i).LeftRight = Int(Rnd * 2)
                    If DropSnow(i).LeftRight > 1 Or DropSnow(i).LeftRight < 0 Then DropSnow(i).LeftRight = 1
                    DropSnow(i).SwitchTime = GetTickCount
                End If
                
                Call DD_BackBuffer.BltFast(DropSnow(i).X + DropSnow(i).Speed, DropSnow(i).Y + DropSnow(i).Speed, DD_TileSurf(ExtraSheets), Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).Y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropSnow(i).Randomized = False
                End If
            End If
        Next i
    End If
        
    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then
        If Int((((50 - RainIntensity) / 50 * 75) + 25) * Rnd) + 1 = 1 Then
            Y = Rnd * 0.65 + 0.35
            DD_BackBuffer.SetFillColor RGB(255 * Y, 255 * Y, 255 * Y)
            Call PlaySoundAnyBuffer("Thunder.wav", 20 * Rnd + 60, 20 * Rnd + 60)
            Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
        End If
    End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).Y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).Y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).X > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropRain(RDNumber).Speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropRain(RDNumber).X = 0
    DropRain(RDNumber).Y = 0
    DropRain(RDNumber).Speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropSnow(RDNumber).Y = 1
    DropSnow(RDNumber).Speed = Int((10 * Rnd) + 2)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropSnow(RDNumber).X = 0
    DropSnow(RDNumber).Y = 0
    DropSnow(RDNumber).Speed = 0
    DropSnow(RDNumber).Randomized = False
End Sub

Sub BltSpell(ByVal Index As Long)
Dim X As Long, Y As Long, i As Long

If Player(Index).SpellNum <= 0 Or Player(Index).SpellNum > MAX_SPELLS Then Exit Sub
If Spell(Player(Index).SpellNum).SpellAnim <= 0 Then Exit Sub

For i = 1 To MAX_SPELL_ANIM
    If Player(Index).SpellAnim(i).CastedSpell = YES Then
        If Player(Index).SpellAnim(i).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then
            If Player(Index).SpellAnim(i).SpellVar > 10 Then
                Player(Index).SpellAnim(i).SpellDone = Player(Index).SpellAnim(i).SpellDone + 1
                Player(Index).SpellAnim(i).SpellVar = 0
            End If
            If GetTickCount > Player(Index).SpellAnim(i).SpellTime + Spell(Player(Index).SpellNum).SpellTime Then
                Player(Index).SpellAnim(i).SpellTime = GetTickCount
                Player(Index).SpellAnim(i).SpellVar = Player(Index).SpellAnim(i).SpellVar + 1
            End If
                        
            Rec.Top = Spell(Player(Index).SpellNum).SpellAnim * PIC_Y
            Rec.Bottom = Rec.Top + PIC_Y
            Rec.Left = Player(Index).SpellAnim(i).SpellVar * PIC_X
            Rec.Right = Rec.Left + PIC_X
            
            If Player(Index).SpellAnim(i).TargetType = 0 Then
                If Player(Index).SpellAnim(i).Target > 0 Then
                    If Player(Index).SpellAnim(i).Target = MyIndex Then
                        X = NewX + sx
                        Y = NewY + sx
                        Call DD_BackBuffer.BltFast(X, Y, DD_SpellAnim, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        X = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx + Player(Player(Index).SpellAnim(i).Target).XOffset
                        Y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx + Player(Player(Index).SpellAnim(i).Target).YOffset
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            Else
                X = MapNpc(Player(Index).SpellAnim(i).Target).X * PIC_X + sx + MapNpc(Player(Index).SpellAnim(i).Target).XOffset
                Y = MapNpc(Player(Index).SpellAnim(i).Target).Y * PIC_Y + sx + MapNpc(Player(Index).SpellAnim(i).Target).YOffset
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Else
            Player(Index).SpellAnim(i).CastedSpell = NO
        End If
    End If
Next i
End Sub

Sub BltEmoticons(ByVal Index As Long)
Dim x2 As Long, y2 As Long
Dim ETime As Long
ETime = 1300
   
    If Player(Index).EmoticonNum < 0 Then Exit Sub
    
    If Player(Index).EmoticonTime + ETime > GetTickCount Then
        If GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 1) Then
            Player(Index).EmoticonVar = 0
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 2) Then
            Player(Index).EmoticonVar = 1
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 3) Then
            Player(Index).EmoticonVar = 2
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 4) Then
            Player(Index).EmoticonVar = 3
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 5) Then
            Player(Index).EmoticonVar = 4
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 6) Then
            Player(Index).EmoticonVar = 5
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 7) Then
            Player(Index).EmoticonVar = 6
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 8) Then
            Player(Index).EmoticonVar = 7
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 9) Then
            Player(Index).EmoticonVar = 8
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 10) Then
            Player(Index).EmoticonVar = 9
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 11) Then
            Player(Index).EmoticonVar = 10
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 12) Then
            Player(Index).EmoticonVar = 11
        End If
        
        Rec.Top = Player(Index).EmoticonNum * PIC_Y
        Rec.Bottom = Rec.Top + PIC_Y
        Rec.Left = Player(Index).EmoticonVar * PIC_X
        Rec.Right = Rec.Left + PIC_X
        
        If Index = MyIndex Then
            x2 = NewX + sx + 16
            y2 = NewY + sx - 32
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            x2 = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + 16
            y2 = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - 32
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_EmoticonSurf, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltArrow(ByVal Index As Long)
Dim X As Long, Y As Long, i As Long, z As Long
Dim BX As Long, BY As Long

For z = 1 To MAX_PLAYER_ARROWS
    If Player(Index).Arrow(z).Arrow > 0 Then
    
        Rec.Top = Player(Index).Arrow(z).ArrowAnim * PIC_Y
        Rec.Bottom = Rec.Top + PIC_Y
        Rec.Left = Player(Index).Arrow(z).ArrowPosition * PIC_X
        Rec.Right = Rec.Left + PIC_X
        
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
                Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If Player(Index).Arrow(z).ArrowPosition = 1 Then
            X = Player(Index).Arrow(z).ArrowX
            Y = Player(Index).Arrow(z).ArrowY - Int(Player(Index).Arrow(z).ArrowVarY / 32)
            If Y < Player(Index).Arrow(z).ArrowY - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                Player(Index).Arrow(z).Arrow = 0
            End If
            
            If Y >= 0 Then
                Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If Player(Index).Arrow(z).ArrowPosition = 2 Then
            X = Player(Index).Arrow(z).ArrowX + Int(Player(Index).Arrow(z).ArrowVarX / 32)
            Y = Player(Index).Arrow(z).ArrowY
            If X > Player(Index).Arrow(z).ArrowX + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                Player(Index).Arrow(z).Arrow = 0
            End If
            
            If X <= MAX_MAPX Then
                Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If Player(Index).Arrow(z).ArrowPosition = 3 Then
            X = Player(Index).Arrow(z).ArrowX - Int(Player(Index).Arrow(z).ArrowVarX / 32)
            Y = Player(Index).Arrow(z).ArrowY
            If X < Player(Index).Arrow(z).ArrowX - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                Player(Index).Arrow(z).Arrow = 0
            End If
            
            If X >= 0 Then
             Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If X >= 0 And X <= MAX_MAPX Then
            If Y >= 0 And Y <= MAX_MAPY Then
                If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_BLOCKICON Then
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

Sub BltMiniMap(ByVal MapNum As Long)
Dim i As Long
Dim X As Integer
Dim Y As Integer
Dim MMx As Long
Dim MMy As Integer

    If MapNum = 0 Then
        MapNum = GetPlayerMap(MyIndex)
        MiniMapMap = MapNum
    End If

    ' Tiles Layer
    Rec.Top = 0
    Rec.Bottom = 6
    Rec.Left = 0
    Rec.Right = 6
   
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            MMx = X * 6
            MMy = Y * 6
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                Rec.Top = 6
                Rec.Bottom = 12
                Rec.Left = 0
                Rec.Right = 6
                Call DD_MiniBuff.BltFast(MMx, MMy, DD_MiniMap, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                Rec.Top = 0
                Rec.Bottom = 6
                Rec.Left = 0
                Rec.Right = 6
                Call DD_MiniBuff.BltFast(MMx, MMy, DD_MiniMap, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next X
    Next Y
    
    ' Map Item Layer
    Rec.Top = 30
    Rec.Bottom = 36
    Rec.Left = 0
    Rec.Right = 6
   
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(i).Num > 0 And MapNum = GetPlayerMap(MyIndex) Then
            X = MapItem(i).X
            Y = MapItem(i).Y
            MMx = X * 6
            MMy = Y * 6
            Call DD_MiniBuff.BltFast(MMx, MMy, DD_MiniMap, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Next i
   
    ' Player Layer
    Rec.Top = 12
    Rec.Bottom = 18
    Rec.Left = 0
    Rec.Right = 6
   
    For i = 1 To MAX_PLAYERS
        If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            X = Player(i).X
            Y = Player(i).Y
            MMx = X * 6
            MMy = Y * 6
            If i <> MyIndex Then
                Call DD_MiniBuff.BltFast(MMx, MMy, DD_MiniMap, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Next i
    
    ' MyPlayer Layer
    If MapNum = GetPlayerMap(MyIndex) Then
        Rec.Top = 18
        Rec.Bottom = 24
        Rec.Left = 0
        Rec.Right = 6
        X = Player(MyIndex).X
        Y = Player(MyIndex).Y
        MMx = X * 6
        MMy = Y * 6
        Call DD_MiniBuff.BltFast(MMx, MMy, DD_MiniMap, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
   
    ' NPC Layer
    Rec.Top = 24
    Rec.Bottom = 30
    Rec.Left = 0
    Rec.Right = 6
   
    For i = 1 To MAX_MAP_NPCS
        If Map(MapNum).Npc(i) And MapNpc(i).HP > 0 And MapNum = GetPlayerMap(MyIndex) Then
            X = MapNpc(i).X
            Y = MapNpc(i).Y
            MMx = X * 6
            MMy = Y * 6
            Call DD_MiniBuff.BltFast(MMx, MMy, DD_MiniMap, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Next i
    
   
   'MiniMap Icons Layer
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            MMx = X * 6
            MMy = Y * 6
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_MINICON Or Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKICON Then
                Rec.Top = Map(MapNum).Tile(X, Y).Data1 * 6
                Rec.Bottom = (Map(MapNum).Tile(X, Y).Data1 + 1) * 6
                Rec.Left = 0
                Rec.Right = 6
                Call DD_MiniBuff.BltFast(MMx, MMy, DD_MiniMap, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next X
    Next Y
    
    With Rec
        .Top = 0
        .Left = 0
        .Right = (MAX_MAPX + 1) * 6
        .Bottom = (MAX_MAPY + 1) * 6
    End With
    
    Call DD_MiniBuff.BltFast(0, 0, DD_MiniOver, Rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Call DD_MiniBuff.BltToDC(frmMirage.picMini.hDC, Rec, Rec)
End Sub
