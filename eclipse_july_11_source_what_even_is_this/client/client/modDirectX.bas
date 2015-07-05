Attribute VB_Name = "modDirectX"
Option Explicit

Public Const TilesInSheets = 14
Public Const ExtraSheets = 6

Public DX As New DirectX7
Public DD As DirectDraw7

Public DD_Clip As DirectDrawClipper

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

Public DD_TileSurf(0 To ExtraSheets) As DirectDrawSurface7
Public DDSD_Tile(0 To ExtraSheets) As DDSURFACEDESC2
Public TileFile(0 To ExtraSheets) As Byte

Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DD_ArrowAnim As DirectDrawSurface7

Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX()

    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
    frmMirage.Show
    
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hWnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMirage.picScreen.hWnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
        
    ' Initialize all surfaces
    Call InitSurfaces
End Sub

Sub InitSurfaces()
Dim key As DDCOLORKEY
Dim i As Long

    ' Check for files existing
    If FileExist("\GFX\sprites.bmp") = False Or FileExist("\GFX\items.bmp") = False Or FileExist("\GFX\bigsprites.bmp") = False Or FileExist("\GFX\emoticons.bmp") = False Or FileExist("\GFX\arrows.bmp") = False Then
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

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal x As Long, ByVal y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

With TmpR
.Left = x
.Top = y
.Right = x
.Bottom = y
End With

TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

With TmpColorKey
.low = TheSurface.GetLockedPixel(x, y)
.high = .low
End With

TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

TheSurface.Unlock TmpR
End Sub

Sub DisplayFx(ByRef surfDisplay As DirectDrawSurface7, intX As Long, intY As Long, intWidth As Long, intHeight As Long, lngROP As Long, blnFxCap As Boolean, Tile As Long)
Dim lngSrcDC As Long
Dim lngDestDC As Long

    lngDestDC = DD_BackBuffer.GetDC
    lngSrcDC = surfDisplay.GetDC
    BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X, Int(Tile / TilesInSheets) * PIC_Y, lngROP
    surfDisplay.ReleaseDC lngSrcDC
    DD_BackBuffer.ReleaseDC lngDestDC
End Sub

Sub Night()
Dim x As Long, y As Long
Dim NewX As Long, NewY As Long
Dim NewX2 As Long, NewY2 As Long
Dim Tile As Long

If TileFile(6) = 0 Then Exit Sub
    
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
        
    For y = NewY To NewY2
        For x = NewX To NewX2
            If Map(GetPlayerMap(MyIndex)).Tile(x, y).Light <= 0 Then
                DisplayFx DD_TileSurf(6), (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31
            Else
                DisplayFx DD_TileSurf(6), (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Map(GetPlayerMap(MyIndex)).Tile(x, y).Light
            End If
        Next x
    Next y
End Sub
Sub WierdNight()
Dim x As Long, y As Long
Dim NewX As Long, NewY As Long
Dim NewX2 As Long, NewY2 As Long
Dim Tile As Long

If TileFile(6) = 0 Then Exit Sub
    
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
        
    For y = NewY To NewY2
        For x = NewX To NewX2
            If Map(GetPlayerMap(MyIndex)).Tile(x, y).Light <= 0 Then
                DisplayFx DD_TileSurf(5), (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31
            Else
                DisplayFx DD_TileSurf(5), (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Map(GetPlayerMap(MyIndex)).Tile(x, y).Light
            End If
        Next x
    Next y
End Sub
Sub BltWeather()
Dim i As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))
    
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
    
    For i = 1 To MAX_RAINDROPS
        If Not ((DropRain(i).x = 0) Or (DropRain(i).y = 0)) Then
            DropRain(i).x = DropRain(i).x + DropRain(i).Speed
            DropRain(i).y = DropRain(i).y + DropRain(i).Speed
            Call DD_BackBuffer.DrawLine(DropRain(i).x, DropRain(i).y, DropRain(i).x + DropRain(i).Speed, DropRain(i).y + DropRain(i).Speed)
            If (DropRain(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                DropRain(i).Randomized = False
            End If
        End If
    Next i
    If TileFile(6) = 1 Then
        rec.Top = Int(14 / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (14 - Int(14 / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
            
        For i = 1 To MAX_RAINDROPS
            If Not ((DropSnow(i).x = 0) Or (DropSnow(i).y = 0)) Then
                DropSnow(i).x = DropSnow(i).x + DropSnow(i).Speed
                DropSnow(i).y = DropSnow(i).y + DropSnow(i).Speed
                Call DD_BackBuffer.BltFast(DropSnow(i).x + DropSnow(i).Speed, DropSnow(i).y + DropSnow(i).Speed, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropSnow(i).Randomized = False
                End If
            End If
        Next i
    End If
        
    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then
        If Int((100 - 1 + 1) * Rnd) + 1 = 8 Then
            DD_BackBuffer.SetFillColor RGB(255, 255, 255)
            Call PlaySound("Thunder.wav")
            Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
        End If
    End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).x > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropRain(RDNumber).Speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropRain(RDNumber).x = 0
    DropRain(RDNumber).y = 0
    DropRain(RDNumber).Speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropSnow(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropSnow(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropSnow(RDNumber).x > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropSnow(RDNumber).Speed = Int((10 * Rnd) + 6)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropSnow(RDNumber).x = 0
    DropSnow(RDNumber).y = 0
    DropSnow(RDNumber).Speed = 0
    DropSnow(RDNumber).Randomized = False
End Sub

Sub BltSpell(ByVal index As Long)
Dim x As Long, y As Long, i As Long

If Player(index).SpellNum <= 0 Or Player(index).SpellNum > MAX_SPELLS Then Exit Sub
If Spell(Player(index).SpellNum).SpellAnim <= 0 Then Exit Sub

For i = 1 To MAX_SPELL_ANIM
    If Player(index).SpellAnim(i).CastedSpell = YES Then
        If Player(index).SpellAnim(i).SpellDone < Spell(Player(index).SpellNum).SpellDone Then
            If Player(index).SpellAnim(i).SpellVar > 10 Then
                Player(index).SpellAnim(i).SpellDone = Player(index).SpellAnim(i).SpellDone + 1
                Player(index).SpellAnim(i).SpellVar = 0
            End If
            If GetTickCount > Player(index).SpellAnim(i).SpellTime + Spell(Player(index).SpellNum).SpellTime Then
                Player(index).SpellAnim(i).SpellTime = GetTickCount
                Player(index).SpellAnim(i).SpellVar = Player(index).SpellAnim(i).SpellVar + 1
            End If
                        
            rec.Top = Spell(Player(index).SpellNum).SpellAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = Player(index).SpellAnim(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X
            
            If Player(index).SpellAnim(i).TargetType = 0 Then
                If Player(index).SpellAnim(i).Target > 0 Then
                    If Player(index).SpellAnim(i).Target = MyIndex Then
                        x = NewX + sx
                        y = NewY + sx
                        Call DD_BackBuffer.BltFast(x, y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        x = GetPlayerX(Player(index).SpellAnim(i).Target) * PIC_X + sx + Player(Player(index).SpellAnim(i).Target).XOffset
                        y = GetPlayerY(Player(index).SpellAnim(i).Target) * PIC_Y + sx + Player(Player(index).SpellAnim(i).Target).YOffset
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            Else
                x = MapNpc(Player(index).SpellAnim(i).Target).x * PIC_X + sx + MapNpc(Player(index).SpellAnim(i).Target).XOffset
                y = MapNpc(Player(index).SpellAnim(i).Target).y * PIC_Y + sx + MapNpc(Player(index).SpellAnim(i).Target).YOffset
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Else
            Player(index).SpellAnim(i).CastedSpell = NO
        End If
    End If
Next i
End Sub

Sub BltEmoticons(ByVal index As Long)
Dim x2 As Long, y2 As Long
Dim ETime As Long
ETime = 1300
   
    If Player(index).EmoticonNum < 0 Then Exit Sub
    
    If Player(index).EmoticonTime + ETime > GetTickCount Then
        If GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 1) Then
            Player(index).EmoticonVar = 0
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 2) Then
            Player(index).EmoticonVar = 1
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 3) Then
            Player(index).EmoticonVar = 2
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 4) Then
            Player(index).EmoticonVar = 3
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 5) Then
            Player(index).EmoticonVar = 4
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 6) Then
            Player(index).EmoticonVar = 5
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 7) Then
            Player(index).EmoticonVar = 6
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 8) Then
            Player(index).EmoticonVar = 7
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 9) Then
            Player(index).EmoticonVar = 8
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 10) Then
            Player(index).EmoticonVar = 9
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 11) Then
            Player(index).EmoticonVar = 10
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 12) Then
            Player(index).EmoticonVar = 11
        End If
        
        rec.Top = Player(index).EmoticonNum * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(index).EmoticonVar * PIC_X
        rec.Right = rec.Left + PIC_X
        
        If index = MyIndex Then
            x2 = NewX + sx + 16
            y2 = NewY + sx - 32
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            x2 = GetPlayerX(index) * PIC_X + sx + Player(index).XOffset + 16
            y2 = GetPlayerY(index) * PIC_Y + sx + Player(index).YOffset - 32
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltArrow(ByVal index As Long)
Dim x As Long, y As Long, i As Long, z As Long
Dim BX As Long, BY As Long

For z = 1 To MAX_PLAYER_ARROWS
    If Player(index).Arrow(z).Arrow > 0 Then
    
        rec.Top = Player(index).Arrow(z).ArrowAnim * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(index).Arrow(z).ArrowPosition * PIC_X
        rec.Right = rec.Left + PIC_X
        
        If GetTickCount > Player(index).Arrow(z).ArrowTime + 30 Then
            Player(index).Arrow(z).ArrowTime = GetTickCount
            Player(index).Arrow(z).ArrowVarX = Player(index).Arrow(z).ArrowVarX + 10
            Player(index).Arrow(z).ArrowVarY = Player(index).Arrow(z).ArrowVarY + 10
        End If
        
        If Player(index).Arrow(z).ArrowPosition = 0 Then
            x = Player(index).Arrow(z).ArrowX
            y = Player(index).Arrow(z).ArrowY + Int(Player(index).Arrow(z).ArrowVarY / 32)
            If y > Player(index).Arrow(z).ArrowY + Arrows(Player(index).Arrow(z).ArrowNum).Range - 2 Then
                Player(index).Arrow(z).Arrow = 0
            End If
            
            If y <= MAX_MAPY Then
                Call DD_BackBuffer.BltFast((Player(index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Player(index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If Player(index).Arrow(z).ArrowPosition = 1 Then
            x = Player(index).Arrow(z).ArrowX
            y = Player(index).Arrow(z).ArrowY - Int(Player(index).Arrow(z).ArrowVarY / 32)
            If y < Player(index).Arrow(z).ArrowY - Arrows(Player(index).Arrow(z).ArrowNum).Range + 2 Then
                Player(index).Arrow(z).Arrow = 0
            End If
            
            If y >= 0 Then
                Call DD_BackBuffer.BltFast((Player(index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Player(index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If Player(index).Arrow(z).ArrowPosition = 2 Then
            x = Player(index).Arrow(z).ArrowX + Int(Player(index).Arrow(z).ArrowVarX / 32)
            y = Player(index).Arrow(z).ArrowY
            If x > Player(index).Arrow(z).ArrowX + Arrows(Player(index).Arrow(z).ArrowNum).Range - 2 Then
                Player(index).Arrow(z).Arrow = 0
            End If
            
            If x <= MAX_MAPX Then
                Call DD_BackBuffer.BltFast((Player(index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(index).Arrow(z).ArrowVarX, (Player(index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If Player(index).Arrow(z).ArrowPosition = 3 Then
            x = Player(index).Arrow(z).ArrowX - Int(Player(index).Arrow(z).ArrowVarX / 32)
            y = Player(index).Arrow(z).ArrowY
            If x < Player(index).Arrow(z).ArrowX - Arrows(Player(index).Arrow(z).ArrowNum).Range + 2 Then
                Player(index).Arrow(z).Arrow = 0
            End If
            
            If x >= 0 Then
             Call DD_BackBuffer.BltFast((Player(index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(index).Arrow(z).ArrowVarX, (Player(index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If x >= 0 And x <= MAX_MAPX Then
            If y >= 0 And y <= MAX_MAPY Then
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Player(index).Arrow(z).Arrow = 0
                End If
            End If
        End If
        
        For i = 1 To MAX_PLAYERS
           If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(i) = x And GetPlayerY(i) = y Then
                    If index = MyIndex Then
                        Call SendData("arrowhit" & SEP_CHAR & 0 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                    End If
                    If index <> i Then Player(index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(i).Num > 0 Then
                If MapNpc(i).x = x And MapNpc(i).y = y Then
                    If index = MyIndex Then
                        Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                    End If
                    Player(index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next i
        
        For BX = 0 To MAX_MAPX
            For BY = 0 To MAX_MAPY
                If Map(GetPlayerMap(MyIndex)).Tile(BX, BY).Type = TILE_TYPE_NPC_SPAWN Then
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        If MapAttributeNpc(i, BX, BY).x = x And MapAttributeNpc(i, BX, BY).y = y Then
                            If index = MyIndex Then
                                Call SendData("arrowhit" & SEP_CHAR & 2 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & BX & SEP_CHAR & BY & SEP_CHAR & END_CHAR)
                            End If
                            Player(index).Arrow(z).Arrow = 0
                            Exit Sub
                        End If
                    Next i
                End If
            Next BY
        Next BX
    End If
Next z
End Sub
