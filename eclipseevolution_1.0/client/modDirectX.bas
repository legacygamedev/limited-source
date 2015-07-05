Attribute VB_Name = "modDirectX"
Option Explicit

Public Const TilesInSheets = 14
Public Const ExtraSheets = 10

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

Public DD_BigSpellAnim As DirectDrawSurface7
Public DDSD_BigSpellAnim As DDSURFACEDESC2

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
    DDSD_Primary.lFlags = DDSD_CAPS And DDSD_ALPHABITDEPTH
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    DDSD_Primary.lAlphaBitDepth = 32
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
Dim Key As DDCOLORKEY
Dim i As Long
Dim DC As Long
Dim strfilename As String
Dim BMU As BitmapUtils
Set BMU = New BitmapUtils

    ' Check for files existing
    If FileExist("\GFX\sprites." & Trim$(ENCRYPT_TYPE)) = False Or FileExist("\GFX\items." & Trim$(ENCRYPT_TYPE)) = False Or FileExist("\GFX\bigsprites." & Trim$(ENCRYPT_TYPE)) = False Or FileExist("\GFX\emoticons." & Trim$(ENCRYPT_TYPE)) = False Or FileExist("\GFX\arrows." & Trim$(ENCRYPT_TYPE)) = False Then
        Call MsgBox("Your missing some GFX files!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If
    
    ' Set the key for masks
    Key.low = 0
    Key.high = 0
    
    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    If Trim$(ENCRYPT_TYPE) = "BMP" Then
    
        ' Init sprite ddsd type and load the bitmap
        DDSD_Sprite.lFlags = DDSD_CAPS
        DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\sprites.bmp", DDSD_Sprite)
        SetMaskColorFromPixel DD_SpriteSurf, 0, 0
        
        ' Init tiles ddsd type and load the bitmap
        For i = 0 To ExtraSheets
            If Dir$(App.Path & "\GFX\tiles" & i & ".bmp") <> "" Then
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
        
        ' Init spells ddsd type and load the bitmap
        DDSD_BigSpellAnim.lFlags = DDSD_CAPS
        DDSD_BigSpellAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_BigSpellAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\bigspells.bmp", DDSD_BigSpellAnim)
        SetMaskColorFromPixel DD_BigSpellAnim, 0, 0
        
        ' Init arrows ddsd type and load the bitmap
        DDSD_ArrowAnim.lFlags = DDSD_CAPS
        DDSD_ArrowAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_ArrowAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\arrows.bmp", DDSD_ArrowAnim)
        SetMaskColorFromPixel DD_ArrowAnim, 0, 0
        
    Else
    
    'TILES FOR CUSTOM ENCRYPTED EXTENSIONS
    
        For i = 0 To ExtraSheets
            If Dir$(App.Path & "\GFX\tiles" & i & "." & Trim$(ENCRYPT_TYPE)) <> "" Then
                
                strfilename = App.Path & "/gfx/tiles" & i & "." & Trim$(ENCRYPT_TYPE)
                BMU.LoadByteData (strfilename)
                BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                BMU.DecompressByteData_ZLib
                DDSD_Tile(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
                DDSD_Tile(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
                DDSD_Tile(i).lWidth = BMU.ImageWidth
                DDSD_Tile(i).lHeight = BMU.ImageHeight
                Set DD_TileSurf(i) = DD.CreateSurface(DDSD_Tile(i))
                'Set DD_TileSurf(i) = DD.CreateSurfaceFromFile(strfilename, DDSD_Tile(i))
                DC = DD_TileSurf(i).GetDC
                
                Call BMU.Blt(DC)
                Call DD_TileSurf(i).ReleaseDC(DC)
                SetMaskColorFromPixel DD_TileSurf(i), 0, 0
                'DD_TileSurf(i).SetColorKey DDCKEY_SRCBLT, Key
             
                TileFile(i) = 1
            Else
                TileFile(i) = 0
            End If
        Next i
    
    'ITEMS FOR CUSTOM ENCRYPTED EXTENSIONS
        strfilename = App.Path & "/gfx/items." & Trim$(ENCRYPT_TYPE)
        BMU.LoadByteData (strfilename)
        BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
        BMU.DecompressByteData_ZLib
        DDSD_Item.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        DDSD_Item.lWidth = BMU.ImageWidth
        DDSD_Item.lHeight = BMU.ImageHeight
        Set DD_ItemSurf = DD.CreateSurface(DDSD_Item)
        'Set DD_ItemSurf = DD.CreateSurfaceFromFile(strfilename, DDSD_Item)
        DC = DD_ItemSurf.GetDC
        Call BMU.Blt(DC)
        Call DD_ItemSurf.ReleaseDC(DC)
        SetMaskColorFromPixel DD_ItemSurf, 0, 0
        'DD_ItemSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    'SPRITES FOR CUSTOM ENCRYPTED EXTENSIONS
        strfilename = App.Path & "/gfx/Sprites." & Trim$(ENCRYPT_TYPE)
        BMU.LoadByteData (strfilename)
        BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
        BMU.DecompressByteData_ZLib
        DDSD_Sprite.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        DDSD_Sprite.lWidth = BMU.ImageWidth
        DDSD_Sprite.lHeight = BMU.ImageHeight
        Set DD_SpriteSurf = DD.CreateSurface(DDSD_Sprite)
        'Set DD_SpriteSurf = DD.CreateSurfaceFromFile(strfilename, DDSD_Sprite)
        DC = DD_SpriteSurf.GetDC
        Call BMU.Blt(DC)
        Call DD_SpriteSurf.ReleaseDC(DC)
        SetMaskColorFromPixel DD_SpriteSurf, 0, 0
        'DD_SpriteSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    'BIG SPRITES FOR CUSTOM ENCRYPTED EXTENSIONS
        strfilename = App.Path & "/gfx/BigSprites." & Trim$(ENCRYPT_TYPE)
        BMU.LoadByteData (strfilename)
        BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
        BMU.DecompressByteData_ZLib
        DDSD_BigSprite.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_BigSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        DDSD_BigSprite.lWidth = BMU.ImageWidth
        DDSD_BigSprite.lHeight = BMU.ImageHeight
        Set DD_BigSpriteSurf = DD.CreateSurface(DDSD_BigSprite)
        'Set DD_BigSpriteSurf = DD.CreateSurfaceFromFile(strfilename, DDSD_BigSprite)
        DC = DD_BigSpriteSurf.GetDC
        Call BMU.Blt(DC)
        Call DD_BigSpriteSurf.ReleaseDC(DC)
        SetMaskColorFromPixel DD_BigSpriteSurf, 0, 0
        'DD_BigSpriteSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    'EMOTICONS FOR CUSTOM ENCRYPTED EXTENSIONS
        strfilename = App.Path & "/gfx/Emoticons." & Trim$(ENCRYPT_TYPE)
        BMU.LoadByteData (strfilename)
        BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
        BMU.DecompressByteData_ZLib
        DDSD_Emoticon.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_Emoticon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        DDSD_Emoticon.lWidth = BMU.ImageWidth
        DDSD_Emoticon.lHeight = BMU.ImageHeight
        Set DD_EmoticonSurf = DD.CreateSurface(DDSD_Emoticon)
        'Set DD_EmoticonSurf = DD.CreateSurfaceFromFile(strfilename, DDSD_Emoticon)
        DC = DD_EmoticonSurf.GetDC
        Call BMU.Blt(DC)
        Call DD_EmoticonSurf.ReleaseDC(DC)
        SetMaskColorFromPixel DD_EmoticonSurf, 0, 0
        'DD_EmoticonSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    'SPELLS FOR CUSTOM ENCRYPTED EXTENSIONS
        strfilename = App.Path & "/gfx/spells." & Trim$(ENCRYPT_TYPE)
        BMU.LoadByteData (strfilename)
        BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
        BMU.DecompressByteData_ZLib
        DDSD_SpellAnim.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_SpellAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        DDSD_SpellAnim.lWidth = BMU.ImageWidth
        DDSD_SpellAnim.lHeight = BMU.ImageHeight
        Set DD_SpellAnim = DD.CreateSurface(DDSD_SpellAnim)
        'Set DD_SpellAnim = DD.CreateSurfaceFromFile(strfilename, DDSD_SpellAnim)
        DC = DD_SpellAnim.GetDC
        Call BMU.Blt(DC)
        Call DD_SpellAnim.ReleaseDC(DC)
        SetMaskColorFromPixel DD_SpellAnim, 0, 0
        'DD_SpellAnim.SetColorKey DDCKEY_SRCBLT, Key
    
    'BIGSPELLS FOR CUSTOM ENCRYPTED EXTENSIONS
        strfilename = App.Path & "/gfx/bigspells." & Trim$(ENCRYPT_TYPE)
        BMU.LoadByteData (strfilename)
        BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
        BMU.DecompressByteData_ZLib
        DDSD_BigSpellAnim.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_BigSpellAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        DDSD_BigSpellAnim.lWidth = BMU.ImageWidth
        DDSD_BigSpellAnim.lHeight = BMU.ImageHeight
        Set DD_BigSpellAnim = DD.CreateSurface(DDSD_BigSpellAnim)
        'Set DD_BigSpellAnim = DD.CreateSurfaceFromFile(strfilename, DDSD_BigSpellAnim)
        DC = DD_BigSpellAnim.GetDC
        Call BMU.Blt(DC)
        Call DD_BigSpellAnim.ReleaseDC(DC)
        SetMaskColorFromPixel DD_BigSpellAnim, 0, 0
        'DD_BigSpellAnim.SetColorKey DDCKEY_SRCBLT, Key
    
    'ARROWS FOR CUSTOM ENCRYPTED EXTENSIONS
        strfilename = App.Path & "/gfx/arrows." & Trim$(ENCRYPT_TYPE)
        BMU.LoadByteData (strfilename)
        BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
        BMU.DecompressByteData_ZLib
        DDSD_ArrowAnim.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_ArrowAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        DDSD_ArrowAnim.lWidth = BMU.ImageWidth
        DDSD_ArrowAnim.lHeight = BMU.ImageHeight
        Set DD_ArrowAnim = DD.CreateSurface(DDSD_ArrowAnim)
        'Set DD_ArrowAnim = DD.CreateSurfaceFromFile(strfilename, DDSD_ArrowAnim)
        DC = DD_ArrowAnim.GetDC
        Call BMU.Blt(DC)
        Call DD_ArrowAnim.ReleaseDC(DC)
        SetMaskColorFromPixel DD_ArrowAnim, 0, 0
        'DD_ArrowAnim.SetColorKey DDCKEY_SRCBLT, Key
    
    End If

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
    Set DD_BigSpellAnim = Nothing
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

If TileFile(10) = 0 Then Exit Sub
    
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
                DisplayFx DD_TileSurf(10), (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31
            Else
                DisplayFx DD_TileSurf(10), (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Map(GetPlayerMap(MyIndex)).Tile(x, y).Light
            End If
        Next x
    Next y
End Sub
Sub WierdNight()
Dim x As Long, y As Long
Dim NewX As Long, NewY As Long
Dim NewX2 As Long, NewY2 As Long
Dim Tile As Long

If TileFile(10) = 0 Then Exit Sub
    
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
    If TileFile(10) = 1 Then
        rec.Top = Int(14 / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (14 - Int(14 / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
            
        For i = 1 To MAX_RAINDROPS
            If Not ((DropSnow(i).x = 0) Or (DropSnow(i).y = 0)) Then
                DropSnow(i).x = DropSnow(i).x + DropSnow(i).Speed
                DropSnow(i).y = DropSnow(i).y + DropSnow(i).Speed
                Call DD_BackBuffer.BltFast(DropSnow(i).x + DropSnow(i).Speed, DropSnow(i).y + DropSnow(i).Speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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

Sub BltSpell(ByVal Index As Long)
Dim x As Long, y As Long, i As Long, z As Long

If Player(Index).SpellNum <= 0 Or Player(Index).SpellNum > MAX_SPELLS Then Exit Sub


For i = 1 To MAX_SPELL_ANIM
'IF SPELL IS NOT BIG
If Spell(Player(Index).SpellNum).Big = 0 Then
    If Player(Index).SpellAnim(i).CastedSpell = YES Then
        If Player(Index).SpellAnim(i).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then
                        
        rec.Top = Spell(Player(Index).SpellNum).SpellAnim * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(Index).SpellAnim(i).SpellVar * PIC_X
        rec.Right = rec.Left + PIC_X
            
            If Player(Index).SpellAnim(i).TargetType = 0 Then
                
                'SMALL: IF TARGET IS A PLAYER
                If Player(Index).SpellAnim(i).Target > 0 Then
                
                    'SMALL: IF TARGET IS SELF
                    If Player(Index).SpellAnim(i).Target = MyIndex Then
                    x = NewX + sx
                    y = NewY + sx
                    Call DD_BackBuffer.BltFast(x, y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    
                    'SMALL: IF TARGET IS ANOTHER PLAYER
                    Else
                    x = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx + Player(Player(Index).SpellAnim(i).Target).XOffset
                    y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx + Player(Player(Index).SpellAnim(i).Target).YOffset
                    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    
                End If
            
            'SMALL: IF TARGET IS AN NPC
            Else
            
                'IF SMALL AND TARGET WILL DIE FROM THIS HIT, USE SCRIPT SPELL INSTEAD
                If MapNpc(Player(Index).SpellAnim(i).Target).HP <= Spell(Player(Index).SpellNum).Data1 = True Then
                
                    For z = 1 To MAX_SCRIPTSPELLS
                        If ScriptSpell(z).CastedSpell = NO Then
                            ScriptSpell(z).SpellNum = Player(Index).SpellNum
                            ScriptSpell(z).SpellDone = 0
                            ScriptSpell(z).SpellVar = 0
                            ScriptSpell(z).SpellTime = GetTickCount
                            ScriptSpell(z).x = MapNpc(Player(Index).SpellAnim(i).Target).x
                            ScriptSpell(z).y = MapNpc(Player(Index).SpellAnim(i).Target).y
                            ScriptSpell(z).CastedSpell = YES
                            Exit For
                        End If
                    Next z
                    
                 
                    'Player(Index).SpellAnim(i).SpellDone = Spell(Player(Index).SpellNum).SpellDone
                
                Else
                
                    x = MapNpc(Player(Index).SpellAnim(i).Target).x * PIC_X + sx + MapNpc(Player(Index).SpellAnim(i).Target).XOffset
                    y = MapNpc(Player(Index).SpellAnim(i).Target).y * PIC_Y + sx + MapNpc(Player(Index).SpellAnim(i).Target).YOffset
                    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                End If
            
            End If
        
        
        'SMALL: ADVANCE SPELL ONE CYCLE
                        
            If GetTickCount > Player(Index).SpellAnim(i).SpellTime + Spell(Player(Index).SpellNum).SpellTime Then
            Player(Index).SpellAnim(i).SpellTime = GetTickCount
            Player(Index).SpellAnim(i).SpellVar = Player(Index).SpellAnim(i).SpellVar + 1
            End If
        
            If Player(Index).SpellAnim(i).SpellVar > 12 Then
            Player(Index).SpellAnim(i).SpellDone = Player(Index).SpellAnim(i).SpellDone + 1
            Player(Index).SpellAnim(i).SpellVar = 0
            End If
        
        Else
        Player(Index).SpellAnim(i).CastedSpell = NO
        End If
    End If
Else
    If Player(Index).SpellAnim(i).CastedSpell = YES Then
        If Player(Index).SpellAnim(i).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then
                       
            rec.Top = Spell(Player(Index).SpellNum).SpellAnim * (PIC_Y * 3)
            rec.Bottom = rec.Top + PIC_Y + 64
            rec.Left = Player(Index).SpellAnim(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X + 64
           
            If Player(Index).SpellAnim(i).TargetType = 0 Then
            
                'BIG: IF TARGET IS A PLAYER
                If Player(Index).SpellAnim(i).Target > 0 Then
                
                    'BIG: IF TARGET IS SELF
                    If Player(Index).SpellAnim(i).Target = MyIndex Then
                        x = NewX + sx - 32
                        y = NewY + sx - 32
                        
                        If y < 0 Then
                        rec.Top = rec.Top + (y * -1)
                        y = 0
                        End If
        
                        If x < 0 Then
                        rec.Left = rec.Left + (x * -1)
                        x = 0
                        End If
        
                        If (x + 64) > (MAX_MAPX * 32) Then
                        rec.Right = rec.Left + 64
                        End If
        
                        If (y + 64) > (MAX_MAPY * 32) Then
                        rec.Bottom = rec.Top + 64
                        End If
                        
                        Call DD_BackBuffer.BltFast(x, y, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    
                    'BIG: IF TARGET IS A DIFFERENT PLAYER
                    Else
                        x = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx - 32 + Player(Player(Index).SpellAnim(i).Target).XOffset
                        y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx - 32 + Player(Player(Index).SpellAnim(i).Target).YOffset
                        
                        If y < 0 Then
                        rec.Top = rec.Top + (y * -1)
                        y = 0
                        End If
                        
                        If x < 0 Then
                        rec.Left = rec.Left + (x * -1)
                        x = 0
                        End If
                        
                        If (x + 64) > (MAX_MAPX * 32) Then
                        rec.Right = rec.Left + 64
                        End If
                        
                        If (y + 64) > (MAX_MAPY * 32) Then
                        rec.Bottom = rec.Top + 64
                        End If
                        
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            
            'BIG: IF TARGET IS AN NPC
            Else
            
                'IF BIG AND TARGET WILL DIE FROM THIS HIT, USE SCRIPT SPELL INSTEAD
                If MapNpc(Player(Index).SpellAnim(i).Target).HP <= Spell(Player(Index).SpellNum).Data1 Then
                
                    For z = 1 To MAX_SCRIPTSPELLS
                        If ScriptSpell(z).CastedSpell = NO Then
                            ScriptSpell(z).SpellNum = Player(Index).SpellNum
                            ScriptSpell(z).SpellDone = 0
                            ScriptSpell(z).SpellVar = 0
                            ScriptSpell(z).SpellTime = GetTickCount
                            ScriptSpell(z).x = MapNpc(Player(Index).SpellAnim(i).Target).x
                            ScriptSpell(z).y = MapNpc(Player(Index).SpellAnim(i).Target).y
                            ScriptSpell(z).CastedSpell = YES
                            Exit For
                        End If
                    Next z
                
                    Player(Index).SpellAnim(i).SpellDone = Spell(Player(Index).SpellNum).SpellDone
                
                Else
            
            
                x = MapNpc(Player(Index).SpellAnim(i).Target).x * PIC_X + sx - 32 + MapNpc(Player(Index).SpellAnim(i).Target).XOffset
                y = MapNpc(Player(Index).SpellAnim(i).Target).y * PIC_Y + sx - 32 + MapNpc(Player(Index).SpellAnim(i).Target).YOffset
                
                        If y < 0 Then
                        rec.Top = rec.Top + (y * -1)
                        y = 0
                        End If
                        
                        If x < 0 Then
                        rec.Left = rec.Left + (x * -1)
                        x = 0
                        End If
                        
                        If (x + 64) > (MAX_MAPX * 32) Then
                        rec.Right = rec.Left + 64
                        End If
                        
                        If (y + 64) > (MAX_MAPY * 32) Then
                        rec.Bottom = rec.Top + 64
                        End If
                
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            End If
            
            'BIG: ADVANCE SPELL ONE CYCLE
            If GetTickCount > Player(Index).SpellAnim(i).SpellTime + Spell(Player(Index).SpellNum).SpellTime Then
            Player(Index).SpellAnim(i).SpellTime = GetTickCount
            Player(Index).SpellAnim(i).SpellVar = Player(Index).SpellAnim(i).SpellVar + 3
            End If
            
            If Player(Index).SpellAnim(i).SpellVar > 36 Then
            Player(Index).SpellAnim(i).SpellDone = Player(Index).SpellAnim(i).SpellDone + 1
            Player(Index).SpellAnim(i).SpellVar = 0
            End If
            
        Else
            Player(Index).SpellAnim(i).CastedSpell = NO
        End If
    End If
End If
Next i
End Sub


Sub BltSpell2()
Dim x As Long, y As Long, i As Long

For i = 1 To MAX_SCRIPTSPELLS
x = ScriptSpell(i).x
y = ScriptSpell(i).y
If ScriptSpell(i).SpellNum > 0 And ScriptSpell(i).SpellNum <= MAX_SPELLS Then
If Spell(ScriptSpell(i).SpellNum).Big = 0 Then
    If ScriptSpell(i).CastedSpell = YES Then
        
        If ScriptSpell(i).SpellDone < Spell(ScriptSpell(i).SpellNum).SpellDone Then
                        
            rec.Top = Spell(ScriptSpell(i).SpellNum).SpellAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = ScriptSpell(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X
    
            'x = NewX * PIC_X
            'y = NewY * PIC_Y

            y = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            x = (x - NewPlayerX) * PIC_X + sx - NewXOffset
                        
            If ScriptSpell(i).SpellVar > 10 Then
                ScriptSpell(i).SpellDone = ScriptSpell(i).SpellDone + 1
                ScriptSpell(i).SpellVar = 0
            End If

            If GetTickCount > ScriptSpell(i).SpellTime + Spell(ScriptSpell(i).SpellNum).SpellTime Then
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).SpellVar = ScriptSpell(i).SpellVar + 1
            End If
                        
            Call DD_BackBuffer.BltFast(x, y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            
        Else
            ScriptSpell(i).CastedSpell = NO
        End If
        
    End If
    
Else

    If ScriptSpell(i).CastedSpell = YES Then
        If ScriptSpell(i).SpellDone < Spell(ScriptSpell(i).SpellNum).SpellDone Then


                       
            rec.Top = Spell(ScriptSpell(i).SpellNum).SpellAnim * (PIC_Y * 3)
            rec.Bottom = rec.Top + PIC_Y + 64
            rec.Left = ScriptSpell(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X + 64
           

                        'x = x * PIC_X - 32
                        'y = y * PIC_Y - 32
                        
                        y = (y - NewPlayerY) * PIC_Y + sx - NewYOffset - PIC_Y
                        x = (x - NewPlayerX) * PIC_X + sx - NewXOffset - PIC_X
                        
                        
        If y < 0 Then
        rec.Top = rec.Top + (y * -1)
        y = 0
        End If
        
        If x < 0 Then
        rec.Left = rec.Left + (x * -1)
        x = 0
        End If
        
        If (x + 64) > (MAX_MAPX * 32) Then
        rec.Right = rec.Left + 64
        End If
        
        If (y + 64) > (MAX_MAPY * 32) Then
        rec.Bottom = rec.Top + 64
        End If
        
                    If ScriptSpell(i).SpellVar > 30 Then
            ScriptSpell(i).SpellDone = ScriptSpell(i).SpellDone + 1
            ScriptSpell(i).SpellVar = 0
            End If

            If GetTickCount > ScriptSpell(i).SpellTime + Spell(ScriptSpell(i).SpellNum).SpellTime Then
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).SpellVar = ScriptSpell(i).SpellVar + 3
            End If
                       
                        Call DD_BackBuffer.BltFast(x, y, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            ScriptSpell(i).CastedSpell = NO
        End If
    End If
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
        
        rec.Top = Player(Index).EmoticonNum * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(Index).EmoticonVar * PIC_X
        rec.Right = rec.Left + PIC_X
        
        If Index = MyIndex Then
            x2 = NewX + sx + 16
            y2 = NewY + sx - 32
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            x2 = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + 16
            y2 = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - 32
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltArrow(ByVal Index As Long)
Dim x As Long, y As Long, i As Long, z As Long
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
            x = Player(Index).Arrow(z).ArrowX
            y = Player(Index).Arrow(z).ArrowY + Int(Player(Index).Arrow(z).ArrowVarY / 32)
            If y > Player(Index).Arrow(z).ArrowY + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                Player(Index).Arrow(z).Arrow = 0
            End If
            
            If y <= MAX_MAPY Then
                Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If Player(Index).Arrow(z).ArrowPosition = 1 Then
            x = Player(Index).Arrow(z).ArrowX
            y = Player(Index).Arrow(z).ArrowY - Int(Player(Index).Arrow(z).ArrowVarY / 32)
            If y < Player(Index).Arrow(z).ArrowY - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                Player(Index).Arrow(z).Arrow = 0
            End If
            
            If y >= 0 Then
                Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If Player(Index).Arrow(z).ArrowPosition = 2 Then
            x = Player(Index).Arrow(z).ArrowX + Int(Player(Index).Arrow(z).ArrowVarX / 32)
            y = Player(Index).Arrow(z).ArrowY
            If x > Player(Index).Arrow(z).ArrowX + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                Player(Index).Arrow(z).Arrow = 0
            End If
            
            If x <= MAX_MAPX Then
                Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If Player(Index).Arrow(z).ArrowPosition = 3 Then
            x = Player(Index).Arrow(z).ArrowX - Int(Player(Index).Arrow(z).ArrowVarX / 32)
            y = Player(Index).Arrow(z).ArrowY
            If x < Player(Index).Arrow(z).ArrowX - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                Player(Index).Arrow(z).Arrow = 0
            End If
            
            If x >= 0 Then
             Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If x >= 0 And x <= MAX_MAPX Then
            If y >= 0 And y <= MAX_MAPY Then
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Player(Index).Arrow(z).Arrow = 0
                End If
            End If
        End If
        
        For i = 1 To MAX_PLAYERS
           If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(i) = x And GetPlayerY(i) = y Then
                    If Index = MyIndex Then
                        Call SendData("arrowhit" & SEP_CHAR & 0 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                    End If
                    If Index <> i Then Player(Index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(i).num > 0 Then
                If MapNpc(i).x = x And MapNpc(i).y = y Then
                    If Index = MyIndex Then
                        Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                    End If
                    Player(Index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next i
        
        For BX = 0 To MAX_MAPX
            For BY = 0 To MAX_MAPY
                If Map(GetPlayerMap(MyIndex)).Tile(BX, BY).Type = TILE_TYPE_NPC_SPAWN Then
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        If MapAttributeNpc(i, BX, BY).x = x And MapAttributeNpc(i, BX, BY).y = y Then
                            If Index = MyIndex Then
                                Call SendData("arrowhit" & SEP_CHAR & 2 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & BX & SEP_CHAR & BY & SEP_CHAR & END_CHAR)
                            End If
                            Player(Index).Arrow(z).Arrow = 0
                            Exit Sub
                        End If
                    Next i
                End If
            Next BY
        Next BX
    End If
Next z
End Sub
