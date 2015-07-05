Attribute VB_Name = "modDirectX"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Private Resp As Long

Public Const TilesInSheets As Byte = 14
Public Const ExtraSheets As Byte = 6

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
    Set DD = DX.DirectDrawCreate(vbNullString)
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

Sub InitDirectSM()

    ' DirectSound7
    Set Ds = DX.DirectSoundCreate(vbNullString)
    Call Ds.SetCooperativeLevel(frmMirage.hWnd, DSSCL_NORMAL)
    
    ' DirectMusic7
    Set loader = DX.DirectMusicLoaderCreate()
    Set perf = DX.DirectMusicPerformanceCreate()
    Call perf.Init(Nothing, 0)
    perf.SetPort -1, 80
    Call perf.SetMasterAutoDownload(True)
    perf.SetMasterVolume (MusicVolume * 42 - 3000)

End Sub

Sub InitSurfaces()
Dim key As DDCOLORKEY
Dim I As Long

    ' Check for files existing
    If FileExist("\GFX\sprites.bmp") = False Or FileExist("\GFX\items.bmp") = False Or FileExist("\GFX\bigsprites.bmp") = False Or FileExist("\GFX\emoticons.bmp") = False Or FileExist("\GFX\arrows.bmp") = False Then
        Call MsgBox("You're missing some graphic files!", vbOKOnly, GAME_NAME)
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
    For I = 0 To ExtraSheets
        If Dir(App.Path & "\GFX\tiles" & I & ".bmp") <> vbNullString Then
            DDSD_Tile(I).lFlags = DDSD_CAPS
            DDSD_Tile(I).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_TileSurf(I) = DD.CreateSurfaceFromFile(App.Path & "\GFX\tiles" & I & ".bmp", DDSD_Tile(I))
            SetMaskColorFromPixel DD_TileSurf(I), 0, 0
            TileFile(I) = 1
        Else
            TileFile(I) = 0
        End If
    Next I
    
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
Dim I As Long

    Set DX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    For I = 0 To ExtraSheets
        If TileFile(I) = 1 Then
            Set DD_TileSurf(I) = Nothing
        End If
    Next I
    Set DD_ItemSurf = Nothing
    Set DD_BigSpriteSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    Set DD_SpellAnim = Nothing
    Set DD_ArrowAnim = Nothing
    If Not (perf Is Nothing) Then perf.CloseDown
    Set Ds = Nothing
    Set loader = Nothing
    Set seg = Nothing
    Set dsbuffer = Nothing
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

Sub BltWeather()
Dim I As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))
    
    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For I = 1 To MAX_RAINDROPS
            If DropRain(I).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = vbNullString Then
                        frmMirage.tmrRainDrop.Interval = 200
                        frmMirage.tmrRainDrop.Tag = "123"
                    End If
                End If
            End If
        Next I
    ElseIf GameWeather = WEATHER_SNOWING Then
        For I = 1 To MAX_RAINDROPS
            If DropSnow(I).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                    If frmMirage.tmrSnowDrop.Tag = vbNullString Then
                        frmMirage.tmrSnowDrop.Interval = 200
                        frmMirage.tmrSnowDrop.Tag = "123"
                    End If
                End If
            End If
        Next I
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= RainIntensity Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If
    
    For I = 1 To MAX_RAINDROPS
        If Not ((DropRain(I).x = 0) Or (DropRain(I).y = 0)) Then
            DropRain(I).x = DropRain(I).x + DropRain(I).Speed
            DropRain(I).y = DropRain(I).y + DropRain(I).Speed
            Call DD_BackBuffer.DrawLine(DropRain(I).x, DropRain(I).y, DropRain(I).x + DropRain(I).Speed, DropRain(I).y + DropRain(I).Speed)
            If (DropRain(I).x > (MAX_MAPX + 1) * PIC_X) Or (DropRain(I).y > (MAX_MAPY + 1) * PIC_Y) Then
                DropRain(I).Randomized = False
            End If
        End If
    Next I
    If TileFile(6) = 1 Then
        rec.Top = Int(14 / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (14 - Int(14 / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
            
        For I = 1 To MAX_RAINDROPS
            If Not ((DropSnow(I).x = 0) Or (DropSnow(I).y = 0)) Then
                DropSnow(I).x = DropSnow(I).x + DropSnow(I).Speed
                DropSnow(I).y = DropSnow(I).y + DropSnow(I).Speed
                Call DD_BackBuffer.BltFast(DropSnow(I).x + DropSnow(I).Speed, DropSnow(I).y + DropSnow(I).Speed, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(I).x > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(I).y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropSnow(I).Randomized = False
                End If
            End If
        Next I
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
Dim x As Long, y As Long, I As Long

If Player(Index).SpellNum <= 0 Or Player(Index).SpellNum > MAX_SPELLS Then Exit Sub
If Spell(Player(Index).SpellNum).SpellAnim <= 0 Then Exit Sub

For I = 1 To MAX_SPELL_ANIM
    If Player(Index).SpellAnim(I).CastedSpell = YES Then
        If Player(Index).SpellAnim(I).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then
            If Player(Index).SpellAnim(I).SpellVar > 10 Then
                Player(Index).SpellAnim(I).SpellDone = Player(Index).SpellAnim(I).SpellDone + 1
                Player(Index).SpellAnim(I).SpellVar = 0
            End If
            If GetTickCount > Player(Index).SpellAnim(I).SpellTime + Spell(Player(Index).SpellNum).SpellTime Then
                Player(Index).SpellAnim(I).SpellTime = GetTickCount
                Player(Index).SpellAnim(I).SpellVar = Player(Index).SpellAnim(I).SpellVar + 1
            End If
                        
            rec.Top = Spell(Player(Index).SpellNum).SpellAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = Player(Index).SpellAnim(I).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X
            
            If Player(Index).SpellAnim(I).TargetType = TARGET_TYPE_PLAYER Then
                If Player(Index).SpellAnim(I).Target > 0 Then
                    If Player(Index).SpellAnim(I).Target = MyIndex Then
                        x = NewX + sx
                        y = NewY + sx
                        Call DD_BackBuffer.BltFast(x, y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        x = GetPlayerX(Player(Index).SpellAnim(I).Target) * PIC_X + sx + Player(Player(Index).SpellAnim(I).Target).XOffset
                        y = GetPlayerY(Player(Index).SpellAnim(I).Target) * PIC_Y + sx + Player(Player(Index).SpellAnim(I).Target).YOffset
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            Else
                If Player(Index).SpellAnim(I).TargetType = TARGET_TYPE_NPC Then
                    x = MapNpc(Player(Index).SpellAnim(I).Target).x * PIC_X + sx + MapNpc(Player(Index).SpellAnim(I).Target).XOffset
                    y = MapNpc(Player(Index).SpellAnim(I).Target).y * PIC_Y + sx + MapNpc(Player(Index).SpellAnim(I).Target).YOffset
                    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    If Player(Index).SpellAnim(I).TargetType = TARGET_TYPE_LOCATION Then
                        x = MakeX(Player(Index).SpellAnim(I).Target) * PIC_X + sx
                        y = MakeY(Player(Index).SpellAnim(I).Target) * PIC_Y + sx
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            End If
        Else
            Player(Index).SpellAnim(I).CastedSpell = NO
        End If
    End If
Next I
End Sub

Sub BltEmoticons(ByVal Index As Long)
Dim x2 As Long, y2 As Long
Dim ETime As Long
ETime = 1300
   
    If Player(Index).EmoticonNum < 0 And Player(Index).EmoticonPlayed Then Exit Sub
    
    If (Player(Index).EmoticonType = EMOTICON_TYPE_SOUND Or Player(Index).EmoticonType = EMOTICON_TYPE_BOTH) And Player(Index).EmoticonPlayed = False And EmoticonSoundOn = YES Then
        Call PlaySound(Player(Index).EmoticonSound)
        Player(Index).EmoticonPlayed = True
    End If
    
    If Player(Index).EmoticonType = EMOTICON_TYPE_IMAGE Or Player(Index).EmoticonType = EMOTICON_TYPE_BOTH Then
        
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
    End If
End Sub

Sub BltArrow(ByVal Index As Long) ' The MAP index, not the player index
Dim x As Long, y As Long, I As Long, z As Long
Dim BX As Long, BY As Long

For z = 1 To MAX_PLAYER_ARROWS

    If Map(Index).Arrow(z).Arrow > 0 Then
   
        rec.Top = Map(Index).Arrow(z).ArrowAnim * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Map(Index).Arrow(z).ArrowPosition * PIC_X
        rec.Right = rec.Left + PIC_X
       
        If GetTickCount > Map(Index).Arrow(z).ArrowTime + 30 Then
            Map(Index).Arrow(z).ArrowTime = GetTickCount
            Map(Index).Arrow(z).ArrowVarX = Map(Index).Arrow(z).ArrowVarX + 10
            Map(Index).Arrow(z).ArrowVarY = Map(Index).Arrow(z).ArrowVarY + 10
        End If
       
        If Map(Index).Arrow(z).ArrowPosition = 0 Then
            x = Map(Index).Arrow(z).ArrowX
            y = Map(Index).Arrow(z).ArrowY + Int(Map(Index).Arrow(z).ArrowVarY / 32)
            If y > Map(Index).Arrow(z).ArrowY + Arrows(Map(Index).Arrow(z).ArrowNum).Range - 2 Then
                Map(Index).Arrow(z).Arrow = 0
            End If
           
            If y <= MAX_MAPY Then
                Call DD_BackBuffer.BltFast((Map(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Map(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Map(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
       
        If Map(Index).Arrow(z).ArrowPosition = 1 Then
            x = Map(Index).Arrow(z).ArrowX
            y = Map(Index).Arrow(z).ArrowY - Int(Map(Index).Arrow(z).ArrowVarY / 32)
            If y < Map(Index).Arrow(z).ArrowY - Arrows(Map(Index).Arrow(z).ArrowNum).Range + 2 Then
                Map(Index).Arrow(z).Arrow = 0
            End If
           
            If y >= 0 Then
                Call DD_BackBuffer.BltFast((Map(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Map(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Map(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
       
        If Map(Index).Arrow(z).ArrowPosition = 2 Then
            x = Map(Index).Arrow(z).ArrowX + Int(Map(Index).Arrow(z).ArrowVarX / 32)
            y = Map(Index).Arrow(z).ArrowY
            If x > Map(Index).Arrow(z).ArrowX + Arrows(Map(Index).Arrow(z).ArrowNum).Range - 2 Then
                Map(Index).Arrow(z).Arrow = 0
            End If
           
            If x <= MAX_MAPX Then
                Call DD_BackBuffer.BltFast((Map(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Map(Index).Arrow(z).ArrowVarX, (Map(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
       
        If Map(Index).Arrow(z).ArrowPosition = 3 Then
            x = Map(Index).Arrow(z).ArrowX - Int(Map(Index).Arrow(z).ArrowVarX / 32)
            y = Map(Index).Arrow(z).ArrowY
            If x < Map(Index).Arrow(z).ArrowX - Arrows(Map(Index).Arrow(z).ArrowNum).Range + 2 Then
                Map(Index).Arrow(z).Arrow = 0
            End If
           
            If x >= 0 Then
             Call DD_BackBuffer.BltFast((Map(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Map(Index).Arrow(z).ArrowVarX, (Map(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
       
        If x >= 0 And x <= MAX_MAPX Then
            If y >= 0 And y <= MAX_MAPY Then
                If Map(Index).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Map(Index).Arrow(z).Arrow = 0
                End If
            End If
        End If
       
        For I = 1 To MAX_PLAYERS
           If IsPlaying(I) And GetPlayerMap(I) = Index Then
                If GetPlayerX(I) = x And GetPlayerY(I) = y Then
                    If Map(Index).Arrow(z).ArrowOwner = MyIndex Then
                        Call SendData(ARROWHIT_CHAR & SEP_CHAR & 0 & SEP_CHAR & I & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
                    End If
                    If Map(Index).Arrow(z).ArrowOwner <> I Then Map(Index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next I
       
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(I).Num > 0 Then
                If MapNpc(I).x = x And MapNpc(I).y = y Then
                    If Map(Index).Arrow(z).ArrowOwner = MyIndex Then
                        Call SendData(ARROWHIT_CHAR & SEP_CHAR & 1 & SEP_CHAR & I & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
                    End If
                    Map(Index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next I
    End If
Next z
End Sub
