Attribute VB_Name = "modDirectX"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit
Public SwitchChar As Byte
Public SwitchServ As Byte
Private Resp As Long
Public Const TilesInSheets = 14
Public Const ExtraSheets = 8
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
Public DDSD_MiniMap As DDSURFACEDESC2
Public DD_MiniMap As DirectDrawSurface7
Public DD_Icon As DirectDrawSurface7
Public DDSD_Icon As DDSURFACEDESC2
Public DD_CorpseAnim As DirectDrawSurface7
Public DDSD_CorpseAnim As DDSURFACEDESC2
Public DD_MiniMapBuffer As DirectDrawSurface7
Public DD_MiniMapBuffer2 As DirectDrawSurface7
Public DD_MiniMapBuffer3 As DirectDrawSurface7
Public DD_MiniMapBuffer4 As DirectDrawSurface7
Public DD_MiniMapBuffer5 As DirectDrawSurface7
Public DD_MiniMapBuffer6 As DirectDrawSurface7
Public DD_MiniMapBuffer7 As DirectDrawSurface7
Public DD_MiniMapBuffer8 As DirectDrawSurface7
Public DD_MiniMapBuffer9 As DirectDrawSurface7
Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX(Optional mmStart As Boolean = False)
    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
 
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hWnd, DDSCL_NORMAL)
        
    If mmStart = False Then
        If windowed() Then
            frmMirage.WindowState = 0
            mclsStyle.Titlebar = True
            frmMirage.Height = frmMirage.Height
        Else
            DD.SetDisplayMode 800, 600, 16, 0, DDSDM_DEFAULT
            mclsStyle.Titlebar = False
        End If
        
        frmMirage.Show
        frmMirage.SetFocus
        
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
    End If
    ' Initialize all surfaces
    Call InitSurfaces(mmStart)
End Sub

Private Sub InitSurface(ByVal filename As String, ByVal StrPW As String, ByRef DDSD As DDSURFACEDESC2, ByRef DDSurf As DirectDrawSurface7)
Dim sDc As Long
Dim BMU As BitmapUtils

    Set BMU = New BitmapUtils
    
    With BMU
        Call .LoadByteData(filename)
        Call .DecryptByteData(StrPW)
        Call .DecompressByteData      'If you want to use zlib, you can change this to .DecompressByteData_ZLib
    End With
    
    DDSD.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD.lWidth = BMU.ImageWidth
    DDSD.lHeight = BMU.ImageHeight
    
    Set DDSurf = DD.CreateSurface(DDSD)
    
    sDc = DDSurf.GetDC
    Call BMU.Blt(sDc)
    Call DDSurf.ReleaseDC(sDc)
    Call SetMaskColorFromPixel(DDSurf, 0, 0)
    
    Set BMU = Nothing
End Sub

Sub InitSurfaces(Optional mmStart As Boolean = False)
Dim Key As DDCOLORKEY
Dim I As Long
Dim StrPW As String

    ' Our password is going to be 'test'.. We set it with Chr$(num)
    ' t = 116, e = 101, s = 115, t = 116
    ' Doing it this way keeps someone from hex editing your password easily.
    ' You can also just do: StrPW = "test" The result would be the same.
    StrPW = GFX_PASSWORD
    
    ' Check for files existing
    If FileExist("\Main\GFX\sprites.gfx") = False Or FileExist("\Main\GFX\items.gfx") = False Or FileExist("\Main\GFX\bigsprites.gfx") = False Or FileExist("\Main\GFX\emoticons.gfx") = False Or FileExist("\Main\GFX\arrows.gfx") = False Or FileExist("\Main\GFX\minimap.gfx") = False Then
        Call MsgBox("Your missing some graphic files!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If
    
    ' Set the key for masks
    Key.low = 0
    Key.high = 0
    
    If Not mmStart Then
        ' Initialize back buffer
        DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
        DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y
        Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
        Set DD_MiniMapBuffer = DD.CreateSurface(DDSD_BackBuffer)
        Set DD_MiniMapBuffer2 = DD.CreateSurface(DDSD_BackBuffer)
        Set DD_MiniMapBuffer3 = DD.CreateSurface(DDSD_BackBuffer)
        Set DD_MiniMapBuffer4 = DD.CreateSurface(DDSD_BackBuffer)
        Set DD_MiniMapBuffer5 = DD.CreateSurface(DDSD_BackBuffer)
        Set DD_MiniMapBuffer6 = DD.CreateSurface(DDSD_BackBuffer)
        Set DD_MiniMapBuffer7 = DD.CreateSurface(DDSD_BackBuffer)
        Set DD_MiniMapBuffer8 = DD.CreateSurface(DDSD_BackBuffer)
        Set DD_MiniMapBuffer9 = DD.CreateSurface(DDSD_BackBuffer)
        
        DD_MiniMapBuffer.SetColorKey DDCKEY_SRCBLT, Key
        DD_MiniMapBuffer2.SetColorKey DDCKEY_SRCBLT, Key
        DD_MiniMapBuffer3.SetColorKey DDCKEY_SRCBLT, Key
        DD_MiniMapBuffer4.SetColorKey DDCKEY_SRCBLT, Key
        DD_MiniMapBuffer5.SetColorKey DDCKEY_SRCBLT, Key
        DD_MiniMapBuffer5.SetColorKey DDCKEY_SRCBLT, Key
        DD_MiniMapBuffer6.SetColorKey DDCKEY_SRCBLT, Key
        DD_MiniMapBuffer7.SetColorKey DDCKEY_SRCBLT, Key
        DD_MiniMapBuffer8.SetColorKey DDCKEY_SRCBLT, Key
        DD_MiniMapBuffer9.SetColorKey DDCKEY_SRCBLT, Key
    
      ' Init tiles ddsd type and load the bitmap
      For I = 0 To ExtraSheets
          If Dir(App.Path & "\Main\GFX\tiles" & I & ".gfx") <> "" Then
              Call InitSurface(App.Path & "\Main\GFX\tiles" & I & ".gfx", StrPW, DDSD_Tile(I), DD_TileSurf(I))
              TileFile(I) = 1
          Else
              TileFile(I) = 0
          End If
      Next I
    
      ' Init big sprites ddsd type and load the bitmap
      Call InitSurface(App.Path & "\Main\GFX\bigsprites.gfx", StrPW, DDSD_BigSprite, DD_BigSpriteSurf)
      
      ' Init emoticons ddsd type and load the bitmap
      Call InitSurface(App.Path & "\Main\GFX\emoticons.gfx", StrPW, DDSD_Emoticon, DD_EmoticonSurf)
      
      ' Init spells ddsd type and load the bitmap
      Call InitSurface(App.Path & "\Main\GFX\spells.gfx", StrPW, DDSD_SpellAnim, DD_SpellAnim)
      
      ' Init arrows ddsd type and load the bitmap
      Call InitSurface(App.Path & "\Main\GFX\arrows.gfx", StrPW, DDSD_ArrowAnim, DD_ArrowAnim)
    End If
    
    ' Init sprite ddsd type and load the bitmap
    Call InitSurface(App.Path & "\Main\GFX\sprites.gfx", StrPW, DDSD_Sprite, DD_SpriteSurf)
    
    ' Init items ddsd type and load the bitmap
    Call InitSurface(App.Path & "\Main\GFX\items.gfx", StrPW, DDSD_Item, DD_ItemSurf)

    ' Init items ddsd type and load the bitmap
      Call InitSurface(App.Path & "\Main\GFX\minimap.gfx", StrPW, DDSD_MiniMap, DD_MiniMap)
      
      ' Init Spell Icons ddsd type and load the bitmap
      Call InitSurface(App.Path & "\Main\GFX\spellicons.gfx", StrPW, DDSD_Icon, DD_Icon)

      ' Init Corpses ddsd type and load the bitmap
      Call InitSurface(App.Path & "\Main\GFX\corpse.gfx", StrPW, DDSD_CorpseAnim, DD_CorpseAnim)
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
    Set DD_MiniMap = Nothing
    Set DD_Icon = Nothing
    Set DD_CorpseAnim = Nothing
    Set DD_MiniMapBuffer = Nothing
    Set DD_MiniMapBuffer2 = Nothing
    Set DD_MiniMapBuffer3 = Nothing
    Set DD_MiniMapBuffer4 = Nothing
    Set DD_MiniMapBuffer5 = Nothing
    Set DD_MiniMapBuffer6 = Nothing
    Set DD_MiniMapBuffer7 = Nothing
    Set DD_MiniMapBuffer8 = Nothing
    Set DD_MiniMapBuffer9 = Nothing
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
Dim X As Long, Y As Long
Dim NewX As Long, NewY As Long
Dim NewX2 As Long, NewY2 As Long
Dim Tile As Long

    If TileFile(6) = 0 Then Exit Sub
    
     If Val(GetVar(App.Path & "\Main\Config\config.ini", "CONFIG", "Night")) = 1 Then
    NewX = GetPlayerX(MyIndex) - roundUp(SCREEN_X / 2) - 1
    NewY = GetPlayerY(MyIndex) - roundUp(SCREEN_Y / 2) - 1
    
    NewX2 = GetPlayerX(MyIndex) + roundUp(SCREEN_X / 2) + 1
    NewY2 = GetPlayerY(MyIndex) + roundUp(SCREEN_Y / 2) + 1
    
    If NewX < 0 Then
        NewX = 0
        NewX2 = SCREEN_X + 1
    End If
    If NewX2 > MAX_MAPX Then
        NewX = MAX_MAPX - SCREEN_X - 2
        NewX2 = MAX_MAPX
    End If
    
    If NewY < 0 Then
        NewY = 0
        NewY2 = SCREEN_Y + 1
    End If
    If NewY2 > MAX_MAPY Then
        NewY = MAX_MAPY - SCREEN_Y - 2
        NewY2 = MAX_MAPY
    End If

    If MAX_MAPX = SCREEN_X And MAX_MAPY = SCREEN_Y Then
        NewY = 0
        NewY2 = SCREEN_Y
        NewX = 0
        NewX2 = SCREEN_X
    End If
        
    For Y = NewY To NewY2
        For X = NewX To NewX2
            If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light <= 0 Then
                DisplayFx DD_TileSurf(6), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31
            Else
                DisplayFx DD_TileSurf(6), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light
            End If
        Next X
    Next Y
End If
End Sub

Sub BltWeather()
Dim I As Long


    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))
    
 If Val(GetVar(App.Path & "\Main\Config\config.ini", "CONFIG", "Weather")) = 1 Then
   
    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For I = 1 To MAX_RAINDROPS
            If DropRain(I).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = "" Then
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
                    If frmMirage.tmrSnowDrop.Tag = "" Then
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
        frmMirage.tmrRainDrop.Tag = ""
    End If
    
    For I = 1 To MAX_RAINDROPS
        If Not ((DropRain(I).X = 0) Or (DropRain(I).Y = 0)) Then
            DropRain(I).X = DropRain(I).X + DropRain(I).speed
            DropRain(I).Y = DropRain(I).Y + DropRain(I).speed
            Call DD_BackBuffer.DrawLine(DropRain(I).X, DropRain(I).Y, DropRain(I).X + DropRain(I).speed, DropRain(I).Y + DropRain(I).speed)
            If (DropRain(I).X > (MAX_MAPX + 1) * PIC_X) Or (DropRain(I).Y > (MAX_MAPY + 1) * PIC_Y) Then
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
            If Not ((DropSnow(I).X = 0) Or (DropSnow(I).Y = 0)) Then
                DropSnow(I).X = DropSnow(I).X + DropSnow(I).speed
                DropSnow(I).Y = DropSnow(I).Y + DropSnow(I).speed
                Call DD_BackBuffer.BltFast(DropSnow(I).X + DropSnow(I).speed, DropSnow(I).Y + DropSnow(I).speed, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(I).X > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(I).Y > (MAX_MAPY + 1) * PIC_Y) Then
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
End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).Y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).Y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).X > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropRain(RDNumber).speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropRain(RDNumber).X = 0
    DropRain(RDNumber).Y = 0
    DropRain(RDNumber).speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropSnow(RDNumber).Y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropSnow(RDNumber).Y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropSnow(RDNumber).X > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropSnow(RDNumber).speed = Int((10 * Rnd) + 6)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropSnow(RDNumber).X = 0
    DropSnow(RDNumber).Y = 0
    DropSnow(RDNumber).speed = 0
    DropSnow(RDNumber).Randomized = False
End Sub

Sub BltSpell(ByVal Index As Long)
Dim X As Long, Y As Long, I As Long

If Player(Index).SpellNum <= 0 Or Player(Index).SpellNum > MAX_SPELLS Then Exit Sub
If Spell(Player(Index).SpellNum).SpellAnim <= 0 Then Exit Sub

For I = 1 To MAX_SPELL_ANIM
    If Player(Index).SpellAnim(I).CastedSpell = YES Then
        If Player(Index).SpellAnim(I).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then
            If Player(Index).SpellAnim(I).SpellVar > 43 Then
                Player(Index).SpellAnim(I).SpellDone = Player(Index).SpellAnim(I).SpellDone + 1
                Player(Index).SpellAnim(I).SpellVar = 0
            End If
            If GetTickCount > Player(Index).SpellAnim(I).SpellTime + Spell(Player(Index).SpellNum).SpellTime Then
                Player(Index).SpellAnim(I).SpellTime = GetTickCount
                Player(Index).SpellAnim(I).SpellVar = Player(Index).SpellAnim(I).SpellVar + 3
            End If
                        
            rec.Top = Spell(Player(Index).SpellNum).SpellAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y + 64
            rec.Left = Player(Index).SpellAnim(I).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X + 64
            
            If Player(Index).SpellAnim(I).TargetType = 0 Then
                If Player(Index).SpellAnim(I).Target > 0 Then
                    If Player(Index).SpellAnim(I).Target = MyIndex Then
                        X = NewX '- 32
                        Y = NewY '- 32
                        Call DD_BackBuffer.BltFast(X, Y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        X = GetPlayerX(Player(Index).SpellAnim(I).Target) * PIC_X - 0 + Player(Player(Index).SpellAnim(I).Target).XOffset
                        Y = GetPlayerY(Player(Index).SpellAnim(I).Target) * PIC_Y - 0 + Player(Player(Index).SpellAnim(I).Target).YOffset
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            Else
                X = MapNpc(Player(Index).SpellAnim(I).Target).X * PIC_X - 0 + MapNpc(Player(Index).SpellAnim(I).Target).XOffset
                Y = MapNpc(Player(Index).SpellAnim(I).Target).Y * PIC_Y - 0 + MapNpc(Player(Index).SpellAnim(I).Target).YOffset
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
    
    If (Player(Index).EmoticonType = EMOTICON_TYPE_SOUND Or Player(Index).EmoticonType = EMOTICON_TYPE_BOTH) And Player(Index).EmoticonPlayed = False And Val(GetVar("Main\Config\config.ini", "CONFIG", "EmoticonSound")) = 1 Then
        Call PlaySound(Player(Index).EmoticonSound)
        Player(Index).EmoticonPlayed = True
    End If
    
    If Val(GetVar(App.Path & "\Main\Config\config.ini", "CONFIG", "Emoticons")) = 1 Then
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
End If
End Sub

Sub BltArrow(ByVal Index As Long)
Dim X As Long, Y As Long, I As Long, z As Long
Dim BX As Long, BY As Long

For z = 1 To MAX_PLAYER_ARROWS
    If Player(Index).Arrow(z).Arrow > 0 Then
    
        rec.Top = Player(Index).Arrow(z).ArrowAnim * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(Index).Arrow(z).ArrowPosition * PIC_X
        rec.Right = rec.Left + PIC_X
        
        If GetTickCount > Player(Index).Arrow(z).ArrowTime + 30 Then
            Player(Index).Arrow(z).ArrowTime = GetTickCount
            Player(Index).Arrow(z).ArrowVarX = Player(Index).Arrow(z).ArrowVarX + 20
            Player(Index).Arrow(z).ArrowVarY = Player(Index).Arrow(z).ArrowVarY + 20
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
                If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                    Player(Index).Arrow(z).Arrow = 0
                End If
            End If
        End If
        
        For I = 1 To MAX_PLAYERS
           If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(I) = X And GetPlayerY(I) = Y Then
                    If Index = MyIndex Then
                        Call SendData("arrowhit" & SEP_CHAR & 0 & SEP_CHAR & I & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR)
                    End If
                    If Index <> I Then Player(Index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next I
        
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(I).Num > 0 Then
                If MapNpc(I).X = X And MapNpc(I).Y = Y Then
                    If Index = MyIndex Then
                        Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & I & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR)
                    End If
                    Player(Index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next I
    End If
Next z
End Sub

Sub BltMiniMap()
Dim I As Long
Dim X As Integer
Dim Y As Integer
Dim MMx As Long
Dim MMy As Integer

    ' Tiles Layer
    ' Select MM Tile to Use for Tiles Layer
    rec.Top = 8
    rec.Bottom = 16
    rec.Left = 0
    rec.Right = 8
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            If Map(Player(MyIndex).Map).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                MMx = 400 + (X * 8)
                MMy = 32 + (Y * 8)
                Call DD_BackBuffer.BltFast(MMx, MMy, DD_MiniMap, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next X
    Next Y
    
    
    ' Player Layer
    ' Select MM Tile to Use for Players Layer
    rec.Top = 16
    rec.Bottom = 24
    rec.Left = 0
    rec.Right = 8
    
    For I = 1 To MAX_PLAYERS
        If Player(I).Map = Player(MyIndex).Map Then
            X = Player(I).X
            Y = Player(I).Y
            MMx = 400 + (X * 8)
            MMy = 32 + (Y * 8)
            If Not I = MyIndex Then
                Call DD_BackBuffer.BltFast(MMx, MMy, DD_MiniMap, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Next I

    ' MyPlayer Layer
    rec.Top = 32
    rec.Bottom = 40
    rec.Left = 0
    rec.Right = 8
    X = Player(MyIndex).X
    Y = Player(MyIndex).Y
    MMx = 400 + (X * 8)
    MMy = 32 + (Y * 8)
    Call DD_BackBuffer.BltFast(MMx, MMy, DD_MiniMap, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    
    ' NPC Layer
    ' Select the MM Tile to use for the NPC Layer
    rec.Top = 24
    rec.Bottom = 32
    rec.Left = 0
    rec.Right = 8
    
    For I = 1 To MAX_MAP_NPCS
        If Map(MyIndex).Npc(I) & MapNpc(I).HP > 0 Then
            X = MapNpc(I).X
            Y = MapNpc(I).Y
            MMx = 400 + (X * 8)
            MMy = 32 + (Y * 8)
            Call DD_BackBuffer.BltFast(MMx, MMy, DD_MiniMap, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Next I
    
End Sub

Sub BltPlayerCorpse(ByVal Index As Integer)
Dim X As Long, Y As Long
rec.Top = 0
rec.Bottom = 32
rec.Left = 0
rec.Right = 32

X = CLng(Player(Index).CorpseX * 32)
Y = CLng(Player(Index).CorpseY * 32)
Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) + sx - NewXOffset, Y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_CorpseAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerCorpseName(ByVal Index As Integer)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    
Color = QBColor(Grey)

        
    ' Draw name
    TextX = Player(Index).CorpseX * PIC_X + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
    TextY = Player(Index).CorpseY * PIC_Y + sx - Int(PIC_Y / 2) - (SIZE_Y - PIC_Y)
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), Color)
End Sub
