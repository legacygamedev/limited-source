Attribute VB_Name = "modDirectX"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

Private Resp As Long

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

Public DDSD_MiniMap As DDSURFACEDESC2
Public DD_MiniMap As DirectDrawSurface7

Public DD_Icon As DirectDrawSurface7
Public DDSD_Icon As DDSURFACEDESC2

Public DD_CorpseAnim As DirectDrawSurface7
Public DDSD_CorpseAnim As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX(Optional mmStart As Boolean = False)
    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
 
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hwnd, DDSCL_NORMAL)
        
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
        DD_Clip.SetHWnd frmMirage.picScreen.hwnd
            
        ' Have the blits to the screen clipped to the picture box
        DD_PrimarySurf.SetClipper DD_Clip
    End If
    ' Initialize all surfaces
    Call InitSurfaces(mmStart)
End Sub

Private Sub InitSurface(ByVal Filename As String, ByVal StrPW As String, ByRef DDSD As DDSURFACEDESC2, ByRef DDSurf As DirectDrawSurface7)
Dim sDc As Long
Dim BMU As BitmapUtils

    Set BMU = New BitmapUtils
    
    With BMU
        Call .LoadByteData(Filename)
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
Dim i As Long
Dim StrPW As String

    ' Our password is going to be 'test'.. We set it with Chr$(num)
    ' t = 116, e = 101, s = 115, t = 116
    ' Doing it this way keeps someone from hex editing your password easily.
    ' You can also just do: StrPW = "test" The result would be the same.
    StrPW = GFX_PASSWORD
    
    ' Check for files existing
    If FileExist("\GFX\sprites.gfx") = False Or FileExist("\GFX\items.gfx") = False Or FileExist("\GFX\bigsprites.gfx") = False Or FileExist("\GFX\emoticons.gfx") = False Or FileExist("\GFX\arrows.gfx") = False Or FileExist("\GFX\minimap.gfx") = False Then
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
    
      ' Init tiles ddsd type and load the bitmap
      For i = 0 To ExtraSheets
          If Dir(App.Path & "\GFX\tiles" & i & ".gfx") <> "" Then
              Call InitSurface(App.Path & "\GFX\tiles" & i & ".gfx", StrPW, DDSD_Tile(i), DD_TileSurf(i))
              TileFile(i) = 1
          Else
              TileFile(i) = 0
          End If
      Next i
    
      ' Init big sprites ddsd type and load the bitmap
      Call InitSurface(App.Path & "\GFX\bigsprites.gfx", StrPW, DDSD_BigSprite, DD_BigSpriteSurf)
      
      ' Init emoticons ddsd type and load the bitmap
      Call InitSurface(App.Path & "\GFX\emoticons.gfx", StrPW, DDSD_Emoticon, DD_EmoticonSurf)
      
      ' Init spells ddsd type and load the bitmap
      Call InitSurface(App.Path & "\GFX\spells.gfx", StrPW, DDSD_SpellAnim, DD_SpellAnim)
      
      ' Init arrows ddsd type and load the bitmap
      Call InitSurface(App.Path & "\GFX\arrows.gfx", StrPW, DDSD_ArrowAnim, DD_ArrowAnim)
    End If
    
    ' Init sprite ddsd type and load the bitmap
    Call InitSurface(App.Path & "\GFX\sprites.gfx", StrPW, DDSD_Sprite, DD_SpriteSurf)
    
    ' Init items ddsd type and load the bitmap
    Call InitSurface(App.Path & "\GFX\items.gfx", StrPW, DDSD_Item, DD_ItemSurf)

    ' Init items ddsd type and load the bitmap
      Call InitSurface(App.Path & "\GFX\minimap.gfx", StrPW, DDSD_MiniMap, DD_MiniMap)
      
      ' Init Spell Icons ddsd type and load the bitmap
      Call InitSurface(App.Path & "\GFX\spellicons.gfx", StrPW, DDSD_Icon, DD_Icon)

      ' Init Corpses ddsd type and load the bitmap
      Call InitSurface(App.Path & "\GFX\corpse.gfx", StrPW, DDSD_CorpseAnim, DD_CorpseAnim)
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
    Set DD_Icon = Nothing
    Set DD_CorpseAnim = Nothing
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
    
     If Val(GetVar(App.Path & "\config.ini", "CONFIG", "Night")) = 1 Then
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
        
    For y = NewY To NewY2
        For x = NewX To NewX2
            If Map(GetPlayerMap(MyIndex)).Tile(x, y).Light <= 0 Then
                DisplayFx DD_TileSurf(6), (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31
            Else
                DisplayFx DD_TileSurf(6), (x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Map(GetPlayerMap(MyIndex)).Tile(x, y).Light
            End If
        Next x
    Next y
End If
End Sub

Sub BltWeather()
Dim i As Long


    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))
    
 If Val(GetVar(App.Path & "\config.ini", "CONFIG", "Weather")) = 1 Then
   
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
            DropRain(i).x = DropRain(i).x + DropRain(i).speed
            DropRain(i).y = DropRain(i).y + DropRain(i).speed
            Call DD_BackBuffer.DrawLine(DropRain(i).x, DropRain(i).y, DropRain(i).x + DropRain(i).speed, DropRain(i).y + DropRain(i).speed)
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
                DropSnow(i).x = DropSnow(i).x + DropSnow(i).speed
                DropSnow(i).y = DropSnow(i).y + DropSnow(i).speed
                Call DD_BackBuffer.BltFast(DropSnow(i).x + DropSnow(i).speed, DropSnow(i).y + DropSnow(i).speed, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).x > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropRain(RDNumber).speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropRain(RDNumber).x = 0
    DropRain(RDNumber).y = 0
    DropRain(RDNumber).speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropSnow(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropSnow(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropSnow(RDNumber).x > (MAX_MAPX + 1) * PIC_X / 4) Then GoTo Start
    DropSnow(RDNumber).speed = Int((10 * Rnd) + 6)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
On Error Resume Next
    DropSnow(RDNumber).x = 0
    DropSnow(RDNumber).y = 0
    DropSnow(RDNumber).speed = 0
    DropSnow(RDNumber).Randomized = False
End Sub

Sub BltSpell(ByVal Index As Long)
Dim x As Long, y As Long, i As Long

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
                        
            rec.Top = Spell(Player(Index).SpellNum).SpellAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = Player(Index).SpellAnim(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X
            
            If Player(Index).SpellAnim(i).TargetType = TARGET_TYPE_PLAYER Then
                If Player(Index).SpellAnim(i).Target > 0 Then
                    If Player(Index).SpellAnim(i).Target = MyIndex Then
                        x = NewX + sx
                        y = NewY + sx
                        Call DD_BackBuffer.BltFast(x, y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        x = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx + Player(Player(Index).SpellAnim(i).Target).XOffset
                        y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx + Player(Player(Index).SpellAnim(i).Target).YOffset
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            Else
                If Player(Index).SpellAnim(i).TargetType = TARGET_TYPE_NPC Then
                    x = MapNpc(Player(Index).SpellAnim(i).Target).x * PIC_X + sx + MapNpc(Player(Index).SpellAnim(i).Target).XOffset
                    y = MapNpc(Player(Index).SpellAnim(i).Target).y * PIC_Y + sx + MapNpc(Player(Index).SpellAnim(i).Target).YOffset
                    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    If Player(Index).SpellAnim(i).TargetType = TARGET_TYPE_LOCATION Then
                        x = MakeX(Player(Index).SpellAnim(i).Target) * PIC_X + sx
                        y = MakeY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
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
   
    If Player(Index).EmoticonNum < 0 And Player(Index).EmoticonPlayed Then Exit Sub
    
    If (Player(Index).EmoticonType = EMOTICON_TYPE_SOUND Or Player(Index).EmoticonType = EMOTICON_TYPE_BOTH) And Player(Index).EmoticonPlayed = False And Val(GetVar("config.ini", "CONFIG", "EmoticonSound")) = 1 Then
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
            If MapNpc(i).Num > 0 Then
                If MapNpc(i).x = x And MapNpc(i).y = y Then
                    If Index = MyIndex Then
                        Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                    End If
                    Player(Index).Arrow(z).Arrow = 0
                    Exit Sub
                End If
            End If
        Next i
    End If
Next z
End Sub

Sub BltMiniMap()
Dim i As Long
Dim x As Integer
Dim y As Integer
Dim MMx As Long
Dim MMy As Integer

    ' Tiles Layer
    ' Select MM Tile to Use for Tiles Layer
    rec.Top = 8
    rec.Bottom = 16
    rec.Left = 0
    rec.Right = 8
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            If Map(Player(MyIndex).Map).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
                MMx = 400 + (x * 8)
                MMy = 32 + (y * 8)
                Call DD_BackBuffer.BltFast(MMx, MMy, DD_MiniMap, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next x
    Next y
    
    
    ' Player Layer
    ' Select MM Tile to Use for Players Layer
    rec.Top = 16
    rec.Bottom = 24
    rec.Left = 0
    rec.Right = 8
    
    For i = 1 To MAX_PLAYERS
        If Player(i).Map = Player(MyIndex).Map Then
            x = Player(i).x
            y = Player(i).y
            MMx = 400 + (x * 8)
            MMy = 32 + (y * 8)
            If Not i = MyIndex Then
                Call DD_BackBuffer.BltFast(MMx, MMy, DD_MiniMap, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Next i

    ' MyPlayer Layer
    rec.Top = 32
    rec.Bottom = 40
    rec.Left = 0
    rec.Right = 8
    x = Player(MyIndex).x
    y = Player(MyIndex).y
    MMx = 400 + (x * 8)
    MMy = 32 + (y * 8)
    Call DD_BackBuffer.BltFast(MMx, MMy, DD_MiniMap, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    
    ' NPC Layer
    ' Select the MM Tile to use for the NPC Layer
    rec.Top = 24
    rec.Bottom = 32
    rec.Left = 0
    rec.Right = 8
    
    For i = 1 To MAX_MAP_NPCS
        If Map(MyIndex).Npc(i) & MapNpc(i).HP > 0 Then
            x = MapNpc(i).x
            y = MapNpc(i).y
            MMx = 400 + (x * 8)
            MMy = 32 + (y * 8)
            Call DD_BackBuffer.BltFast(MMx, MMy, DD_MiniMap, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Next i
    
End Sub

Sub BltPlayerCorpse(ByVal Index As Integer)
Dim x As Long, y As Long
rec.Top = 0
rec.Bottom = 32
rec.Left = 0
rec.Right = 32

x = CLng(Player(Index).CorpseX * 32)
y = CLng(Player(Index).CorpseY * 32)
Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_CorpseAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
