Attribute VB_Name = "modDirectX"
Option Explicit

Public DX As DirectX7
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

Public DD_player_head As DirectDrawSurface7
Public DDSD_player_head As DDSURFACEDESC2

Public DD_player_body As DirectDrawSurface7
Public DDSD_player_body As DDSURFACEDESC2

Public DD_player_legs As DirectDrawSurface7
Public DDSD_player_legs As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX()
    On Error GoTo DXErr

    ' Initialize DirextX
    Set DX = New DirectX7

    ' Initialize DirectDraw
    Set DD = DX.DirectDrawCreate(vbNullString)

    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hWnd, DDSCL_NORMAL)

    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)

    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)

    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMirage.picScreen.hWnd

    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip

    ' Initialize all surfaces
    Call InitSurfaces
    Exit Sub

    ' Error handling
DXErr:
    Call MsgBox("Error initializing DirectDraw! Make sure you have DirectX 7 or higher installed and a compatible graphics device. Err: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    Call GameDestroy
    End
End Sub

Sub InitSurfaces()
    Dim Key As DDCOLORKEY
    Dim i As Long
    Dim DC As Long

    ' Check for files existing
    If Not FileExists("\GFX\Sprites.bmp") Or Not FileExists("\GFX\Items.bmp") Or Not FileExists("\GFX\BigSprites.bmp") Or Not FileExists("\GFX\Emoticons.bmp") Or Not FileExists("\GFX\Arrows.bmp") Then
        Call MsgBox("Your missing some graphic files!", vbOKOnly, GAME_NAME)
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

    ' Init sprite ddsd type and load the bitmap
    DDSD_Sprite.lFlags = DDSD_CAPS
    DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\Sprites.bmp", DDSD_Sprite)
    SetMaskColorFromPixel DD_SpriteSurf, 0, 0

    ' Init tiles ddsd type and load the bitmap
    For i = 0 To ExtraSheets
        If Dir$(App.Path & "\GFX\Tiles" & i & ".bmp") <> vbNullString Then
            DDSD_Tile(i).lFlags = DDSD_CAPS
            DDSD_Tile(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_TileSurf(i) = DD.CreateSurfaceFromFile(App.Path & "\GFX\Tiles" & i & ".bmp", DDSD_Tile(i))
            SetMaskColorFromPixel DD_TileSurf(i), 0, 0
            TileFile(i) = 1
        Else
            TileFile(i) = 0
        End If
    Next i

    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\Items.bmp", DDSD_Item)
    SetMaskColorFromPixel DD_ItemSurf, 0, 0

    ' Init big sprites ddsd type and load the bitmap
    DDSD_BigSprite.lFlags = DDSD_CAPS
    DDSD_BigSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_BigSpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\BigSprites.bmp", DDSD_BigSprite)
    SetMaskColorFromPixel DD_BigSpriteSurf, 0, 0

    ' Init emoticons ddsd type and load the bitmap
    DDSD_Emoticon.lFlags = DDSD_CAPS
    DDSD_Emoticon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_EmoticonSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\Emoticons.bmp", DDSD_Emoticon)
    SetMaskColorFromPixel DD_EmoticonSurf, 0, 0

    ' Init spells ddsd type and load the bitmap
    DDSD_SpellAnim.lFlags = DDSD_CAPS
    DDSD_SpellAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpellAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\Spells.bmp", DDSD_SpellAnim)
    SetMaskColorFromPixel DD_SpellAnim, 0, 0

    ' Init spells ddsd type and load the bitmap
    DDSD_BigSpellAnim.lFlags = DDSD_CAPS
    DDSD_BigSpellAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_BigSpellAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\BigSpells.bmp", DDSD_BigSpellAnim)
    SetMaskColorFromPixel DD_BigSpellAnim, 0, 0

    ' Init arrows ddsd type and load the bitmap
    DDSD_ArrowAnim.lFlags = DDSD_CAPS
    DDSD_ArrowAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ArrowAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\Arrows.bmp", DDSD_ArrowAnim)
    SetMaskColorFromPixel DD_ArrowAnim, 0, 0

    If CustomPlayers <> 0 Then
        ' Init head ddsd type and load the bitmap
        DDSD_player_head.lFlags = DDSD_CAPS
        DDSD_player_head.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_player_head = DD.CreateSurfaceFromFile(App.Path & "\GFX\heads.bmp", DDSD_player_head)
        SetMaskColorFromPixel DD_player_head, 0, 0

        ' Init body ddsd type and load the bitmap
        DDSD_player_body.lFlags = DDSD_CAPS
        DDSD_player_body.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_player_body = DD.CreateSurfaceFromFile(App.Path & "\GFX\bodys.bmp", DDSD_player_body)
        SetMaskColorFromPixel DD_player_body, 0, 0

        ' Init legs ddsd type and load the bitmap
        DDSD_player_legs.lFlags = DDSD_CAPS
        DDSD_player_legs.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_player_legs = DD.CreateSurfaceFromFile(App.Path & "\GFX\legs.bmp", DDSD_player_legs)
        SetMaskColorFromPixel DD_player_legs, 0, 0
    End If
End Sub

Sub DestroyDirectX()
    Dim i As Long

    Set DX = Nothing
    Set DD = Nothing

    Set DD_Clip = Nothing

    Set DD_PrimarySurf = Nothing
    Set DD_BackBuffer = Nothing

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

    Set DD_player_head = Nothing
    Set DD_player_body = Nothing
    Set DD_player_legs = Nothing

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

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal y As Long)
    Dim TmpR As RECT
    Dim TmpDDSD As DDSURFACEDESC2
    Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = X
        .Top = y
        .Right = X
        .Bottom = y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .low = TheSurface.GetLockedPixel(X, y)
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

Public Function GetScreenLeft(ByVal Index As Long) As Long
    GetScreenLeft = GetPlayerX(Index) - 11
End Function

Public Function GetScreenTop(ByVal Index As Long) As Long
    GetScreenTop = GetPlayerY(Index) - 8
End Function

Public Function GetScreenRight(ByVal Index As Long) As Long
    GetScreenRight = GetPlayerX(Index) + 10
End Function

Public Function GetScreenBottom(ByVal Index As Long) As Long
    GetScreenBottom = GetPlayerY(Index) + 8
End Function

Sub Night()
    Dim X As Long, y As Long

    If TileFile(10) = 0 Then
        Exit Sub
    End If

    For y = ScreenY To ScreenY2
        For X = ScreenX To ScreenX2
            If Map(GetPlayerMap(MyIndex)).Tile(X, y).light <= 0 Then
                DisplayFx DD_TileSurf(10), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31
            Else
                DisplayFx DD_TileSurf(10), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, CLng(Map(GetPlayerMap(MyIndex)).Tile(X, y).light)
            End If
        Next X
    Next y
End Sub

Sub BltTile2(ByVal X As Long, ByVal y As Long, ByVal Tile As Long)
    If TileFile(10) = 0 Then
        Exit Sub
    End If

    rec.Top = Int(Tile / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
'    DisplayFx DD_TileSurf(10), (x - NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, 32, 16, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Tile
End Sub

Sub BltPlayerText(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim intLoop As Integer
    Dim intLoop2 As Integer

    Dim bytLineCount As Byte
    Dim bytLineLength As Byte
    Dim strLine(0 To MAX_LINES - 1) As String
    Dim strWords() As String
    strWords() = Split(Bubble(Index).Text, " ")

    If Len(Bubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(Bubble(Index).Text) * 9) \ PIC_X)

        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If

    TextX = GetPlayerX(Index) * PIC_X + Player(Index).xOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - Int(PIC_Y) + 75

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)

        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1

        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1

            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If

            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)

            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 4) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 4) - 4, 4)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 4, TextY - (NewPlayerY * PIC_Y) + sx - NewYOffset, strLine(intLoop), QBColor(WHITE))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub Bltscriptbubble(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal y As Long, ByVal Colour As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim intLoop As Integer
    Dim intLoop2 As Integer

    Dim bytLineCount As Byte
    Dim bytLineLength As Byte
    Dim strLine(0 To MAX_LINES - 1) As String
    Dim strWords() As String

    strWords() = Split(ScriptBubble(Index).Text, " ")

    If Len(ScriptBubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(ScriptBubble(Index).Text) * 9) \ PIC_X)

        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If

    ' TextX = X * PIC_X + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextX = X * PIC_X - 22
    TextY = y * PIC_Y - 22

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1

        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1

            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If

            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)

            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 4)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY, strLine(intLoop), QBColor(Colour))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub BltPlayerBars(ByVal Index As Long)
    Dim X As Long, y As Long

    X = (GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
    y = (GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset


    If Player(Index).HP = 0 Then
        Exit Sub
    End If
    If SpriteSize = 1 Then
        ' draws the back bars
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, y - 30, X + 32, y - 34)

        ' draws HP
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(X, y - 30, X + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), y - 34)
    Else
        If SpriteSize = 2 Then
            ' draws the back bars
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
            Call DD_BackBuffer.DrawBox(X, y - 30 - PIC_Y, X + 32, y - 34 - PIC_Y)

            ' draws HP
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
            Call DD_BackBuffer.DrawBox(X, y - 30 - PIC_Y, X + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), y - 34 - PIC_Y)
        Else
            ' draws the back bars
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
            Call DD_BackBuffer.DrawBox(X, y + 2, X + 32, y - 2)

            ' draws HP
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
            Call DD_BackBuffer.DrawBox(X, y + 2, X + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), y - 2)
        End If
    End If
End Sub

Sub BltNpcBars(ByVal Index As Long)
    Dim X As Long, y As Long

    On Error GoTo BltNpcBars_Error

    If MapNpc(Index).HP = 0 Then
        Exit Sub
    End If
    If MapNpc(Index).num < 1 Then
        Exit Sub
    End If

    If Npc(MapNpc(Index).num).Big = 1 Then
        X = (MapNpc(Index).X * PIC_X + sx - 9 + MapNpc(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset

        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, y + 32, X + 50, y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        If MapNpc(Index).MaxHp < 1 Then
            Call DD_BackBuffer.DrawBox(X, y + 32, X + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHp + 1) / 100) * 50), y + 36)
        Else
            Call DD_BackBuffer.DrawBox(X, y + 32, X + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 50), y + 36)
        End If
    Else
        X = (MapNpc(Index).X * PIC_X + sx + MapNpc(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset

        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, y + 32, X + 32, y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))

        If MapNpc(Index).MaxHp < 1 Then
            Call DD_BackBuffer.DrawBox(X, y + 32, X + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHp + 1) / 100) * 32), y + 36)
        Else
            Call DD_BackBuffer.DrawBox(X, y + 32, X + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 32), y + 36)
        End If

    End If


    On Error GoTo 0
    Exit Sub

BltNpcBars_Error:

    If Err.Number = DDERR_CANTCREATEDC Then

    End If

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
                    If frmMirage.tmrRainDrop.Tag = vbNullString Then
                        frmMirage.tmrRainDrop.Interval = 100
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
                    If frmMirage.tmrSnowDrop.Tag = vbNullString Then
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
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If

    For i = 1 To MAX_RAINDROPS
        If Not ((DropRain(i).X = 0) Or (DropRain(i).y = 0)) Then
            rec.Top = 0
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = 6 * PIC_X
            rec.Right = rec.Left + PIC_X
            DropRain(i).X = DropRain(i).X + DropRain(i).speed
            DropRain(i).y = DropRain(i).y + DropRain(i).speed
            Call DD_BackBuffer.BltFast(DropRain(i).X + DropRain(i).speed, DropRain(i).y + DropRain(i).speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            If (DropRain(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).y > (MAX_MAPY + 1) * PIC_Y) Then
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
            If Not ((DropSnow(i).X = 0) Or (DropSnow(i).y = 0)) Then
                DropSnow(i).X = DropSnow(i).X + DropSnow(i).speed
                DropSnow(i).y = DropSnow(i).y + DropSnow(i).speed
                Call DD_BackBuffer.BltFast(DropSnow(i).X + DropSnow(i).speed, DropSnow(i).y + DropSnow(i).speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropSnow(i).Randomized = False
                End If
            End If
        Next i
    End If

    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then
        If Int((100 - 1 + 1) * Rnd) + 1 = 8 Then
            DD_BackBuffer.SetFillColor RGB(255, 255, 255)

            Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
        End If
    End If
End Sub

Sub BltMapWeather()
    Dim i As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))

    If Map(GetPlayerMap(MyIndex)).Weather = 1 Or Map(GetPlayerMap(MyIndex)).Weather = 3 Then
        For i = 1 To MAX_RAINDROPS
            If DropRain(i).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                End If
            End If
        Next i
        For i = 1 To MAX_RAINDROPS
            If Not ((DropRain(i).X = 0) Or (DropRain(i).y = 0)) Then
                rec.Top = (14 - Int(14 / TilesInSheets)) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = 6 * PIC_X
                rec.Right = rec.Left + PIC_X
                DropRain(i).X = DropRain(i).X + DropRain(i).speed
                DropRain(i).y = DropRain(i).y + DropRain(i).speed
                Call DD_BackBuffer.BltFast(DropRain(i).X + DropRain(i).speed, DropRain(i).y + DropRain(i).speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropRain(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropRain(i).Randomized = False
                End If
            End If
        Next i

        If Map(GetPlayerMap(MyIndex)).Weather = 3 Then
            If Int((100 - 1 + 1) * Rnd) + 1 < 3 Then
                DD_BackBuffer.SetFillColor RGB(255, 255, 255)

                Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
            End If
        End If

    ElseIf Map(GetPlayerMap(MyIndex)).Weather = 2 Then
        For i = 1 To MAX_RAINDROPS
            If DropSnow(i).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                End If
            End If
        Next i
        If TileFile(10) = 1 Then
            rec.Top = Int(14 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (14 - Int(14 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            For i = 1 To MAX_RAINDROPS
                If Not ((DropSnow(i).X = 0) Or (DropSnow(i).y = 0)) Then
                    DropSnow(i).X = DropSnow(i).X + DropSnow(i).speed
                    DropSnow(i).y = DropSnow(i).y + DropSnow(i).speed
                    Call DD_BackBuffer.BltFast(DropSnow(i).X + DropSnow(i).speed, DropSnow(i).y + DropSnow(i).speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    If (DropSnow(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                        DropSnow(i).Randomized = False
                    End If
                End If
            Next i
        End If
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= frmMirage.tmrRainDrop.Interval Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).X > (MAX_MAPX + 1) * PIC_X / 4) Then
        GoTo Start
    End If
    DropRain(RDNumber).speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
    On Error Resume Next
    DropRain(RDNumber).X = 0
    DropRain(RDNumber).y = 0
    DropRain(RDNumber).speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropSnow(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropSnow(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropSnow(RDNumber).X > (MAX_MAPX + 1) * PIC_X / 4) Then
        GoTo Start
    End If
    DropSnow(RDNumber).speed = Int((10 * Rnd) + 6)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
    On Error Resume Next
    DropSnow(RDNumber).X = 0
    DropSnow(RDNumber).y = 0
    DropSnow(RDNumber).speed = 0
    DropSnow(RDNumber).Randomized = False
End Sub

Sub BltSpell(ByVal Index As Long)
    Dim X As Long, y As Long, i As Long

    If Player(Index).SpellNum <= 0 Or Player(Index).SpellNum > MAX_SPELLS Then
        Exit Sub
    End If


    For i = 1 To MAX_SPELL_ANIM
        ' IF SPELL IS NOT BIG
        If Spell(Player(Index).SpellNum).Big = 0 Then
            If Player(Index).SpellAnim(i).CastedSpell = YES Then
                If Player(Index).SpellAnim(i).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then

                    rec.Top = Spell(Player(Index).SpellNum).SpellAnim * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    rec.Left = Player(Index).SpellAnim(i).SpellVar * PIC_X
                    rec.Right = rec.Left + PIC_X

                    If Player(Index).SpellAnim(i).TargetType = 0 Then

                        ' SMALL: IF TARGET IS A PLAYER
                        If Player(Index).SpellAnim(i).Target > 0 Then

                            ' SMALL: IF TARGET IS SELF
                            If Player(Index).SpellAnim(i).Target = MyIndex Then
                                X = NewX + sx
                                y = NewY + sx
                                Call DD_BackBuffer.BltFast(X, y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                            ' SMALL: IF TARGET IS ANOTHER PLAYER
                            Else
                                X = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx + Player(Player(Index).SpellAnim(i).Target).xOffset
                                y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx + Player(Player(Index).SpellAnim(i).Target).yOffset
                                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            End If
                        End If

                    ' SMALL: IF TARGET IS AN NPC
                    Else
                        X = MapNpc(Player(Index).SpellAnim(i).Target).X * PIC_X + sx + MapNpc(Player(Index).SpellAnim(i).Target).xOffset
                        y = MapNpc(Player(Index).SpellAnim(i).Target).y * PIC_Y + sx + MapNpc(Player(Index).SpellAnim(i).Target).yOffset
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If


' SMALL: ADVANCE SPELL ONE CYCLE

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

                        ' BIG: IF TARGET IS A PLAYER
                        If Player(Index).SpellAnim(i).Target > 0 Then

                            ' BIG: IF TARGET IS SELF
                            If Player(Index).SpellAnim(i).Target = MyIndex Then
                                X = NewX + sx - 32
                                y = NewY + sx - 32

                                If y < 0 Then
                                    rec.Top = rec.Top + (y * -1)
                                    y = 0
                                End If

                                If X < 0 Then
                                    rec.Left = rec.Left + (X * -1)
                                    X = 0
                                End If

                                If (X + 64) > (MAX_MAPX * 32) Then
                                    rec.Right = rec.Left + 64
                                End If

                                If (y + 64) > (MAX_MAPY * 32) Then
                                    rec.Bottom = rec.Top + 64
                                End If

                                Call DD_BackBuffer.BltFast(X, y, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                            ' BIG: IF TARGET IS A DIFFERENT PLAYER
                            Else
                                X = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx - 32 + Player(Player(Index).SpellAnim(i).Target).xOffset
                                y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx - 32 + Player(Player(Index).SpellAnim(i).Target).yOffset

                                If y < 0 Then
                                    rec.Top = rec.Top + (y * -1)
                                    y = 0
                                End If

                                If X < 0 Then
                                    rec.Left = rec.Left + (X * -1)
                                    X = 0
                                End If

                                If (X + 64) > (MAX_MAPX * 32) Then
                                    rec.Right = rec.Left + 64
                                End If

                                If (y + 64) > (MAX_MAPY * 32) Then
                                    rec.Bottom = rec.Top + 64
                                End If

                                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            End If
                        End If

                    ' BIG: IF TARGET IS AN NPC
                    Else
                        X = MapNpc(Player(Index).SpellAnim(i).Target).X * PIC_X + sx - 32 + MapNpc(Player(Index).SpellAnim(i).Target).xOffset
                        y = MapNpc(Player(Index).SpellAnim(i).Target).y * PIC_Y + sx - 32 + MapNpc(Player(Index).SpellAnim(i).Target).yOffset

                        If y < 0 Then
                            rec.Top = rec.Top + (y * -1)
                            y = 0
                        End If

                        If X < 0 Then
                            rec.Left = rec.Left + (X * -1)
                            X = 0
                        End If

                        If (X + 64) > (MAX_MAPX * 32) Then
                            rec.Right = rec.Left + 64
                        End If

                        If (y + 64) > (MAX_MAPY * 32) Then
                            rec.Bottom = rec.Top + 64
                        End If

                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' BIG: ADVANCE SPELL ONE CYCLE
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

' Scripted Spell
Sub BltScriptSpell(ByVal i As Long)
    Dim rec As RECT
    Dim X As Long, y As Long

    X = ScriptSpell(i).X
    y = ScriptSpell(i).y

    If Spell(ScriptSpell(i).SpellNum).Big = 0 Then
        If ScriptSpell(i).SpellDone < Spell(ScriptSpell(i).SpellNum).SpellDone Then
            rec.Top = Spell(ScriptSpell(i).SpellNum).SpellAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = ScriptSpell(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X

            X = X * PIC_X + sx
            y = y * PIC_Y + sx

            If ScriptSpell(i).SpellVar > 10 Then
                ScriptSpell(i).SpellDone = ScriptSpell(i).SpellDone + 1
                ScriptSpell(i).SpellVar = 0
            End If

            If GetTickCount > ScriptSpell(i).SpellTime + Spell(ScriptSpell(i).SpellNum).SpellTime Then
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).SpellVar = ScriptSpell(i).SpellVar + 1
            End If

            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X), y - (NewPlayerY * PIC_Y), DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

        Else ' spell is done
            ScriptSpell(i).CastedSpell = NO
        End If
    Else
        If ScriptSpell(i).SpellDone < Spell(ScriptSpell(i).SpellNum).SpellDone Then
            rec.Top = Spell(ScriptSpell(i).SpellNum).SpellAnim * (PIC_Y * 3)
            rec.Bottom = rec.Top + PIC_Y + 64
            rec.Left = ScriptSpell(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X + 64

            X = X * PIC_X + sx - 32
            y = y * PIC_Y + sx - 32

            If y < 0 Then
                rec.Top = rec.Top + (y * -1)
                y = 0
            End If

            If X < 0 Then
                rec.Left = rec.Left + (X * -1)
                X = 0
            End If

            If (X + 64) > (MAX_MAPX * 32) Then
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

            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X), y - (NewPlayerY * PIC_Y), DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else 'spell is done
            ScriptSpell(i).CastedSpell = NO
        End If
    End If
End Sub

Sub BltEmoticons(ByVal Index As Long)
    Dim x2 As Long, y2 As Long
    Dim ETime As Long
    ETime = 1300

    If Player(Index).EmoticonNum < 0 Then
        Exit Sub
    End If

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
            x2 = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + 16
            y2 = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32

            If y2 < 0 Then
                Exit Sub
            End If

            Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub Bltgrapple(ByVal Index As Long)
    Dim z As Integer
    Dim BX As Long, BY As Long

    If Player(Index).HookShotX > 0 Or Player(Index).HookShotY <> 0 Then

        Select Case Player(Index).HookShotDir
            Case 0
                z = 1
            Case 1
                z = 0
            Case 2
                z = 3
            Case 3
                z = 2
        End Select

        rec.Top = Player(Index).HookShotAnim * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = z * PIC_X
        rec.Right = rec.Left + PIC_X

        If GetTickCount > Player(Index).HookShotTime + 50 Then
            If Player(Index).HookShotSucces = 1 Then
                If Index = MyIndex Then
                Call SendData("endshot" & SEP_CHAR & 1 & END_CHAR)
                End If
                Player(Index).HookShotX = 0
                Player(Index).HookShotY = 0
            Else
                If Index = MyIndex Then
                Call SendData("endshot" & SEP_CHAR & 0 & END_CHAR)
                End If
                Player(Index).HookShotX = 0
                Player(Index).HookShotY = 0
            End If
        End If

        BX = GetPlayerX(Index)
        BY = GetPlayerY(Index)

        If Player(Index).HookShotDir = DIR_DOWN Then
            Do While BY <= Player(Index).HookShotToY
                If BY <= MAX_MAPY Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BY = BY + 1
            Loop
        End If

        If Player(Index).HookShotDir = DIR_UP Then
            Do While BY >= Player(Index).HookShotToY
                If BY >= 0 Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BY = BY - 1
            Loop
        End If

        If Player(Index).HookShotDir = DIR_RIGHT Then
            Do While BX <= Player(Index).HookShotToX
                If BX <= MAX_MAPX Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BX = BX + 1
            Loop
        End If

        If Player(Index).HookShotDir = DIR_LEFT Then
            Do While BX >= Player(Index).HookShotToX
                If BX >= 0 Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BX = BX - 1
            Loop
        End If
    End If
End Sub

Sub BltArrow(ByVal Index As Long)
    Dim X As Long
    Dim y As Long
    Dim i As Long
    Dim z As Long

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
                y = Player(Index).Arrow(z).ArrowY + Int(Player(Index).Arrow(z).ArrowVarY / 32)

                If y > Player(Index).Arrow(z).ArrowY + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                    Player(Index).Arrow(z).Arrow = 0
                End If

                If y <= MAX_MAPY Then
                    Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If Player(Index).Arrow(z).ArrowPosition = 1 Then
                X = Player(Index).Arrow(z).ArrowX
                y = Player(Index).Arrow(z).ArrowY - Int(Player(Index).Arrow(z).ArrowVarY / 32)

                If y < Player(Index).Arrow(z).ArrowY - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                    Player(Index).Arrow(z).Arrow = 0
                End If

                If y >= 0 Then
                    Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If Player(Index).Arrow(z).ArrowPosition = 2 Then
                X = Player(Index).Arrow(z).ArrowX + Int(Player(Index).Arrow(z).ArrowVarX / 32)
                y = Player(Index).Arrow(z).ArrowY

                If X > Player(Index).Arrow(z).ArrowX + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                    Player(Index).Arrow(z).Arrow = 0
                End If

                If X <= MAX_MAPX Then
                    Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If Player(Index).Arrow(z).ArrowPosition = 3 Then
                X = Player(Index).Arrow(z).ArrowX - Int(Player(Index).Arrow(z).ArrowVarX / 32)
                y = Player(Index).Arrow(z).ArrowY

                If X < Player(Index).Arrow(z).ArrowX - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                    Player(Index).Arrow(z).Arrow = 0
                End If

                If X >= 0 Then
                    Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If X >= 0 And X <= MAX_MAPX Then
                If y >= 0 And y <= MAX_MAPY Then
                    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_BLOCKED Then
                        Player(Index).Arrow(z).Arrow = 0
                    End If
                End If
            End If

            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If GetPlayerX(i) = X Then
                            If GetPlayerY(i) = y Then
                                If Index = MyIndex Then
                                    Call SendData("arrowhit" & SEP_CHAR & 0 & SEP_CHAR & i & SEP_CHAR & X & SEP_CHAR & y & END_CHAR)
                                End If

                                If Index <> i Then
                                    Player(Index).Arrow(z).Arrow = 0
                                End If

                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next i

            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    If MapNpc(i).X = X Then
                        If MapNpc(i).y = y Then
                            If Index = MyIndex Then
                                Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & i & SEP_CHAR & X & SEP_CHAR & y & END_CHAR)
                            End If

                            Player(Index).Arrow(z).Arrow = 0

                            Exit Sub
                        End If
                    End If
                End If
            Next i
        End If
    Next z
End Sub

Sub BltLevelUp(ByVal Index As Long)
    Dim rec As RECT
    Dim X As Integer
    Dim y As Integer

    If Player(Index).LevelUpT + 3000 > GetTickCount Then
        If GetPlayerMap(Index) = GetPlayerMap(MyIndex) Then
            rec.Top = PIC_Y * 2
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = PIC_X * 4
            rec.Right = rec.Left + 96

            X = GetPlayerX(Index) * PIC_X + Player(Index).xOffset + sx
            y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset + sx

            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - PIC_X - NewXOffset, y - (NewPlayerY * PIC_Y) - PIC_Y - NewYOffset - 8, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

            If Player(Index).LevelUp >= 3 Then
                Player(Index).LevelUp = Player(Index).LevelUp - 1
            ElseIf Player(Index).LevelUp >= 1 Then
                Player(Index).LevelUp = Player(Index).LevelUp + 1
            End If
        Else
            Player(Index).LevelUpT = 0
        End If
    End If
End Sub

Sub BltSpriteChange(ByVal X As Long, ByVal y As Long)
    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_SPRITE_CHANGE Then
        If SpriteSize = 0 Then
            rec.Top = Map(GetPlayerMap(MyIndex)).Tile(X, y).Data1 * PIC_Y + 16
            rec.Bottom = rec.Top + PIC_Y - 16
        Else
            rec.Top = Map(GetPlayerMap(MyIndex)).Tile(X, y).Data1 * 64 + 16
            rec.Bottom = rec.Top + 64 - 16
        End If
        
        rec.Left = 96
        rec.Right = rec.Left + PIC_X

        X = X * PIC_X + sx
        y = y * PIC_Y + (sx / 2) '- 16

        If y < 0 Then
            Exit Sub
        End If

        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub
