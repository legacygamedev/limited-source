Attribute VB_Name = "modDirectDraw7"
Option Explicit

' ******************************************
' **               rootSource               **
' ** DirectDraw7, renders graphics        **
' ******************************************

Private DD As DirectDraw7 ' DirectDraw7 Object

Private DD_Clip As DirectDrawClipper ' Clipper object

' primary surface
Private DDS_Primary As DirectDrawSurface7
Private DDSD_Primary As DDSURFACEDESC2

' back buffer
Public DDS_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

' gfx buffers
Public DDS_Item() As DD_BufferRec
Public DDS_Sprite() As DD_BufferRec
Public DDS_Spell() As DD_BufferRec
Public DDS_Tile As DD_BufferRec
Public DDS_Misc As DD_BufferRec
Private Type DD_BufferRec
    Surface As DirectDrawSurface7
    SurfDescription As DDSURFACEDESC2
    SurfTimer As Long
End Type

Public DDSD_Temp As DDSURFACEDESC2

Public Pixelation As Long

' Number of graphic files
Public NumTileSets As Long
Public NumSprites As Long
Public NumSpells As Long
Public NumItems As Long

Public Const SurfaceTimerMax As Long = 20000

' ** These values are used for transparency, comment out if getting mask color from the bitmap
' ** MASK_COLOR is the "transparency" color
'Public Const MASK_COLOR As Long = 0   ' RGB(0,0,0)
'Public Key As DXVBLib.DDCOLORKEY

' ********************
' ** Initialization **
' ********************
Public Function InitDirectDraw() As Boolean
On Error GoTo ErrorHandle

    ' Set the key for masks
    'With Key
    '    .Low = MASK_COLOR
    '    .High = MASK_COLOR
    'End With
    
    Call DestroyDirectDraw ' clear out everything
    
    ' Initialize direct draw
    Set DD = DX7.DirectDrawCreate(vbNullString) ' empty string forces primary device

    ' dictates how we access the screen and how other programs
    ' running at the same time will be allowed to access the screen as well.
    Call DD.SetCooperativeLevel(frmMainGame.hwnd, DDSCL_NORMAL)
        
    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
        
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hWnd with the clipper
    Call DD_Clip.SetHWnd(frmMainGame.picScreen.hwnd)
        
    ' Have the blits to the screen clipped to the picture box
    Call DDS_Primary.SetClipper(DD_Clip) ' method attaches a clipper object to, or deletes one from, a surface.
    
    Call InitBackBuffer
    
    Call DDS_BackBuffer.SetClipper(DD_Clip)
    
    InitDirectDraw = True
    
    Exit Function
    
ErrorHandle:

    Select Case Err
    
        Case -2147024770
            Call MsgBox("dx7vb.dll is either not found or is not registered, try re-installing directX or adding the file to your system directory.")
            Call DestroyGame
    
        Case 91
            Call MsgBox("DirectX7 master object not created.")
            Call DestroyGame

    End Select
    
    InitDirectDraw = False
    
End Function

Public Sub ReInitDD()
    Call InitDirectDraw
    Call InitTileSurf(map(5).TileSet)
    
    If Editor = EDITOR_MAP Then
        Call InitDDSurf("misc", DDS_Misc)
    End If
    
End Sub

Private Sub InitBackBuffer()
Dim rec As DXVBLib.RECT
  
    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (MAX_MAPX + 1) * PIC_X
        .lHeight = (MAX_MAPY + 1) * PIC_Y
    End With

    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)

    Call DDS_BackBuffer.BltColorFill(rec, 0) ' clear out trash
    
End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
Private Sub SetMaskColorFromPixel(ByRef Surface As DirectDrawSurface7, ByVal x As Long, ByVal y As Long)
  Dim TmpR As DXVBLib.RECT
  Dim TmpDDSD As DXVBLib.DDSURFACEDESC2
  Dim TmpColorKey As DXVBLib.DDCOLORKEY

    With TmpR
        .Left = x
        .Top = y
        .Right = x
        .Bottom = y
    End With

    Surface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = Surface.GetLockedPixel(x, y)
        .High = .Low
    End With

    Surface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    Surface.Unlock TmpR
    
End Sub

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(FileName As String, ByRef DD_SurfBuffer As DD_BufferRec)
'On Error GoTo ErrorHandle
    
    ' Set path
    FileName = App.Path & GFX_PATH & FileName & GFX_EXT
    
    ' Clear buffer
    Call DD_ClearBuffer(DD_SurfBuffer)
     
    ' set flags
    DD_SurfBuffer.SurfDescription.lFlags = DDSD_Temp.lFlags
    DD_SurfBuffer.SurfDescription.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
    
    ' init object
    Set DD_SurfBuffer.Surface = DD.CreateSurfaceFromFile(FileName, DD_SurfBuffer.SurfDescription)
    
    ' set color key (handle transparency)
    'Call DD_SurfBuffer.Surface.SetColorKey(DDCKEY_SRCBLT, Key) ' uses MASK_COLOR
    Call SetMaskColorFromPixel(DD_SurfBuffer.Surface, 0, 0)
    
    Exit Sub
    
ErrorHandle:

    Select Case Err
        
        ' File not found
        Case 53
            MsgBox "missing file: " & FileName
            Call DestroyGame
        
        ' DirectDraw does not have enough memory to perform the operation.
        Case DDERR_OUTOFMEMORY
            MsgBox "Out of system memory"
            Call DestroyGame
            
        ' DirectDraw does not have enough display memory to perform the operation.
        Case DDERR_OUTOFVIDEOMEMORY
            Call DevMsg("Out of video memory, attempting to re-initialize using system memory", BrightRed)

            DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY

            Call ReInitDD
                  
    End Select
    
End Sub

Public Sub InitTileSurf(ByVal TileSet As Integer)
    If TileSet < 1 Or TileSet > NumTileSets Then Exit Sub
    Call InitDDSurf("tiles" & TileSet, DDS_Tile)
End Sub

Public Sub DestroyDirectDraw()
Dim i As Long
    
    Call DD_ClearBuffer(DDS_Misc)
    Call DD_ClearBuffer(DDS_Tile)
    
    For i = 1 To NumItems
        Call DD_ClearBuffer(DDS_Item(i))
    Next

    For i = 1 To NumSpells
        Call DD_ClearBuffer(DDS_Spell(i))
    Next

    For i = 1 To NumSprites
        Call DD_ClearBuffer(DDS_Sprite(i))
    Next

    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    
    Set DD_Clip = Nothing
    Set DD = Nothing

End Sub

' ***********************
Private Function CheckSurfaces() As Boolean
On Error GoTo ErrorHandle

    ' Check if we need to restore surfaces
    If NeedToRestoreSurfaces Then
        DD.RestoreAllSurfaces
    End If
    
    CheckSurfaces = True
    Exit Function
    
ErrorHandle:

    Call ReInitDD
    CheckSurfaces = False
    
End Function

Private Function NeedToRestoreSurfaces() As Boolean
    If Not DD.TestCooperativeLevel = DD_OK Then
        NeedToRestoreSurfaces = True
    End If
End Function

' ***********************
Private Sub DD_ReadySurface(ByVal FileName As String, ByRef DD_SurfBuffer As DD_BufferRec)
    DD_SurfBuffer.SurfTimer = GetTickCount + SurfaceTimerMax
    If DD_SurfBuffer.Surface Is Nothing Then
        Call InitDDSurf(FileName, DD_SurfBuffer)
    End If
End Sub
Public Sub DD_CheckSurfTimer(ByRef DD_SurfBuffer As DD_BufferRec)
    If DD_SurfBuffer.SurfTimer > 0 Then
        If DD_SurfBuffer.SurfTimer < GetTickCount Then
            Call DD_ClearBuffer(DD_SurfBuffer)
        End If
    End If
End Sub
Public Sub DD_ClearBuffer(ByRef DD_SurfBuffer As DD_BufferRec)
    Set DD_SurfBuffer.Surface = Nothing
    Call ZeroMemory(ByVal VarPtr(DD_SurfBuffer.SurfDescription), LenB(DD_SurfBuffer.SurfDescription))
    DD_SurfBuffer.SurfTimer = 0
End Sub
' ************************


' **************
' ** Blitting **
' **************

Private Sub DD_BltFast(ByVal dx As Long, ByVal dy As Long, ByRef ddS As DirectDrawSurface7, srcRect As RECT, trans As CONST_DDBLTFASTFLAGS)
On Error GoTo ErrorHandle

Dim BltFastError As Long
    
    Dim destRect As RECT
    'BltFastError = DDS_BackBuffer.BltFast(dx, dy, ddS, srcRECT, trans)
    With destRect
        .Left = dx
        .Top = dy
        .Right = .Left + (srcRect.Right - srcRect.Left)
        .Bottom = .Top + (srcRect.Bottom - srcRect.Top)
    End With
    
    Call DDS_BackBuffer.Blt(destRect, ddS, srcRect, DDBLT_KEYSRC)
    'MsgBox srcRECT.Left & " " & srcRECT.Top

    If BltFastError <> 0 Then
        Select Case BltFastError
            Case DDERR_INVALIDRECT
                Call DevMsg("(DD_BltFast) Error Rendering Graphics: The provided rectangle was invalid." & srcRect.Right & " " & srcRect.Bottom & " " & srcRect.Top & " " & srcRect.Left, BrightRed)
        End Select
    End If
    
    Exit Sub

ErrorHandle:

    Select Case Err
        Case 5
            Call DevMsg("(DD_BltFast) Error Rendering Graphics: Attempting to copy from a surface thats not initialized.", BrightRed)
    End Select
    
End Sub

Private Function DD_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DXVBLib.RECT, dRECT As DXVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
On Error GoTo ErrorHandle
    
    If Clear Then
        picBox.Cls
    End If
    
    Call Surface.BltToDC(picBox.hDC, sRECT, dRECT)
    picBox.Refresh

    DD_BltToDC = True
    Exit Function

ErrorHandle:
    DD_BltToDC = False
    
    Call DevMsg("(DD_BltToDC) Error Rendering Graphics: Unable to draw to a picturebox.", BrightRed)

End Function

Public Sub BltMapTile(ByVal x As Long, ByVal y As Long)
    ' Blit out ground tile without transparency
    Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(0), DDBLTFAST_WAIT)

    If MapAnim = 0 Or ((MapTilePosition(x, y).Layer(2).Top = 0) And (MapTilePosition(x, y).Layer(2).Left = 0)) Then
        If ((MapTilePosition(x, y).Layer(1).Top <> 0) Or (MapTilePosition(x, y).Layer(1).Left <> 0)) Then
            If TempTile(x + NewPlayerX, y + NewPlayerY).DoorOpen = NO Then
                Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(1), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Else
        Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(2), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    If MapAnim = 0 Or ((MapTilePosition(x, y).Layer(4).Top = 0) And (MapTilePosition(x, y).Layer(4).Left = 0)) Then
        If ((MapTilePosition(x, y).Layer(3).Top <> 0) Or (MapTilePosition(x, y).Layer(3).Left <> 0)) Then
            Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(3), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(4), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
End Sub

Public Sub BltMapFringeTile(ByVal x As Long, ByVal y As Long)
    If MapAnim = 0 Or ((MapTilePosition(x, y).Layer(6).Top = 0) And (MapTilePosition(x, y).Layer(6).Left = 0)) Then
        If ((MapTilePosition(x, y).Layer(5).Top <> 0) Or (MapTilePosition(x, y).Layer(5).Left <> 0)) Then
            Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(5), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(6), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    If MapAnim = 0 Or ((MapTilePosition(x, y).Layer(8).Top = 0) And (MapTilePosition(x, y).Layer(8).Left = 0)) Then
        If ((MapTilePosition(x, y).Layer(7).Top <> 0) Or (MapTilePosition(x, y).Layer(7).Left <> 0)) Then
            Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(7), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        Call DD_BltFast(MapTilePosition(x, y).PosX - NewXOffset, MapTilePosition(x, y).PosY - NewYOffset, DDS_Tile.Surface, MapTilePosition(x, y).Layer(8), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Public Sub BltItem(ByVal ItemNum As Long, ByVal MapNum As Long)
Dim PicNum As Integer
Dim x As Long
Dim y As Long
Dim x2 As Long
Dim y2 As Long
Dim insight As Boolean
Dim rec As DXVBLib.RECT

    PicNum = Item(MapItem(ItemNum, MapNum).Num).Pic
    
    If PicNum < 1 Or PicNum > NumItems Then Exit Sub
    
    Call DD_ReadySurface("Items\" & PicNum, DDS_Item(PicNum))

            insight = False
                Select Case MapNum
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
If insight Then

    With rec
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
    
    x = MapItem(ItemNum, MapNum).x
    y = MapItem(ItemNum, MapNum).y
    
    Call DD_BltFast(MapTilePosition(x, y).PosX + StaticX + x2 * PIC_X, MapTilePosition(x, y).PosY + StaticY + y2 * PIC_Y, DDS_Item(PicNum).Surface, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
End Sub

Private Sub BltSprite(ByVal Sprite As Long, ByVal x As Long, ByVal y As Long, rec As DXVBLib.RECT)
    If Sprite < 1 Or Sprite > NumSprites Then Exit Sub
   
    Call DD_BltFast(x, y, DDS_Sprite(Sprite).Surface, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Private Sub BltSpell(ByVal spellnum As Long, ByVal x As Long, y As Long, rec As DXVBLib.RECT)
    If spellnum < 1 Or spellnum > NumSpells Then Exit Sub
    
    Call DD_ReadySurface("Spells\" & spellnum, DDS_Spell(spellnum))
    
    Call DD_BltFast(x + StaticX, y + StaticY, DDS_Spell(spellnum).Surface, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim x2 As Long, y2 As Long
Dim i As Long
Dim insight As Boolean
Dim Sprite As Long, spriteleft As Long
Dim rec As DXVBLib.RECT
Dim spellnum As Long


    Sprite = GetPlayerSprite(Index)
    If Sprite = 0 Then Sprite = 1
    Call DD_ReadySurface("sprites\" & Sprite, DDS_Sprite(Sprite))

            insight = False
                Select Case Player(Index).map
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
                
            If insight = True Then
                Anim = 0
                If Player(Index).Attacking = 0 Then
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            If Player(Index).YOffset < PIC_Y \ 2 Then Anim = 1
                        Case DIR_DOWN
                            If Player(Index).YOffset < PIC_Y \ -2 Then Anim = 1
                        Case DIR_LEFT
                            If Player(Index).XOffset < PIC_Y \ 2 Then Anim = 1
                        Case DIR_RIGHT
                            If Player(Index).XOffset < PIC_Y \ -2 Then Anim = 1
                    End Select
                Else
                    If Player(Index).AttackTimer + 500 > GetTickCount Then
                        Anim = 2
                    End If
                End If
                
                If Player(Index).AttackTimer + 1000 < GetTickCount Then
                    Player(Index).Attacking = 0
                    Player(Index).AttackTimer = 0
                End If
                
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spriteleft = DIR_UP
        Case DIR_RIGHT
            spriteleft = DIR_RIGHT
        Case DIR_DOWN
            spriteleft = DIR_DOWN
        Case DIR_LEFT
            spriteleft = DIR_LEFT
    End Select
                
    With rec
        .Top = 0
        .Bottom = DDS_Sprite(Sprite).SurfDescription.lHeight
        .Left = (spriteleft * 3 + Anim) * (DDS_Sprite(Sprite).SurfDescription.lWidth / 12)
        .Right = .Left + (DDS_Sprite(Sprite).SurfDescription.lWidth / 12)
    End With
    
    ' Calculate the X
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((DDS_Sprite(Sprite).SurfDescription.lWidth / 12 - 32) / 2)
    ' Is the player's height more than 32..?
    If (DDS_Sprite(Sprite).SurfDescription.lHeight) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((DDS_Sprite(Sprite).SurfDescription.lHeight) - 32)
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
    End If
    
    ' Check if its out of bounds because of the offset
    ' Is player's Y less than 0..?
    'If y < 0 Then
    '    With rec
    '        .Top = .Top - y
    '    End With
    '    y = 0
    'End If

    ' Is player's X less than 0..?
    'If x < 0 Then
    '    With rec
    '        .Left = .Left + (x * -1)
    '    End With
    '    x = 0
    'End If

    ' Is player's X more than max map values..?
    'If x + (DDS_Sprite(Sprite).SurfDescription.lWidth / 12) > MAX_MAPX * 32 + 32 Then
    '    With rec
    '        .Right = .Right + (x - (MAX_MAPX * 32))
    '    End With
    'End If
        
    Call BltSprite(Sprite, x2 * PIC_X + x + StaticX, y2 * PIC_Y + y + StaticY, rec)
    
        For i = 1 To MAX_SPELLANIM
        spellnum = Player(Index).SpellAnimations(i).spellnum
        
        If spellnum > 0 Then
            If Spell(spellnum).Pic > 0 Then
        
                If Player(Index).SpellAnimations(i).Timer < GetTickCount Then
                    Player(Index).SpellAnimations(i).FramePointer = Player(Index).SpellAnimations(i).FramePointer + 1
                    Player(Index).SpellAnimations(i).Timer = GetTickCount + 120
                                                                            
                    If Player(Index).SpellAnimations(i).FramePointer >= DDS_Spell(Spell(spellnum).Pic).SurfDescription.lWidth \ SIZE_X Then
                        Player(Index).SpellAnimations(i).spellnum = 0
                        Player(Index).SpellAnimations(i).Timer = 0
                        Player(Index).SpellAnimations(i).FramePointer = 0
                    End If
                End If
            
                If Player(Index).SpellAnimations(i).spellnum > 0 Then
                    With rec
                        .Top = 0
                        .Bottom = SIZE_Y
                        .Left = Player(Index).SpellAnimations(i).FramePointer * SIZE_X
                        .Right = .Left + SIZE_X
                    End With
            
                    Call BltSpell(Spell(spellnum).Pic, x2 * PIC_X + x, y2 * PIC_Y + y, rec)
                End If
                
            End If
        End If
        
    Next
                'X = 0 * PIC_X + GetPlayerX(Index) * PIC_X + Player(Index).XOffset - NewPlayerX * PIC_X - NewXOffset
                'Y = 0 * PIC_Y + GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - NewPlayerY * PIC_Y - NewYOffset
                
                'Call DrawTexture(X, Y, PIC_X, PIC_Y, (GetPlayerDir(i) * 3 + Anim) * PIC_X, GetPlayerSprite(i) * PIC_Y, PIC_X, PIC_Y, 255, 255, 255, 255, 512, 16384, 0)
    End If
End Sub


Public Sub BltNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim Anim As Long
Dim i As Long
Dim spellnum As Long
Dim x, x2 As Long
Dim y, y2 As Long
Dim Sprite As Long, spriteleft As Long
Dim rec As DXVBLib.RECT
Dim insight As Boolean

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum, MapNum).Num <= 0 Then
        Exit Sub
    End If

    Sprite = Npc(MapNpc(MapNpcNum, MapNum).Num).Sprite

            insight = False
                Select Case MapNum
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
If insight Then
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum, MapNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum, MapNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum, MapNum).YOffset < SIZE_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum, MapNum).YOffset < SIZE_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum, MapNum).XOffset < SIZE_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum, MapNum).XOffset < SIZE_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum, MapNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum, MapNum)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    'Pre-load graphic to get the width and height used in calculation
    Call DD_ReadySurface("sprites\" & Sprite, DDS_Sprite(Sprite))
   
    Select Case MapNpc(MapNpcNum, MapNum).Dir
        Case DIR_UP
            spriteleft = DIR_UP
        Case DIR_RIGHT
            spriteleft = DIR_RIGHT
        Case DIR_DOWN
            spriteleft = DIR_DOWN
        Case DIR_LEFT
            spriteleft = DIR_LEFT
    End Select
   
    With rec
        .Top = 0
        .Bottom = DDS_Sprite(Sprite).SurfDescription.lHeight
        .Left = (spriteleft * 3 + Anim) * (DDS_Sprite(Sprite).SurfDescription.lWidth / 12)
        .Right = .Left + (DDS_Sprite(Sprite).SurfDescription.lWidth / 12)
    End With
    
    With MapNpc(MapNpcNum, MapNum)
        ' Calculate X
        x = .x * PIC_X + .XOffset - ((DDS_Sprite(Sprite).SurfDescription.lWidth / 12 - 32) / 2)
        ' Is sprite more than 32..?
        If ((DDS_Sprite(Sprite).SurfDescription.lHeight) - 32) > 0 Then
            ' Create a 32 pixel offset for larger sprites
            y = MapNpc(MapNpcNum, MapNum).y * PIC_Y + MapNpc(MapNpcNum, MapNum).YOffset - ((DDS_Sprite(Sprite).SurfDescription.lHeight) - 32)
        Else
            ' Proceed as normal
            y = MapNpc(MapNpcNum, MapNum).y * PIC_Y + MapNpc(MapNpcNum, MapNum).YOffset
        End If
    End With
    
    ' Check if its out of bounds because of the offset
    ' Is player's Y less than 0..?
    If y < 0 Then
        With rec
            .Top = .Top - y
        End With
        y = 0
    End If

    ' Is player's X less than 0..?
    If x < 0 Then
        With rec
            .Left = .Left + (x * -1)
            '.Right = .Left + 48 - (x * -1)
        End With
        x = 0
    End If

    ' Is player's X more than max map values..?
    If x + (DDS_Sprite(Sprite).SurfDescription.lWidth / 12) > MAX_MAPX * 32 + 32 Then
        With rec
            .Right = .Right + (x - (MAX_MAPX * 32))
        End With
    End If
        
    Call BltSprite(Sprite, x2 * PIC_X + x + StaticX, y2 * PIC_Y + y + StaticY, rec)
    
    ' ** Blit Spells Animations **
    For i = 1 To MAX_SPELLANIM
        spellnum = MapNpc(MapNpcNum, MapNum).SpellAnimations(i).spellnum
        
        If spellnum > 0 Then
            If Spell(spellnum).Pic > 0 Then
            
                If MapNpc(MapNpcNum, MapNum).SpellAnimations(i).Timer < GetTickCount Then
                    MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer = MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer + 1
                    MapNpc(MapNpcNum, MapNum).SpellAnimations(i).Timer = GetTickCount + 120
                    
                    If MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer >= DDS_Spell(Spell(spellnum).Pic).SurfDescription.lWidth \ SIZE_X Then
                        MapNpc(MapNpcNum, MapNum).SpellAnimations(i).spellnum = 0
                        MapNpc(MapNpcNum, MapNum).SpellAnimations(i).Timer = 0
                        MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer = 0
                    End If
                End If
            
                If MapNpc(MapNpcNum, MapNum).SpellAnimations(i).spellnum > 0 Then
                    With rec
                        .Top = 0
                        .Bottom = SIZE_Y
                        .Left = MapNpc(MapNpcNum, MapNum).SpellAnimations(i).FramePointer * SIZE_X
                        .Right = .Left + SIZE_X
                    End With
            
                    Call BltSpell(Spell(spellnum).Pic, x + x2 * PIC_X, y + y2 * PIC_Y, rec)
                End If
                
            End If
        End If
        
    Next
    End If
End Sub

' ******************
' ** Game Editors **
' ******************

Public Sub BltMapEditor()
Dim Height As Long
Dim Width As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    Height = DDS_Tile.SurfDescription.lHeight
    Width = DDS_Tile.SurfDescription.lWidth
    
    dRECT.Top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width
    
    frmMainGame.picBackSelect.Height = Height
    frmMainGame.picBackSelect.Width = Width
   
    Call DD_BltToDC(DDS_Tile.Surface, sRECT, dRECT, frmMainGame.picBackSelect)
End Sub

Public Sub BltMapEditorTilePreview()
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    sRECT.Top = EditorTileY * PIC_Y
    sRECT.Bottom = sRECT.Top + PIC_Y
    sRECT.Left = EditorTileX * PIC_X
    sRECT.Right = sRECT.Left + PIC_X
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    
    Call DD_BltToDC(DDS_Tile.Surface, sRECT, dRECT, frmMainGame.picSelect)
End Sub

Public Sub BltTileOutline()
Dim x As Long
Dim y As Long
Dim rec As DXVBLib.RECT

    If Not isInBounds Then Exit Sub
    
    x = CurX
    y = CurY
    
    Call DD_BltFast(MapTilePosition(x, y).PosX, MapTilePosition(x, y).PosY, DDS_Misc.Surface, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub MapItemEditorBltItem()
Dim ItemNum As Integer
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    ItemNum = Item(frmMapItem.scrlItem.Value).Pic

    If ItemNum < 1 Or ItemNum > NumItems Then
        frmMapItem.picPreview.Cls
        Exit Sub
    End If
    
    Call DD_ReadySurface("Items\" & ItemNum, DDS_Item(ItemNum))

    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call DD_BltToDC(DDS_Item(ItemNum).Surface, sRECT, dRECT, frmMapItem.picPreview)
    
End Sub

Public Sub BltInventory(ItemNum As Integer)
Dim PicNum As Integer
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    PicNum = Item(ItemNum).Pic

    If PicNum < 1 Or PicNum > NumItems Then
        frmMainGame.picInvSelected.Cls
        Exit Sub
    End If

    Call DD_ReadySurface("Items\" & PicNum, DDS_Item(PicNum))
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call DD_BltToDC(DDS_Item(PicNum).Surface, sRECT, dRECT, frmMainGame.picInvSelected)
    
End Sub

Public Sub KeyItemEditorBltItem()
Dim ItemNum As Integer
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    ItemNum = Item(frmMapKey.scrlItem.Value).Pic
    
    If ItemNum < 1 Or ItemNum > NumItems Then
        frmMapKey.picPreview.Cls
        Exit Sub
    End If
    
    Call DD_ReadySurface("Items\" & ItemNum, DDS_Item(ItemNum))

    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call DD_BltToDC(DDS_Item(ItemNum).Surface, sRECT, dRECT, frmMapKey.picPreview)
    
End Sub

Public Sub ItemEditorBltItem()
Dim ItemPic As Integer
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    ItemPic = frmItemEditor.scrlPic.Value
    
    If ItemPic < 1 Or ItemPic > NumItems Then
        frmItemEditor.picPic.Cls
        Exit Sub
    End If
    
    Call DD_ReadySurface("Items\" & ItemPic, DDS_Item(ItemPic))

    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call DD_BltToDC(DDS_Item(ItemPic).Surface, sRECT, dRECT, frmItemEditor.picPic)
    
End Sub

Public Sub NpcEditorBltSprite()
Dim SpritePic As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    SpritePic = frmNpcEditor.scrlSprite.Value

    If SpritePic < 1 Or SpritePic > NumSprites Then
        frmNpcEditor.picSprite.Cls
        Exit Sub
    End If
    
    Call DD_ReadySurface("sprites\" & SpritePic, DDS_Sprite(SpritePic))
    
    sRECT.Top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = PIC_X * 3 ' facing down
    sRECT.Right = sRECT.Left + SIZE_X
    
    dRECT.Top = 0
    dRECT.Bottom = SIZE_Y
    dRECT.Left = 0
    dRECT.Right = SIZE_X

    Call DD_BltToDC(DDS_Sprite(SpritePic).Surface, sRECT, dRECT, frmNpcEditor.picSprite)

End Sub

Public Sub SpellEditorBltSpell()
Dim SpellPic As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    SpellPic = frmSpellEditor.scrlPic.Value

    If SpellPic < 1 Or SpellPic > NumSpells Then
        frmSpellEditor.picPic.Cls
        Exit Sub
    End If
    
    Call DD_ReadySurface("spells\" & SpellPic, DDS_Spell(SpellPic))
    
    sRECT.Top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = frmSpellEditor.scrlFrame.Value * SIZE_X
    sRECT.Right = sRECT.Left + SIZE_X
    
    dRECT.Top = 0
    dRECT.Bottom = SIZE_Y
    dRECT.Left = 0
    dRECT.Right = SIZE_X

    Call DD_BltToDC(DDS_Spell(SpellPic).Surface, sRECT, dRECT, frmSpellEditor.picPic)

End Sub

' *************************

Public Sub Render_Graphics()
'On Error GoTo ErrorHandle
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim n As Long
    Dim rec As DXVBLib.RECT
    Dim rec_pos As DXVBLib.RECT
    
    If Not CheckSurfaces Then Exit Sub

    If frmMainGame.WindowState = vbMinimized Then Exit Sub
    
    Call DDS_BackBuffer.BltColorFill(rec, 0)
 
    If GettingMap Then
        TexthDC = DDS_BackBuffer.GetDC ' Lock backbuffer to draw text
        Call DrawText(TexthDC, 35, 35, "Receiving Map...", QBColor(BrightCyan))
    Else
    
        ' blit lower tiles
        For x = -1 To MAX_MAPX + 1
            For y = -1 To MAX_MAPY + 1
                Call BltMapTile(x, y)
            Next
        Next
    
        ' Blit out the items
        For n = 1 To 9
            If (tMap(n) <> 0) And (tMap(n) < MAX_MAPS + 1) Then
                For i = 1 To MAX_MAP_ITEMS
                    If MapItem(i, tMap(n)).Num > 0 Then
                        Call BltItem(i, tMap(n))
                    End If
                Next
            End If
        Next
        
        ' Blit out players and NPCs
        For y = -1 To MAX_MAPY + 1
            For i = 1 To High_Index
                If Player(i).y = y Then
                    Call BltPlayer(i)
                End If
            Next
            
            For n = 1 To 9
                If (tMap(n) <> 0) And (tMap(n) < MAX_MAPS + 1) Then
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(i, tMap(n)).y = y Then
                            Call BltNpc(i, tMap(n))
                        End If
                    Next
                End If
            Next
        Next
        
        
        ' blit out upper tiles
        For x = -1 To MAX_MAPX + 1
            For y = -1 To MAX_MAPY + 1
                Call BltMapFringeTile(x, y)
            Next
        Next
     
        ' blit out a square at mouse cursor
        If Editor = EDITOR_MAP Then
            Call BltTileOutline
        End If
        
        ' ********************
        ' *** TEXT DRAWING ***
        ' ********************
        ' Lock backbuffer to draw text
        TexthDC = DDS_BackBuffer.GetDC
    
        ' draw player names
        For i = 1 To High_Index
            Call DrawPlayerName(i)
            Call DrawPlayerGuildName(i)
        Next
        
        ' Draw map name
        Call DrawText(TexthDC, DrawMapNameX, DrawMapNameY, map(5).Name, DrawMapNameColor)
        
       ' draw FPS
        If BFPS Then
            Call DrawText(TexthDC, (MAX_MAPX - 1) * PIC_X - 4, 1, Trim$("FPS: " & GameFPS), QBColor(Yellow))
        End If
    
        ' draw cursor and player location
        If BLoc Then
            Call DrawText(TexthDC, 0, 1, Trim$("Cur X: " & CurX & " Y: " & CurY), QBColor(Yellow))
            Call DrawText(TexthDC, 0, 15, Trim$("Loc X: " & GetPlayerX(MyIndex) & " Y: " & GetPlayerY(MyIndex)), QBColor(Yellow))
            Call DrawText(TexthDC, 0, 27, Trim$(" (Map #" & GetPlayerMap(MyIndex) & ")"), QBColor(Yellow))
        End If
                        
        ' draw map attributes
        If Editor = EDITOR_MAP Then
            If frmMainGame.optAttribs.Value Then
                Call DrawMapAttributes
            End If
        End If
        
    End If
    
    ' *********************
    ' *********************
        
    ' Release DC
    Call DDS_BackBuffer.ReleaseDC(TexthDC)
    
    ' Get the rect to blit to
    Call DX7.GetWindowRect(frmMainGame.picScreen.hwnd, rec_pos)

    'Call Pixelate(Pixelation)
    ' Blit the backbuffer
    Call DDS_Primary.Blt(rec_pos, DDS_BackBuffer, rec, DDBLT_WAIT)

    Exit Sub
    
ErrorHandle:

    Select Case Err
    
        Case 91
            Sleep 100
            Call ReInitDD
            Err.Clear
            Exit Sub
        
        Case Else
            Call DevMsg("(Render_Graphics) Error Rendering Graphics - Unhandled Error", BrightRed)
            Call DevMsg("(Render_Graphics) Error Number : " & Err & " - " & Err.Description, BrightRed)
    
    End Select
        
End Sub

Private Sub Pixelate(ByVal Subs As Long)
Dim srcRect As RECT
Dim destRect As RECT

If Subs <> 0 Then

    srcRect.Bottom = (MAX_MAPY + 1) * PIC_Y
    srcRect.Right = (MAX_MAPX + 1) * PIC_X
    
    destRect.Bottom = srcRect.Bottom / (Subs + 1)
    destRect.Right = destRect.Bottom * (((MAX_MAPY + 1) * PIC_Y) / ((MAX_MAPX + 1) * PIC_X))   'srcRect.Right / (Subs + 1)
    
    Call DDS_BackBuffer.Blt(destRect, DDS_BackBuffer, srcRect, DDBLT_WAIT)
    Call DDS_BackBuffer.Blt(srcRect, DDS_BackBuffer, destRect, DDBLT_WAIT)
    
End If

End Sub

Public Sub DrawSelChar(ByVal i As Long)
Dim rec As RECT
Dim rec_pos As RECT

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 3 * PIC_X
        .Right = .Left + PIC_X
    End With

    With rec_pos
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
    
    If CharSprites(i) = 0 Then
        Exit Sub
    End If
    
    Call DD_ReadySurface("sprites\" & CharSprites(i), DDS_Sprite(CharSprites(i)))

    DD_BltToDC DDS_Sprite(CharSprites(i)).Surface, rec, rec_pos, frmMainMenu.picSelChar
End Sub

Public Sub DrawNewChar()
Dim rec As RECT
Dim rec_pos As RECT
Dim ListIndexSprite As Long

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 3 * PIC_X
        .Right = .Left + PIC_X
    End With

    With rec_pos
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
   
    ListIndexSprite = frmMainMenu.lstChars.ListIndex
   
    If Class(ListIndexSprite).Sprite = 0 Then Exit Sub
   
    Call DD_ReadySurface("sprites\" & Class(ListIndexSprite).Sprite, DDS_Sprite(Class(ListIndexSprite).Sprite))
   
    DD_BltToDC DDS_Sprite(CharSprites(Class(ListIndexSprite).Sprite)).Surface, rec, rec_pos, frmMainMenu.picPic

End Sub
