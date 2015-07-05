Attribute VB_Name = "modDirectDraw7"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ** Renders graphics                     **
' ******************************************

Public DD As DirectDraw7 ' DirectDraw7 Object

Public DD_Clip As DirectDrawClipper ' Clipper object

' primary surface
Public DDS_Primary As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

' back buffer
Public DDS_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

' gfx buffers
Public DDS_Item() As DirectDrawSurface7
Public DDS_Sprite() As DirectDrawSurface7
Public DDS_Spell() As DirectDrawSurface7
Public DDS_Tile As DirectDrawSurface7
Public DDS_Misc As DirectDrawSurface7

Public DDSD_Item() As DDSURFACEDESC2
Public DDSD_Sprite() As DDSURFACEDESC2
Public DDSD_Spell() As DDSURFACEDESC2
Public DDSD_Tile As DDSURFACEDESC2
Public DDSD_Misc As DDSURFACEDESC2

Public DDSD_Temp As DDSURFACEDESC2

Public Const SurfaceTimerMax As Long = 200000
Public SpriteTimer() As Long
Public SpellTimer() As Long
Public ItemTimer() As Long

' Number of graphic files
Public NumTileSets As Long
Public NumSprites As Long
Public NumSpells As Long
Public NumItems As Long

' ** These values are used for transparency, comment out if getting mask color from the bitmap
' ** MASK_COLOR is the "transparency" color
'Public Const MASK_COLOR As Long = 0   ' RGB(0,0,0)
'Public Key As DDCOLORKEY

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
    Call DD.SetCooperativeLevel(frmMirage.hWnd, DDSCL_NORMAL)
        
    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
        
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    Call DD_Clip.SetHWnd(frmMirage.picScreen.hWnd)
        
    ' Have the blits to the screen clipped to the picture box
    Call DDS_Primary.SetClipper(DD_Clip) ' method attaches a clipper object to, or deletes one from, a surface.
    
    Call InitBackBuffer
    
    InitDirectDraw = True
    
    Exit Function
    
ErrorHandle:

    Select Case Err.Number
    
        Case 91
            Call MsgBox("DirectX7 master object not created.")

    End Select
    
    InitDirectDraw = False
    
End Function

Private Sub InitBackBuffer()
Dim rec As DXVBLib.RECT
  
    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (MAX_MAPX + 3) * PIC_X
        .lHeight = (MAX_MAPY + 3) * PIC_Y
    End With

    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)

    Call DDS_BackBuffer.BltColorFill(rec, 0)

End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
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

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(FileName As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)
On Error GoTo ErrorHandle
    
    ' Set path
    FileName = App.Path & GFX_PATH & FileName & GFX_EXT
    
    ' Destroy surface if it exist
    If Not Surf Is Nothing Then
        Set Surf = Nothing
        Call ZeroMemory(ByVal VarPtr(SurfDesc), LenB(SurfDesc))
    End If
    
    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
    
    ' init object
    Set Surf = DD.CreateSurfaceFromFile(FileName, SurfDesc)
    
    'Call Surf.SetColorKey(DDCKEY_SRCBLT, Key) ' MASK_COLOR
    Call SetMaskColorFromPixel(Surf, 0, 0)
    
    Exit Sub
    
ErrorHandle:

    Select Case Err.Number
        
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
            MsgBox "Out of video memory, attempting to re-initialize using system memory"
            
            DDSD_Temp.lFlags = DDSD_CAPS
            DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY

            Call ReInitDD
                  
    End Select
    
End Sub

Public Function CheckSurfaces() As Boolean
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

Public Sub InitTileSurf(ByVal TileSet As Integer)
    If TileSet < 1 Or TileSet > NumTileSets Then Exit Sub
    
    Call InitDDSurf("tiles" & TileSet, DDSD_Tile, DDS_Tile)
End Sub

Public Sub ReInitDD()
    Call InitDirectDraw
    Call InitTileSurf(Map.TileSet)
    
    If InMapEditor Then
        Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    End If
    
End Sub

Public Sub DestroyDirectDraw()
Dim i As Long

    ' Unload DirectDraw
    Set DDS_Misc = Nothing
    Set DDS_Tile = Nothing
    
    For i = 1 To NumItems
        Set DDS_Item(i) = Nothing
    Next

    For i = 1 To NumSpells
        Set DDS_Spell(i) = Nothing
    Next

    For i = 1 To NumSprites
        Set DDS_Sprite(i) = Nothing
    Next

    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    
    Set DD_Clip = Nothing
    Set DD = Nothing

End Sub

' **************
' ** Blitting **
' **************

Public Sub Engine_BltFast(ByVal dx As Long, ByVal dy As Long, ByRef ddS As DirectDrawSurface7, srcRECT As RECT, trans As CONST_DDBLTFASTFLAGS)
On Error GoTo ErrorHandle:
    
    If Not ddS Is Nothing Then
        Call DDS_BackBuffer.BltFast(ConvertMapX(dx), ConvertMapY(dy), ddS, srcRECT, trans)
    End If
    
    Exit Sub

ErrorHandle:

    Select Case Err.Number
    
        Case 5
            Call DevMsg("Attempting to copy from a surface thats not initialized.", BrightRed)
    
    End Select
    
End Sub

Public Function Engine_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DXVBLib.RECT, dRECT As DXVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
On Error GoTo ErrorHandle
    
    If Clear Then
        picBox.Cls
    End If
    
    Call Surface.BltToDC(picBox.hdc, sRECT, dRECT)
    picBox.Refresh

    Engine_BltToDC = True
    Exit Function

ErrorHandle:
    ' returns false on error
    Engine_BltToDC = False

End Function

Public Sub BltMapTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DXVBLib.RECT

    With Map.Tile(X, Y)
    
        rec.Top = (.Ground \ TILESHEET_WIDTH) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (.Ground Mod TILESHEET_WIDTH) * PIC_X
        rec.Right = rec.Left + PIC_X
        Call Engine_BltFast(X * PIC_X, Y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT)
    
        If MapAnim = 0 Or .Anim <= 0 Then
            If .Mask > 0 Then
                If TempTile(X, Y).DoorOpen = NO Then
                    rec.Top = (.Mask \ TILESHEET_WIDTH) * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    rec.Left = (.Mask Mod TILESHEET_WIDTH) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    Call Engine_BltFast(X * PIC_X, Y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If
        Else
            ' Is there an animation tile to draw?
            If .Anim > 0 Then
                rec.Top = (.Anim \ TILESHEET_WIDTH) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = (.Anim Mod TILESHEET_WIDTH) * PIC_X
                rec.Right = rec.Left + PIC_X
                Call Engine_BltFast(X * PIC_X, Y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
        
        If .Mask2 > 0 Then
            If TempTile(X, Y).DoorOpen = NO Then
                rec.Top = (.Mask2 \ TILESHEET_WIDTH) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = (.Mask2 Mod TILESHEET_WIDTH) * PIC_X
                rec.Right = rec.Left + PIC_X
                Call Engine_BltFast(X * PIC_X, Y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    
    End With

End Sub

Public Sub BltMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DXVBLib.RECT

    With Map.Tile(X, Y)
        If .Fringe > 0 Then
            rec.Top = (.Fringe \ TILESHEET_WIDTH) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (.Fringe Mod TILESHEET_WIDTH) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call Engine_BltFast(X * PIC_X, Y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        If .Fringe2 > 0 Then
            rec.Top = (.Fringe2 \ TILESHEET_WIDTH) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (.Fringe2 Mod TILESHEET_WIDTH) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call Engine_BltFast(X * PIC_X, Y * PIC_Y, DDS_Tile, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End With
    
End Sub

Public Sub BltItem(ByVal ItemNum As Long)
Dim PicNum As Integer
Dim rec As DXVBLib.RECT

    PicNum = Item(MapItem(ItemNum).Num).Pic
    
    If PicNum < 1 Or PicNum > NumItems Then Exit Sub

    ItemTimer(PicNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(PicNum) Is Nothing Then
        Call InitDDSurf("Items\" & PicNum, DDSD_Item(PicNum), DDS_Item(PicNum))
    End If

    With rec
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
    
    Call Engine_BltFast(MapItem(ItemNum).X * PIC_X, MapItem(ItemNum).Y * PIC_Y, DDS_Item(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Private Sub BltSprite(ByVal Sprite As Long, ByVal X As Long, Y As Long, rec As DXVBLib.RECT)
    If Sprite < 1 Or Sprite > NumSprites Then Exit Sub
    
    SpriteTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Sprite(Sprite) Is Nothing Then
        Call InitDDSurf("sprites\" & Sprite, DDSD_Sprite(Sprite), DDS_Sprite(Sprite))
    End If
    
    Call Engine_BltFast(X, Y, DDS_Sprite(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Private Sub BltSpell(ByVal SpellNum As Long, ByVal X As Long, Y As Long, rec As DXVBLib.RECT)
    If SpellNum < 1 Or SpellNum > NumSpells Then Exit Sub
    
    SpellTimer(SpellNum) = GetTickCount + SurfaceTimerMax

    If DDS_Spell(SpellNum) Is Nothing Then
        Call InitDDSurf("Spells\" & SpellNum, DDSD_Spell(SpellNum), DDS_Spell(SpellNum))
    End If
    
    Call Engine_BltFast(X, Y, DDS_Spell(SpellNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim i As Long
Dim SpellNum As Long
Dim X As Long
Dim Y As Long
Dim Sprite As Long
Dim rec As DXVBLib.RECT

    Sprite = GetPlayerSprite(Index)

    ' Check for animation
    Anim = 0
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset < SIZE_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).YOffset < SIZE_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).XOffset < SIZE_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).XOffset < SIZE_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If

    
    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .Top = 0
        .Bottom = SIZE_Y
        .Left = (GetPlayerDir(Index) * 3 + Anim) * SIZE_X
        .Right = .Left + SIZE_X
    End With
    
    X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
    Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4 ' to raise the sprite by 4 pixels
    
    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        With rec
            .Top = .Top + (Y * -1)
        End With
    End If
        
    Call BltSprite(Sprite, X, Y, rec)
    
    ' ** Blit Spells Animations **
    For i = 1 To MAX_SPELLANIM
        SpellNum = Player(Index).SpellAnimations(i).SpellNum
        
        If SpellNum > 0 Then
        
            If Player(Index).SpellAnimations(i).Timer < GetTickCount Then
            
                Player(Index).SpellAnimations(i).FramePointer = Player(Index).SpellAnimations(i).FramePointer + 1
                Player(Index).SpellAnimations(i).Timer = GetTickCount + 120
                
                If Player(Index).SpellAnimations(i).FramePointer >= DDSD_Spell(Spell(SpellNum).Pic).lWidth \ SIZE_X Then
                    Player(Index).SpellAnimations(i).SpellNum = 0
                    Player(Index).SpellAnimations(i).Timer = 0
                    Player(Index).SpellAnimations(i).FramePointer = 0
                End If

            End If
        
            If Player(Index).SpellAnimations(i).SpellNum > 0 Then
                With rec
                    .Top = 0
                    .Bottom = SIZE_Y
                    .Left = Player(Index).SpellAnimations(i).FramePointer * SIZE_X
                    .Right = .Left + SIZE_X
                End With
        
                Call BltSpell(Spell(SpellNum).Pic, X, Y, rec)
            End If
            
        End If
    Next


End Sub

Public Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim i As Long
Dim SpellNum As Long
Dim X As Long
Dim Y As Long
Dim Sprite As Long
Dim rec As DXVBLib.RECT

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    Sprite = Npc(MapNpc(MapNpcNum).Num).Sprite

    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < SIZE_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < SIZE_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < SIZE_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < SIZE_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .Top = 0
        .Bottom = SIZE_Y
        .Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * SIZE_X
        .Right = .Left + SIZE_X
    End With
    
    With MapNpc(MapNpcNum)
        X = .X * PIC_X + .XOffset
        Y = .Y * PIC_Y + .YOffset - 4 ' to raise the sprite by 4 pixels
    End With
    
    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        rec.Top = rec.Top + (Y * -1)
    End If
        
    Call BltSprite(Sprite, X, Y, rec)
    
    ' ** Blit Spells Animations **
    For i = 1 To MAX_SPELLANIM
        SpellNum = MapNpc(MapNpcNum).SpellAnimations(i).SpellNum
        
        If SpellNum > 0 Then
        
            If MapNpc(MapNpcNum).SpellAnimations(i).Timer < GetTickCount Then
            
                MapNpc(MapNpcNum).SpellAnimations(i).FramePointer = MapNpc(MapNpcNum).SpellAnimations(i).FramePointer + 1
                MapNpc(MapNpcNum).SpellAnimations(i).Timer = GetTickCount + 120
                
                If MapNpc(MapNpcNum).SpellAnimations(i).FramePointer >= DDSD_Spell(Spell(SpellNum).Pic).lWidth \ SIZE_X Then
                    MapNpc(MapNpcNum).SpellAnimations(i).SpellNum = 0
                    MapNpc(MapNpcNum).SpellAnimations(i).Timer = 0
                    MapNpc(MapNpcNum).SpellAnimations(i).FramePointer = 0
                End If

            End If
        
            If MapNpc(MapNpcNum).SpellAnimations(i).SpellNum > 0 Then
                With rec
                    .Top = 0
                    .Bottom = SIZE_Y
                    .Left = MapNpc(MapNpcNum).SpellAnimations(i).FramePointer * SIZE_X
                    .Right = .Left + SIZE_X
                End With
        
                Call BltSpell(Spell(SpellNum).Pic, X, Y, rec)
            End If
            
        End If
    Next
    
End Sub

' ******************
' ** Game Editors **
' ******************

Public Sub BltMapEditor()
Dim Height As Long
Dim Width As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    Height = DDSD_Tile.lHeight
    Width = DDSD_Tile.lWidth
    
    dRECT.Top = 0
    dRECT.Bottom = Height
    dRECT.Left = 0
    dRECT.Right = Width
    
    frmMirage.picBackSelect.Height = Height
    frmMirage.picBackSelect.Width = Width
   
    Call Engine_BltToDC(DDS_Tile, sRECT, dRECT, frmMirage.picBackSelect)
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
    
    Call Engine_BltToDC(DDS_Tile, sRECT, dRECT, frmMirage.picSelect)
End Sub

Public Sub BltTileOutline()
Dim rec As DXVBLib.RECT
    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With
    
    Call Engine_BltFast(CurX * PIC_X, CurY * PIC_Y, DDS_Misc, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
    
    ItemTimer(ItemNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(ItemNum) Is Nothing Then
        Call InitDDSurf("Items\" & ItemNum, DDSD_Item(ItemNum), DDS_Item(ItemNum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call Engine_BltToDC(DDS_Item(ItemNum), sRECT, dRECT, frmMapItem.picPreview)
    
End Sub

Public Sub BltInventory(ItemNum As Integer)
Dim PicNum As Integer
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Sub

     PicNum = Item(ItemNum).Pic
    
    If PicNum < 1 Or PicNum > NumItems Then
        frmMirage.picInvSelected.Cls
        Exit Sub
    End If
    
    ItemTimer(PicNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(PicNum) Is Nothing Then
        Call InitDDSurf("Items\" & PicNum, DDSD_Item(PicNum), DDS_Item(PicNum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call Engine_BltToDC(DDS_Item(PicNum), sRECT, dRECT, frmMirage.picInvSelected)
    
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
    
    ItemTimer(ItemNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(ItemNum) Is Nothing Then
        Call InitDDSurf("Items\" & ItemNum, DDSD_Item(ItemNum), DDS_Item(ItemNum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call Engine_BltToDC(DDS_Item(ItemNum), sRECT, dRECT, frmMapKey.picPreview)
    
End Sub

Public Sub ItemEditorBltItem()
Dim ItemNum As Integer
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    ItemNum = frmItemEditor.scrlPic.Value
    
    If ItemNum < 1 Or ItemNum > NumItems Then
        frmItemEditor.picPic.Cls
        Exit Sub
    End If

    ItemTimer(ItemNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(ItemNum) Is Nothing Then
        Call InitDDSurf("Items\" & ItemNum, DDSD_Item(ItemNum), DDS_Item(ItemNum))
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    dRECT.Top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X

    Call Engine_BltToDC(DDS_Item(ItemNum), sRECT, dRECT, frmItemEditor.picPic)
    
End Sub

Public Sub NpcEditorBltSprite()
Dim Sprite As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    Sprite = frmNpcEditor.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumSprites Then
        frmNpcEditor.picSprite.Cls
        Exit Sub
    End If
    
    SpriteTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Sprite(Sprite) Is Nothing Then
        Call InitDDSurf("sprites\" & Sprite, DDSD_Sprite(Sprite), DDS_Sprite(Sprite))
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = PIC_X * 3 ' facing down
    sRECT.Right = sRECT.Left + SIZE_X
    
    dRECT.Top = 0
    dRECT.Bottom = SIZE_Y
    dRECT.Left = 0
    dRECT.Right = SIZE_X

    Call Engine_BltToDC(DDS_Sprite(Sprite), sRECT, dRECT, frmNpcEditor.picSprite)

End Sub

Public Sub SpellEditorBltSpell()
Dim SpellNum As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    SpellNum = frmSpellEditor.scrlPic.Value

    If SpellNum < 1 Or SpellNum > NumSpells Then
        frmSpellEditor.picPic.Cls
        Exit Sub
    End If
    
    SpellTimer(SpellNum) = GetTickCount + SurfaceTimerMax

    If DDS_Spell(SpellNum) Is Nothing Then
        Call InitDDSurf("spells\" & SpellNum, DDSD_Spell(SpellNum), DDS_Spell(SpellNum))
    End If
    
    frmSpellEditor.scrlFrame.Max = (DDSD_Spell(SpellNum).lWidth \ SIZE_X) - 1
    
    sRECT.Top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = frmSpellEditor.scrlFrame.Value * SIZE_X
    sRECT.Right = sRECT.Left + SIZE_X
    
    dRECT.Top = 0
    dRECT.Bottom = SIZE_Y
    dRECT.Left = 0
    dRECT.Right = SIZE_X

    Call Engine_BltToDC(DDS_Spell(SpellNum), sRECT, dRECT, frmSpellEditor.picPic)

End Sub

Public Sub Render_Graphics()
On Error GoTo ErrorHandle
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim rec As DXVBLib.RECT
    Dim rec_pos As DXVBLib.RECT
    
    If Not CheckSurfaces Then Exit Sub

    If frmMirage.WindowState = vbMinimized Then Exit Sub
 
    If GettingMap Then
        
        TexthDC = DDS_BackBuffer.GetDC ' Lock the backbuffer so we can draw text and names
        
        ' Check if we are getting a map, and if we are tell them so
        Call DrawText(TexthDC, 35, 35, "Receiving Map...", QBColor(BrightCyan))
        
    Else
    
        UpdateCamera
        
        DDS_BackBuffer.BltColorFill rec_pos, 0
        
        ' blit lower tiles
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapTile(X, Y)
                End If
            Next
        Next
    
        ' Blit out the items
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call BltItem(i)
            End If
        Next
        
        ' Blit out players
        For i = 1 To PlayersOnMapHighIndex
            Call BltPlayer(PlayersOnMap(i))
        Next
                
        ' Blit out the npcs
        For i = 1 To High_Npc_Index
            Call BltNpc(i)
        Next
        
        ' blit out upper tiles
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapFringeTile(X, Y)
                End If
            Next
        Next
     
        ' blit out a square at mouse cursor
        If InMapEditor Then
            Call BltTileOutline
        End If
    
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DDS_BackBuffer.GetDC
        
        ' draw FPS
        If BFPS Then
            Call DrawText(TexthDC, Camera.Right - (Len("FPS: " & GameFPS) * 8), Camera.Top + 1, Trim$("FPS: " & GameFPS), QBColor(Yellow))
        End If
    
        ' draw cursor, player X and Y locations
        If BLoc Then
            Call DrawText(TexthDC, Camera.Left, Camera.Top + 1, Trim$("cur x: " & CurX & " y: " & CurY), QBColor(Yellow))
            Call DrawText(TexthDC, Camera.Left, Camera.Top + 15, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), QBColor(Yellow))
            Call DrawText(TexthDC, Camera.Left, Camera.Top + 27, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), QBColor(Yellow))
        End If
    
        ' draw player names
        For i = 1 To PlayersOnMapHighIndex
            Call DrawPlayerName(PlayersOnMap(i))
        Next
                        
        ' Blit out map attributes
        If InMapEditor Then
            Call BltMapAttributes
        End If

        ' Draw map name
        Call DrawText(TexthDC, DrawMapNameX, DrawMapNameY, Map.Name, DrawMapNameColor)
    End If
        
Continue:

    ' Release DC
    Call DDS_BackBuffer.ReleaseDC(TexthDC)
    
    ' Get the rect to blit to
    Call DX7.GetWindowRect(frmMirage.picScreen.hWnd, rec_pos)

    ' Blit the backbuffer
    Call DDS_Primary.Blt(rec_pos, DDS_BackBuffer, Camera, DDBLT_WAIT)

    Exit Sub
    
ErrorHandle:
    
    If Err.Number = 91 Then
        Sleep 100
        Call ReInitDD
        Err.Clear
        Exit Sub
    End If
    
On Error Resume Next

    If Not CheckSurfaces Then Exit Sub ' surfaces can get lost, check again
    
    TexthDC = DDS_BackBuffer.GetDC ' Lock the backbuffer so we can draw text and names
    Call DrawText(TexthDC, 10, 15, "Error Rendering Graphics - Unhandled Error", QBColor(BrightRed))
    Call DrawText(TexthDC, 10, 28, "Error Number : " & Err.Number & " - " & Err.Description, QBColor(BrightCyan))
    
    GoTo Continue
            
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

    StartX = GetPlayerX(MyIndex) - ((MAX_MAPX + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((MAX_MAPY + 1) \ 2) - 1
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
    
    EndX = StartX + (MAX_MAPX + 1) + 1
    EndY = StartY + (MAX_MAPY + 1) + 1
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
    ConvertMapX = X - (TileView.Left * PIC_X)
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ConvertMapY = Y - (TileView.Top * PIC_Y)
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
