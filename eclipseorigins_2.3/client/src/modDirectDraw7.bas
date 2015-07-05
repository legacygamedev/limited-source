Attribute VB_Name = "modDirectDraw7"
Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectDraw7 Object
Public DD As DirectDraw7
' Clipper object
Public DD_Clip As DirectDrawClipper

' primary surface
Public DDS_Primary As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

' back buffer
Public DDS_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

' Used for pre-rendering
Public DDS_Map As DirectDrawSurface7
Public DDSD_Map As DDSURFACEDESC2

' gfx buffers
Public DDS_Item() As DirectDrawSurface7 ' arrays
Public DDS_Character() As DirectDrawSurface7
Public DDS_Paperdoll() As DirectDrawSurface7
Public DDS_Tileset() As DirectDrawSurface7
Public DDS_Resource() As DirectDrawSurface7
Public DDS_Animation() As DirectDrawSurface7
Public DDS_SpellIcon() As DirectDrawSurface7
Public DDS_Face() As DirectDrawSurface7
Public DDS_Door As DirectDrawSurface7 ' singes
Public DDS_Blood As DirectDrawSurface7
Public DDS_Misc As DirectDrawSurface7
Public DDS_Direction As DirectDrawSurface7
Public DDS_Target As DirectDrawSurface7
Public DDS_Bars As DirectDrawSurface7
Public DDS_Event As DirectDrawSurface7

' descriptions
Public DDSD_Temp As DDSURFACEDESC2 ' arrays
Public DDSD_Item() As DDSURFACEDESC2
Public DDSD_Character() As DDSURFACEDESC2
Public DDSD_Paperdoll() As DDSURFACEDESC2
Public DDSD_Tileset() As DDSURFACEDESC2
Public DDSD_Resource() As DDSURFACEDESC2
Public DDSD_Animation() As DDSURFACEDESC2
Public DDSD_SpellIcon() As DDSURFACEDESC2
Public DDSD_Face() As DDSURFACEDESC2
Public DDSD_Door As DDSURFACEDESC2 ' singles
Public DDSD_Blood As DDSURFACEDESC2
Public DDSD_Misc As DDSURFACEDESC2
Public DDSD_Direction As DDSURFACEDESC2
Public DDSD_Target As DDSURFACEDESC2
Public DDSD_Bars As DDSURFACEDESC2
Public DDSD_Event As DDSURFACEDESC2

' timers
Public Const SurfaceTimerMax As Long = 10000
Public CharacterTimer() As Long
Public PaperdollTimer() As Long
Public ItemTimer() As Long
Public ResourceTimer() As Long
Public AnimationTimer() As Long
Public SpellIconTimer() As Long
Public FaceTimer() As Long

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public NumItems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDirectDraw() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear DD7
    Call DestroyDirectDraw
    
    ' Init Direct Draw
    Set DD = DX7.DirectDrawCreate(vbNullString)
    
    ' Windowed
    DD.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL

    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        .lBackBufferCount = 1
    End With
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMain.picScreen.hWnd
    
    ' Have the blits to the screen clipped to the picture box
    DDS_Primary.SetClipper DD_Clip
    
    ' Initialise the surfaces
    InitSurfaces
    
    ' We're done
    InitDirectDraw = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub InitSurfaces()
Dim rec As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' DirectDraw Surface memory management setting
    DDSD_Temp.lFlags = DDSD_CAPS
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    
    ' clear out everything for re-init
    Set DDS_BackBuffer = Nothing

    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (MAX_MAPX + 3) * PIC_X
        .lHeight = (MAX_MAPY + 3) * PIC_Y
    End With
    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' load persistent surfaces
    If FileExist(App.Path & "\data files\graphics\door.bmp", True) Then Call InitDDSurf("door", DDSD_Door, DDS_Door)
    If FileExist(App.Path & "\data files\graphics\direction.bmp", True) Then Call InitDDSurf("direction", DDSD_Direction, DDS_Direction)
    If FileExist(App.Path & "\data files\graphics\target.bmp", True) Then Call InitDDSurf("target", DDSD_Target, DDS_Target)
    If FileExist(App.Path & "\data files\graphics\misc.bmp", True) Then Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    If FileExist(App.Path & "\data files\graphics\blood.bmp", True) Then Call InitDDSurf("blood", DDSD_Blood, DDS_Blood)
    If FileExist(App.Path & "\data files\graphics\bars.bmp", True) Then Call InitDDSurf("bars", DDSD_Bars, DDS_Bars)
    If FileExist(App.Path & "\data files\graphics\event.bmp", True) Then Call InitDDSurf("event", DDSD_Event, DDS_Event)
    
    ' count the blood sprites
    BloodCount = DDSD_Blood.lWidth / 32
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With TmpR
        .Left = X
        .top = Y
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
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetMaskColorFromPixel", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(fileName As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Set path
    fileName = App.Path & GFX_PATH & fileName & GFX_EXT

    ' Destroy surface if it exist
    If Not Surf Is Nothing Then
        Set Surf = Nothing
        Call ZeroMemory(ByVal VarPtr(SurfDesc), LenB(SurfDesc))
    End If

    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
    
    ' init object
    Set Surf = DD.CreateSurfaceFromFile(fileName, SurfDesc)
    
    ' Set mask
    Call SetMaskColorFromPixel(Surf, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitDDSurf", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function CheckSurfaces() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if we need to restore surfaces
    If Not DD.TestCooperativeLevel = DD_OK Then
        CheckSurfaces = False
    Else
        CheckSurfaces = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "CheckSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function NeedToRestoreSurfaces() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not DD.TestCooperativeLevel = DD_OK Then
        NeedToRestoreSurfaces = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "NeedToRestoreSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub ReInitDD()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call InitDirectDraw
    
    LoadTilesets
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ReInitDD", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyDirectDraw()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Unload DirectDraw
    Set DDS_Misc = Nothing
    
    For i = 1 To NumTileSets
        Set DDS_Tileset(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i))
    Next

    For i = 1 To NumItems
        Set DDS_Item(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i))
    Next

    For i = 1 To NumCharacters
        Set DDS_Character(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i))
    Next
    
    For i = 1 To NumPaperdolls
        Set DDS_Paperdoll(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i))
    Next
    
    For i = 1 To NumResources
        Set DDS_Resource(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i))
    Next
    
    For i = 1 To NumAnimations
        Set DDS_Animation(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i))
    Next
    
    For i = 1 To NumSpellIcons
        Set DDS_SpellIcon(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i))
    Next
    
    For i = 1 To NumFaces
        Set DDS_Face(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Face(i)), LenB(DDSD_Face(i))
    Next
    
    Set DDS_Blood = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Blood), LenB(DDSD_Blood)
    
    Set DDS_Door = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Door), LenB(DDSD_Door)
    
    Set DDS_Direction = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Direction), LenB(DDSD_Direction)
    
    Set DDS_Target = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Target), LenB(DDSD_Target)

    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    Set DD_Clip = Nothing
    Set DD = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Blitting **
' **************
Public Sub Engine_BltFast(ByVal dx As Long, ByVal dy As Long, ByRef ddS As DirectDrawSurface7, srcRECT As RECT, trans As CONST_DDBLTFASTFLAGS)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler


    If Not ddS Is Nothing Then
        Call DDS_BackBuffer.BltFast(dx, dy, ddS, srcRECT, trans)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Engine_BltFast", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Engine_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Clear Then
        picBox.Cls
    End If

    Call Surface.BltToDC(picBox.hDC, sRECT, dRECT)
    picBox.Refresh
    Engine_BltToDC = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Engine_BltToDC", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub BltDirection(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' render grid
    rec.top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.Bottom = rec.top + 32
    Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' render dir blobs
    For i = 1 To 4
        rec.Left = (i - 1) * 8
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(i)) Then
            rec.top = 8
        Else
            rec.top = 16
        End If
        rec.Bottom = rec.top + 8
        'render!
        Call Engine_BltFast(ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDirection", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTarget(ByVal X As Long, ByVal Y As Long)
Dim sRECT As DxVBLib.RECT
Dim width As Long, height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub
    
    width = DDSD_Target.lWidth / 2
    height = DDSD_Target.lHeight

    With sRECT
        .top = 0
        .Bottom = height
        .Left = 0
        .Right = width
    End With
    
    X = X - ((width - 32) / 2)
    Y = Y - (height / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTarget", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltHover(ByVal tType As Long, ByVal target As Long, ByVal X As Long, ByVal Y As Long)
Dim sRECT As DxVBLib.RECT
Dim width As Long, height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub
    
    width = DDSD_Target.lWidth / 2
    height = DDSD_Target.lHeight

    With sRECT
        .top = 0
        .Bottom = height
        .Left = width
        .Right = .Left + width
    End With
    
    X = X - ((width - 32) / 2)
    Y = Y - (height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHover", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile?
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                ' sort out rec
                rec.top = .Layer(i).Y * PIC_Y
                rec.Bottom = rec.top + PIC_Y
                rec.Left = .Layer(i).X * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "BltMapTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' skip tile if tileset isn't set
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                ' sort out rec
                rec.top = .Layer(i).Y * PIC_Y
                rec.Bottom = rec.top + PIC_Y
                rec.Left = .Layer(i).X * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapFringeTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltDoor(ByVal X As Long, ByVal Y As Long)
Dim rec As DxVBLib.RECT
Dim x2 As Long, y2 As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' sort out animation
    With TempTile(X, Y)
        If .DoorAnimate = 1 Then ' opening
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame < 4 Then
                    .DoorFrame = .DoorFrame + 1
                Else
                    .DoorAnimate = 2 ' set to closing
                End If
                .DoorTimer = GetTickCount
            End If
        ElseIf .DoorAnimate = 2 Then ' closing
            If .DoorTimer + 100 < GetTickCount Then
                If .DoorFrame > 1 Then
                    .DoorFrame = .DoorFrame - 1
                Else
                    .DoorAnimate = 0 ' end animation
                End If
                .DoorTimer = GetTickCount
            End If
        End If
        
        If .DoorFrame = 0 Then .DoorFrame = 1
    End With

    With rec
        .top = 0
        .Bottom = DDSD_Door.lHeight
        .Left = ((TempTile(X, Y).DoorFrame - 1) * (DDSD_Door.lWidth / 4))
        .Right = .Left + (DDSD_Door.lWidth / 4)
    End With

    x2 = (X * PIC_X)
    y2 = (Y * PIC_Y) - (DDSD_Door.lHeight / 2) + 4
    Call DDS_BackBuffer.BltFast(ConvertMapX(x2), ConvertMapY(y2), DDS_Door, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDoor", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltBlood(ByVal Index As Long)
Dim rec As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Blood(Index)
        ' check if we should be seeing it
        If .Timer + 20000 < GetTickCount Then Exit Sub
        
        rec.top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        Engine_BltFast ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), DDS_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBlood", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim i As Long
Dim width As Long, height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim X As Long, Y As Long
Dim lockindex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    AnimationTimer(Sprite) = GetTickCount + SurfaceTimerMax
    
    If DDS_Animation(Sprite) Is Nothing Then
        Call InitDDSurf("animations\" & Sprite, DDSD_Animation(Sprite), DDS_Animation(Sprite))
    End If
    
    ' total width divided by frame count
    width = DDSD_Animation(Sprite).lWidth / FrameCount
    height = DDSD_Animation(Sprite).lHeight
    
    sRECT.top = 0
    sRECT.Bottom = height
    sRECT.Left = (AnimInstance(Index).FrameIndex(Layer) - 1) * width
    sRECT.Right = sRECT.Left + width
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (width / 2) + Player(lockindex).XOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (height / 2) + Player(lockindex).YOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (width / 2) + MapNpc(lockindex).XOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (height / 2) + MapNpc(lockindex).YOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (height / 2)
    End If
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    ' Clip to screen
    If Y < 0 Then

        With sRECT
            .top = .top - Y
        End With

        Y = 0
    End If

    If X < 0 Then

        With sRECT
            .Left = .Left - X
        End With

        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    
    Call Engine_BltFast(X, Y, DDS_Animation(Sprite), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimation", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItem(ByVal itemnum As Long)
Dim PicNum As Long
Dim rec As DxVBLib.RECT
Dim MaxFrames As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if it's not us then don't render
    If MapItem(itemnum).playerName <> vbNullString Then
        If MapItem(itemnum).playerName <> Trim$(GetPlayerName(MyIndex)) Then Exit Sub
    End If
    
    ' get the picture
    PicNum = Item(MapItem(itemnum).num).Pic

    If PicNum < 1 Or PicNum > NumItems Then Exit Sub
    ItemTimer(PicNum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(PicNum) Is Nothing Then
        Call InitDDSurf("items\" & PicNum, DDSD_Item(PicNum), DDS_Item(PicNum))
    End If

    If DDSD_Item(PicNum).lWidth > 64 Then ' has more than 1 frame
        With rec
            .top = 0
            .Bottom = 32
            .Left = (MapItem(itemnum).Frame * 32)
            .Right = .Left + 32
        End With
    Else
        With rec
            .top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If

    Call Engine_BltFast(ConvertMapX(MapItem(itemnum).X * PIC_X), ConvertMapY(MapItem(itemnum).Y * PIC_Y), DDS_Item(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ScreenshotMap()
Dim X As Long, Y As Long, i As Long, rec As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' clear the surface
    Set DDS_Map = Nothing
    
    ' Initialize it
    With DDSD_Map
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (Map.MaxX + 1) * 32
        .lHeight = (Map.MaxY + 1) * 32
    End With
    Set DDS_Map = DD.CreateSurface(DDSD_Map)
    
    ' render the tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For i = MapLayer.Ground To MapLayer.Mask2
                    ' skip tile?
                    If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                        ' sort out rec
                        rec.top = .Layer(i).Y * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        rec.Left = .Layer(i).X * PIC_X
                        rec.Right = rec.Left + PIC_X
                        ' render
                        DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Next
            End With
        Next
    Next
    
    ' render the resources
    For Y = 0 To Map.MaxY
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
                            Call BltMapResource(i, True)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' render the tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For i = MapLayer.Fringe To MapLayer.Fringe2
                    ' skip tile?
                    If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                        ' sort out rec
                        rec.top = .Layer(i).Y * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        rec.Left = .Layer(i).X * PIC_X
                        rec.Right = rec.Left + PIC_X
                        ' render
                        DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                Next
            End With
        Next
    Next
    
    ' dump and save
    frmMain.picSSMap.width = DDSD_Map.lWidth
    frmMain.picSSMap.height = DDSD_Map.lHeight
    rec.top = 0
    rec.Left = 0
    rec.Bottom = DDSD_Map.lHeight
    rec.Right = DDSD_Map.lWidth
    Engine_BltToDC DDS_Map, rec, rec, frmMain.picSSMap
    SavePicture frmMain.picSSMap.Image, App.Path & "\map" & GetPlayerMap(MyIndex) & ".jpg"
    
    ' let them know we did it
    AddText "Screenshot of map #" & GetPlayerMap(MyIndex) & " saved.", BrightGreen
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotMap", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapResource(ByVal Resource_num As Long, Optional ByVal screenShot As Boolean = False)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As DxVBLib.RECT
Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).Data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' Load early
    If DDS_Resource(Resource_sprite) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource_sprite, DDSD_Resource(Resource_sprite), DDS_Resource(Resource_sprite))
    End If

    ' src rect
    With rec
        .top = 0
        .Bottom = DDSD_Resource(Resource_sprite).lHeight
        .Left = 0
        .Right = DDSD_Resource(Resource_sprite).lWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (DDSD_Resource(Resource_sprite).lWidth / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - DDSD_Resource(Resource_sprite).lHeight + 32
    
    ' render it
    If Not screenShot Then
        Call BltResource(Resource_sprite, X, Y, rec)
    Else
        Call ScreenshotResource(Resource_sprite, X, Y, rec)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltResource(ByVal Resource As Long, ByVal dx As Long, dy As Long, rec As DxVBLib.RECT)
Dim X As Long
Dim Y As Long
Dim width As Long
Dim height As Long
Dim destRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ResourceTimer(Resource) = GetTickCount + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If

    X = ConvertMapX(dx)
    Y = ConvertMapY(dy)
    
    width = (rec.Right - rec.Left)
    height = (rec.Bottom - rec.top)

    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If

    ' End clipping
    Call Engine_BltFast(X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScreenshotResource(ByVal Resource As Long, ByVal X As Long, Y As Long, rec As DxVBLib.RECT)
Dim width As Long
Dim height As Long
Dim destRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ResourceTimer(Resource) = GetTickCount + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If
    
    width = (rec.Right - rec.Left)
    height = (rec.Bottom - rec.top)

    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_Map.lHeight Then
        rec.Bottom = rec.Bottom - (Y + height - DDSD_Map.lHeight)
    End If

    If X + width > DDSD_Map.lWidth Then
        rec.Right = rec.Right - (X + width - DDSD_Map.lWidth)
    End If

    ' End clipping
    'Call Engine_BltFast(x, y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    DDS_Map.BltFast X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRECT As RECT
Dim barWidth As Long
Dim i As Long, npcNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' dynamic bar calculations
    sWidth = DDSD_Bars.lWidth
    sHeight = DDSD_Bars.lHeight / 4
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(npcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).XOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).YOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (Npc(npcNum).HP / sWidth)) * sWidth
                
                ' draw bar background
                With sRECT
                    .top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                
                ' draw the bar proper
                With sRECT
                    .top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 35 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' draw bar background
            With sRECT
                .top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            
            ' draw the bar proper
            With sRECT
                .top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).YOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' draw bar background
        With sRECT
            .top = sHeight * 1 ' HP bar background
            .Left = 0
            .Right = .Left + sWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
       
        ' draw the bar proper
        With sRECT
            .top = 0 ' HP bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).XOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).YOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    With sRECT
                        .top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    
                    ' draw the bar proper
                    With sRECT
                        .top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            End If
        Next
    End If
                    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBars", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltHotbar()
Dim sRECT As RECT, dRECT As RECT, i As Long, num As String, n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.picHotbar.Cls
    For i = 1 To MAX_HOTBAR
        With dRECT
            .top = HotbarTop
            .Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .Bottom = .top + 32
            .Right = .Left + 32
        End With
        
        With sRECT
            .top = 0
            .Left = 32
            .Bottom = 32
            .Right = 64
        End With
        
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).Name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        If DDS_Item(Item(Hotbar(i).Slot).Pic) Is Nothing Then
                            Call InitDDSurf("Items\" & Item(Hotbar(i).Slot).Pic, DDSD_Item(Item(Hotbar(i).Slot).Pic), DDS_Item(Item(Hotbar(i).Slot).Pic))
                        End If
                        Engine_BltToDC DDS_Item(Item(Hotbar(i).Slot).Pic), sRECT, dRECT, frmMain.picHotbar, False
                    End If
                End If
            Case 2 ' spell
                With sRECT
                    .top = 0
                    .Left = 0
                    .Bottom = 32
                    .Right = 32
                End With
                If Len(Spell(Hotbar(i).Slot).Name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 Then
                        If DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon) Is Nothing Then
                            Call InitDDSurf("Spellicons\" & Spell(Hotbar(i).Slot).Icon, DDSD_SpellIcon(Spell(Hotbar(i).Slot).Icon), DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon))
                        End If
                        ' check for cooldown
                        For n = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(n) = Hotbar(i).Slot Then
                                ' has spell
                                If Not SpellCD(i) = 0 Then
                                    sRECT.Left = 32
                                    sRECT.Right = 64
                                End If
                            End If
                        Next
                        Engine_BltToDC DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon), sRECT, dRECT, frmMain.picHotbar, False
                    End If
                End If
        End Select
        
        ' render the letters
        num = "F" & str(i)
        DrawText frmMain.picHotbar.hDC, dRECT.Left + 2, dRECT.top + 16, num, QBColor(White)
    Next
    frmMain.picHotbar.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHotbar", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPlayer(ByVal Index As Long)
Dim anim As Byte, i As Long, X As Long, Y As Long
Dim Sprite As Long, spritetop As Long
Dim rec As DxVBLib.RECT
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    ' Reset frame
    If Player(Index).Step = 3 Then
        anim = 0
    ElseIf Player(Index).Step = 1 Then
        anim = 2
    End If
    
    ' Check for attacking animation
    If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If Player(Index).Attacking = 1 Then
            anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset > 8) Then anim = Player(Index).Step
            Case DIR_DOWN
                If (Player(Index).YOffset < -8) Then anim = Player(Index).Step
            Case DIR_LEFT
                If (Player(Index).XOffset > 8) Then anim = Player(Index).Step
            Case DIR_RIGHT
                If (Player(Index).XOffset < -8) Then anim = Player(Index).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .top = spritetop * (DDSD_Character(Sprite).lHeight / 4)
        .Bottom = .top + (DDSD_Character(Sprite).lHeight / 4)
        .Left = anim * (DDSD_Character(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 4)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((DDSD_Character(Sprite).lWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (DDSD_Character(Sprite).lHeight) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((DDSD_Character(Sprite).lHeight / 4) - 32)
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
    End If

    ' render the actual sprite
    Call BltSprite(Sprite, X, Y, rec)
    
    ' check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, anim, spritetop)
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayer", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltNpc(ByVal MapNpcNum As Long)
Dim anim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
Dim rec As DxVBLib.RECT
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    
    Sprite = Npc(MapNpc(MapNpcNum).num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    attackspeed = 1000

    ' Reset frame
    anim = 0
    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset > 8) Then anim = MapNpc(MapNpcNum).Step
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < -8) Then anim = MapNpc(MapNpcNum).Step
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset > 8) Then anim = MapNpc(MapNpcNum).Step
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < -8) Then anim = MapNpc(MapNpcNum).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .top = (DDSD_Character(Sprite).lHeight / 4) * spritetop
        .Bottom = .top + DDSD_Character(Sprite).lHeight / 4
        .Left = anim * (DDSD_Character(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 4)
    End With

    ' Calculate the X
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset - ((DDSD_Character(Sprite).lWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (DDSD_Character(Sprite).lHeight / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset - ((DDSD_Character(Sprite).lHeight / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
    End If

    Call BltSprite(Sprite, X, Y, rec)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltNpc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPaperdoll(ByVal x2 As Long, ByVal y2 As Long, ByVal Sprite As Long, ByVal anim As Long, ByVal spritetop As Long)
Dim rec As DxVBLib.RECT
Dim X As Long, Y As Long
Dim width As Long, height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("Paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If
    
    With rec
        .top = spritetop * (DDSD_Paperdoll(Sprite).lHeight / 4)
        .Bottom = .top + (DDSD_Paperdoll(Sprite).lHeight / 4)
        .Left = anim * (DDSD_Paperdoll(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Paperdoll(Sprite).lWidth / 4)
    End With
    
    ' clipping
    X = ConvertMapX(x2)
    Y = ConvertMapY(y2)
    width = (rec.Right - rec.Left)
    height = (rec.Bottom - rec.top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Paperdoll(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltSprite(ByVal Sprite As Long, ByVal x2 As Long, y2 As Long, rec As DxVBLib.RECT)
Dim X As Long
Dim Y As Long
Dim width As Long
Dim height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(x2)
    Y = ConvertMapY(y2)
    width = (rec.Right - rec.Left)
    height = (rec.Bottom - rec.top)

    ' clipping
    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    ' /clipping
    
    Call Engine_BltFast(X, Y, DDS_Character(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltAnimatedInvItems()
Dim i As Long
Dim itemnum As Long, itempic As Long
Dim X As Long, Y As Long
Dim MaxFrames As Byte
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS

        If MapItem(i).num > 0 Then
            itempic = Item(MapItem(i).num).Pic

            If itempic < 1 Or itempic > NumItems Then Exit Sub
            MaxFrames = (DDSD_Item(itempic).lWidth / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < MaxFrames - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 1
            End If
        End If

    Next

    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= NumItems Then
                If DDSD_Item(itempic).lWidth > 64 Then
                    MaxFrames = (DDSD_Item(itempic).lWidth / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(i) < MaxFrames - 1 Then
                        InvItemFrame(i) = InvItemFrame(i) + 1
                    Else
                        InvItemFrame(i) = 1
                    End If

                    With rec
                        .top = 0
                        .Bottom = 32
                        .Left = (DDSD_Item(itempic).lWidth / 2) + (InvItemFrame(i) * 32) ' middle to get the start of inv gfx, then +32 for each frame
                        .Right = .Left + 32
                    End With

                    With rec_pos
                        .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    ' We'll now re-blt the item, and place the currency value over it again :P
                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.top + 22
                        X = rec_pos.Left - 4
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        DrawText frmMain.picInventory.hDC, X, Y, ConvertCurrency(Amount), QBColor(Yellow)

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If

    Next

    frmMain.picInventory.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimatedInvItems", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltFace()
Dim rec As RECT, rec_pos As RECT, faceNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub
    
    frmMain.picFace.Cls
    
    faceNum = GetPlayerSprite(MyIndex)
    
    If faceNum <= 0 Or faceNum > NumFaces Then Exit Sub

    With rec
        .top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    With rec_pos
        .top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    ' Load face if not loaded, and reset timer
    FaceTimer(faceNum) = GetTickCount + SurfaceTimerMax

    If DDS_Face(faceNum) Is Nothing Then
        Call InitDDSurf("Faces\" & faceNum, DDSD_Face(faceNum), DDS_Face(faceNum))
    End If

    Engine_BltToDC DDS_Face(faceNum), rec, rec_pos, frmMain.picFace, False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltFace", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltEquipment()
Dim i As Long, itemnum As Long, itempic As Long
Dim rec As RECT, rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumItems = 0 Then Exit Sub
    
    frmMain.picCharacter.Cls

    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(MyIndex, i)

        If itemnum > 0 Then
            itempic = Item(itemnum).Pic

            With rec
                .top = 0
                .Bottom = 32
                .Left = 32
                .Right = 64
            End With

            With rec_pos
                .top = EqTop
                .Bottom = .top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            ' Load item if not loaded, and reset timer
            ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

            If DDS_Item(itempic) Is Nothing Then
                Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
            End If

            Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picCharacter, False
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltEquipment", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltInventory()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT
Dim colour As Long
Dim tmpItem As Long, amountModifier As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' reset gold label
    frmMain.lblGold.Caption = "0g"
    
    frmMain.picInventory.Cls

    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            
            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                    If TradeYourOffer(X).num = i Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(X).Value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(X).Value
                            End If
                        End If
                    End If
                Next
            End If

            If itempic > 0 And itempic <= NumItems Then
                If DDSD_Item(itempic).lWidth <= 64 Then ' more than 1 frame is handled by anim sub

                    With rec
                        .top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With

                    With rec_pos
                        .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                    If DDS_Item(itempic) Is Nothing Then
                        Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                    End If

                    Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.top + 22
                        X = rec_pos.Left - 4
                        
                        Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            colour = QBColor(White)
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            colour = QBColor(Yellow)
                        ElseIf Amount > 10000000 Then
                            colour = QBColor(BrightGreen)
                        End If
                        
                        DrawText frmMain.picInventory.hDC, X, Y, Format$(ConvertCurrency(str(Amount)), "#,###,###,###"), colour

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If
NextLoop:
    Next
    
    frmMain.picInventory.Refresh
    'update animated items
    BltAnimatedInvItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventory", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltTrade()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picYourTrade.Cls
    frmMain.picTheirTrade.Cls
    
    For i = 1 To MAX_INV
        ' blt your own offer
        itemnum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= NumItems Then
                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With rec_pos
                    .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If

                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picYourTrade, False

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).Value > 1 Then
                    Y = rec_pos.top + 22
                    X = rec_pos.Left - 4
                    
                    Amount = TradeYourOffer(i).Value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picYourTrade.hDC, X, Y, ConvertCurrency(str(Amount)), colour
                End If
            End If
        End If
            
        ' blt their offer
        itemnum = TradeTheirOffer(i).num

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= NumItems Then
                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With rec_pos
                    .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If

                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picTheirTrade, False

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).Value > 1 Then
                    Y = rec_pos.top + 22
                    X = rec_pos.Left - 4
                    
                    Amount = TradeTheirOffer(i).Value
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picTheirTrade.hDC, X, Y, ConvertCurrency(str(Amount)), colour
                End If
            End If
        End If
    Next
    
    frmMain.picYourTrade.Refresh
    frmMain.picTheirTrade.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTrade", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltPlayerSpells()
Dim i As Long, X As Long, Y As Long, spellnum As Long, spellicon As Long
Dim Amount As String
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picSpells.Cls

    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i)

        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellicon = Spell(spellnum).Icon

            If spellicon > 0 And spellicon <= NumSpellIcons Then
            
                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 0
                    .Right = 32
                End With
                
                If Not SpellCD(i) = 0 Then
                    rec.Left = 32
                    rec.Right = 64
                End If

                With rec_pos
                    .top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                    .Bottom = .top + PIC_Y
                    .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load spellicon if not loaded, and reset timer
                SpellIconTimer(spellicon) = GetTickCount + SurfaceTimerMax

                If DDS_SpellIcon(spellicon) Is Nothing Then
                    Call InitDDSurf("SpellIcons\" & spellicon, DDSD_SpellIcon(spellicon), DDS_SpellIcon(spellicon))
                End If

                Engine_BltToDC DDS_SpellIcon(spellicon), rec, rec_pos, frmMain.picSpells, False
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayerSpells", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltShop()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As String
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    frmMain.picShopItems.Cls

    For i = 1 To MAX_TRADES
        itemnum = Shop(InShop).TradeItem(i).Item 'GetPlayerInvItemNum(MyIndex, i)
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            If itempic > 0 And itempic <= NumItems Then
            
                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With
                
                With rec_pos
                    .top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                    .Bottom = .top + PIC_Y
                    .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                    .Right = .Left + PIC_X
                End With
                
                ' Load item if not loaded, and reset timer
                ItemTimer(itempic) = GetTickCount + SurfaceTimerMax
                
                If DDS_Item(itempic) Is Nothing Then
                    Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
                End If
                
                Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picShopItems, False
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    Y = rec_pos.top + 22
                    X = rec_pos.Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = QBColor(White)
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picShopItems.hDC, X, Y, ConvertCurrency(Amount), colour
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltShop", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltInventoryItem(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT
Dim itemnum As Long, itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)

    If itemnum > 0 And itemnum <= MAX_ITEMS Then
        itempic = Item(itemnum).Pic
        
        If itempic = 0 Then Exit Sub

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = DDSD_Item(itempic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .top = 2
            .Bottom = .top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

        If DDS_Item(itempic) Is Nothing Then
            Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
        End If

        Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picTempInv, False

        With frmMain.picTempInv
            .top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltDraggedSpell(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT
Dim spellnum As Long, spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = PlayerSpells(DragSpell)

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon
        
        If spellpic = 0 Then Exit Sub

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .top = 2
            .Bottom = .top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = GetTickCount + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("Spellicons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        Engine_BltToDC DDS_SpellIcon(spellpic), rec, rec_pos, frmMain.picTempSpell, False

        With frmMain.picTempSpell
            .top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItemDesc(ByVal itemnum As Long)
Dim rec As RECT, rec_pos As RECT
Dim itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.picItemDescPic.Cls
    
    If itemnum > 0 And itemnum <= MAX_ITEMS Then
        itempic = Item(itemnum).Pic

        If itempic = 0 Then Exit Sub
        
        ' Load item if not loaded, and reset timer
        ItemTimer(itempic) = GetTickCount + SurfaceTimerMax

        If DDS_Item(itempic) Is Nothing Then
            Call InitDDSurf("Items\" & itempic, DDSD_Item(itempic), DDS_Item(itempic))
        End If

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = DDSD_Item(itempic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        Engine_BltToDC DDS_Item(itempic), rec, rec_pos, frmMain.picItemDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItemDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltSpellDesc(ByVal spellnum As Long)
Dim rec As RECT, rec_pos As RECT
Dim spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.picSpellDescPic.Cls

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon

        If spellpic <= 0 Or spellpic > NumSpellIcons Then Exit Sub
        
        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = GetTickCount + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("SpellIcons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        Engine_BltToDC DDS_SpellIcon(spellpic), rec, rec_pos, frmMain.picSpellDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSpellDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_BltTileset()
Dim height As Long
Dim width As Long
Dim Tileset As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    ' make sure it's loaded
    If DDS_Tileset(Tileset) Is Nothing Then
        Call InitDDSurf("tilesets\" & Tileset, DDSD_Tileset(Tileset), DDS_Tileset(Tileset))
    End If
    
    height = DDSD_Tileset(Tileset).lHeight
    width = DDSD_Tileset(Tileset).lWidth
    
    dRECT.top = 0
    dRECT.Bottom = height
    dRECT.Left = 0
    dRECT.Right = width
    
    frmEditor_Map.picBackSelect.height = height
    frmEditor_Map.picBackSelect.width = width
    
    Call Engine_BltToDC(DDS_Tileset(Tileset), sRECT, dRECT, frmEditor_Map.picBackSelect)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltTileset", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTileOutline()
Dim rec As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optBlock.Value Then Exit Sub

    With rec
        .top = 0
        .Bottom = .top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    Call Engine_BltFast(ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), DDS_Misc, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTileOutline", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NewCharacterBltSprite()
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim width As Long, height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.Value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("Characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If
    
    width = DDSD_Character(Sprite).lWidth / 4
    height = DDSD_Character(Sprite).lHeight / 4
    
    frmMenu.picSprite.width = width
    frmMenu.picSprite.height = height
    
    sRECT.top = 0
    sRECT.Bottom = sRECT.top + height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + width
    
    dRECT.top = 0
    dRECT.Bottom = height
    dRECT.Left = 0
    dRECT.Right = width
    
    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmMenu.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterBltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltMapItem()
Dim itemnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = Item(frmEditor_Map.scrlMapItem.Value).Pic

    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    ItemTimer(itemnum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itemnum) Is Nothing Then
        Call InitDDSurf("Items\" & itemnum, DDSD_Item(itemnum), DDS_Item(itemnum))
    End If

    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(itemnum), sRECT, dRECT, frmEditor_Map.picMapItem)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltMapItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltKey()
Dim itemnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = Item(frmEditor_Map.scrlMapKey.Value).Pic

    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If

    ItemTimer(itemnum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itemnum) Is Nothing Then
        Call InitDDSurf("Items\" & itemnum, DDSD_Item(itemnum), DDS_Item(itemnum))
    End If

    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(itemnum), sRECT, dRECT, frmEditor_Map.picMapKey)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltKey", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltItem()
Dim itemnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = frmEditor_Item.scrlPic.Value

    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ItemTimer(itemnum) = GetTickCount + SurfaceTimerMax

    If DDS_Item(itemnum) Is Nothing Then
        Call InitDDSurf("Items\" & itemnum, DDSD_Item(itemnum), DDS_Item(itemnum))
    End If

    ' rect for source
    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRECT = sRECT
    Call Engine_BltToDC(DDS_Item(itemnum), sRECT, dRECT, frmEditor_Item.picItem)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltPaperdoll()
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Item.picPaperdoll.Cls
    
    Sprite = frmEditor_Item.scrlPaperdoll.Value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    PaperdollTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If

    ' rect for source
    sRECT.top = 0
    sRECT.Bottom = DDSD_Paperdoll(Sprite).lHeight
    sRECT.Left = 0
    sRECT.Right = DDSD_Paperdoll(Sprite).lWidth
    ' same for destination as source
    dRECT = sRECT
    
    Call Engine_BltToDC(DDS_Paperdoll(Sprite), sRECT, dRECT, frmEditor_Item.picPaperdoll)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_BltIcon()
Dim iconnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    iconnum = frmEditor_Spell.scrlIcon.Value
    
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    SpellIconTimer(iconnum) = GetTickCount + SurfaceTimerMax
    
    If DDS_SpellIcon(iconnum) Is Nothing Then
        Call InitDDSurf("SpellIcons\" & iconnum, DDSD_SpellIcon(iconnum), DDS_SpellIcon(iconnum))
    End If
    
    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    
    Call Engine_BltToDC(DDS_SpellIcon(iconnum), sRECT, dRECT, frmEditor_Spell.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_BltIcon", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_BltAnim()
Dim Animationnum As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
Dim i As Long
Dim width As Long, height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim ShouldRender As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value
        
        If Animationnum < 1 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                frmEditor_Animation.picSprite(i).Cls
            
                AnimationTimer(Animationnum) = GetTickCount + SurfaceTimerMax
                
                If DDS_Animation(Animationnum) Is Nothing Then
                    Call InitDDSurf("animations\" & Animationnum, DDSD_Animation(Animationnum), DDS_Animation(Animationnum))
                End If
                
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
                    ' total width divided by frame count
                    width = DDSD_Animation(Animationnum).lWidth / frmEditor_Animation.scrlFrameCount(i).Value
                    height = DDSD_Animation(Animationnum).lHeight
                    
                    sRECT.top = 0
                    sRECT.Bottom = height
                    sRECT.Left = (AnimEditorFrame(i) - 1) * width
                    sRECT.Right = sRECT.Left + width
                    
                    dRECT.top = 0
                    dRECT.Bottom = height
                    dRECT.Left = 0
                    dRECT.Right = width
                    
                    Call Engine_BltToDC(DDS_Animation(Animationnum), sRECT, dRECT, frmEditor_Animation.picSprite(i))
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_BltAnim", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_BltSprite()
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_NPC.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    CharacterTimer(Sprite) = GetTickCount + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    sRECT.top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = PIC_X * 3 ' facing down
    sRECT.Right = sRECT.Left + SIZE_X
    dRECT.top = 0
    dRECT.Bottom = SIZE_Y
    dRECT.Left = 0
    dRECT.Right = SIZE_X
    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmEditor_NPC.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_BltSprite()
Dim Sprite As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        ResourceTimer(Sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picNormalPic)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        ResourceTimer(Sprite) = GetTickCount + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picExhaustedPic)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check if automation is screwed
    If Not CheckSurfaces Then
        ' exit out and let them know we need to re-init
        ReInitSurfaces = True
        Exit Sub
    Else
        ' if we need to fix the surfaces then do so
        If ReInitSurfaces Then
            ReInitSurfaces = False
            ReInitDD
        End If
    End If
    
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub
    
    ' update the viewpoint
    UpdateCamera
    
    ' update animation editor
    If Editor = EDITOR_ANIMATION Then
        EditorAnim_BltAnim
    End If
    
    ' fill it with black
    DDS_BackBuffer.BltColorFill rec_pos, 0
    
    ' blit lower tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapTile(X, Y)
                End If
            Next
        Next
    End If

    ' render the decals
    For i = 1 To MAX_BYTE
        Call BltBlood(i)
    Next

    ' Blit out the items
    If NumItems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call BltItem(i)
            End If
        Next
    End If
    
    If Map.CurrentEvents > 0 Then
        For i = 1 To Map.CurrentEvents
            If Map.MapEvents(i).Position = 0 Then
                BltEvent i
            End If
        Next
    End If
    
    ' draw animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                BltAnimation i, 0
            End If
        Next
    End If

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For Y = 0 To Map.MaxY
        If NumCharacters > 0 Then
        
            If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    If Map.MapEvents(i).Position = 1 Then
                        If Y = Map.MapEvents(i).Y Then
                            BltEvent i
                        End If
                    End If
                Next
            End If
            
            ' Players
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).Y = Y Then
                        Call BltPlayer(i)
                    End If
                End If
            Next
            
            
        
            ' Npcs
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Y = Y Then
                    Call BltNpc(i)
                End If
            Next
        End If
        
        ' Resources
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
                            Call BltMapResource(i)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                BltAnimation i, 1
            End If
        Next
    End If

    ' blit out upper tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapFringeTile(X, Y)
                End If
            Next
        Next
    End If
    
    If Map.CurrentEvents > 0 Then
        For i = 1 To Map.CurrentEvents
            If Map.MapEvents(i).Position = 2 Then
                BltEvent i
            End If
        Next
    End If
    
    ' blit out a square at mouse cursor
    If InMapEditor Then
        If frmEditor_Map.optBlock.Value = True Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.top To TileView.Bottom
                    If IsValidMapPoint(X, Y) Then
                        Call BltDirection(X, Y)
                    End If
                Next
            Next
        End If
        Call BltTileOutline
    End If
    
    ' Render the bars
    BltBars
    
    ' Blt the target icon
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            BltTarget (Player(myTarget).X * 32) + Player(myTarget).XOffset, (Player(myTarget).Y * 32) + Player(myTarget).YOffset
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            BltTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).XOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).YOffset
        End If
    End If
    
    ' blt the hover icon
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Map = Player(MyIndex).Map Then
                If CurX = Player(i).X And CurY = Player(i).Y Then
                    If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                        ' dont render lol
                    Else
                        BltHover TARGET_TYPE_PLAYER, i, (Player(i).X * 32) + Player(i).XOffset, (Player(i).Y * 32) + Player(i).YOffset
                    End If
                End If
            End If
        End If
    Next
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then
                If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                    ' dont render lol
                Else
                    BltHover TARGET_TYPE_NPC, i, (MapNpc(i).X * 32) + MapNpc(i).XOffset, (MapNpc(i).Y * 32) + MapNpc(i).YOffset
                End If
            End If
        End If
    Next
    
    If frmEditor_Events.Visible Then
        EditorEvent_BltGraphic
    End If
    
    If InMapEditor Then
        If frmEditor_Map.optEvent.Value = True Then
            BltEvents
        End If
    End If

    ' Lock the backbuffer so we can draw text and names
    TexthDC = DDS_BackBuffer.GetDC

    ' draw FPS
    If BFPS Then
        Call DrawText(TexthDC, Camera.Right - (Len("FPS: " & GameFPS) * 8), Camera.top + 1, Trim$("FPS: " & GameFPS), QBColor(Yellow))
    End If

    ' draw cursor, player X and Y locations
    If BLoc Then
        Call DrawText(TexthDC, Camera.Left, Camera.top + 1, Trim$("cur x: " & CurX & " y: " & CurY), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.top + 15, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.top + 27, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), QBColor(Yellow))
    End If

    ' draw player names
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(i)
        End If
    Next
    
    For i = 1 To Map.CurrentEvents
        If Map.MapEvents(i).Visible = 1 Then
            If Map.MapEvents(i).ShowName = 1 Then
                DrawEventName (i)
            End If
        End If
    Next
    
    ' draw npc names
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            Call DrawNpcName(i)
        End If
    Next
    
    For i = 1 To Action_HighIndex
        Call BltActionMsg(i)
    Next i

    ' Blit out map attributes
    If InMapEditor Then
        Call BltMapAttributes
    End If

    ' Draw map name
    Call DrawText(TexthDC, DrawMapNameX, DrawMapNameY, Map.Name, DrawMapNameColor)

    ' Release DC
    DDS_BackBuffer.ReleaseDC TexthDC
    
    ' Get rec
    With rec
        .top = Camera.top
        .Bottom = .top + ScreenY
        .Left = Camera.Left
        .Right = .Left + ScreenX
    End With
    
    ' rec_pos
    With rec_pos
        .Bottom = ((MAX_MAPY + 1) * PIC_Y)
        .Right = ((MAX_MAPX + 1) * PIC_X)
    End With
    
    ' Flip and render
    DX7.GetWindowRect frmMain.picScreen.hWnd, rec_pos
    DDS_Primary.Blt rec_pos, DDS_BackBuffer, rec, DDBLT_WAIT
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "Render_Graphics", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    offsetX = Player(MyIndex).XOffset + PIC_X
    offsetY = Player(MyIndex).YOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).XOffset > 0 Then
                offsetX = Player(MyIndex).XOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).YOffset > 0 Then
                offsetY = Player(MyIndex).YOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        offsetX = 32
        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).XOffset < 0 Then
                offsetX = Player(MyIndex).XOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        offsetY = 32
        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).YOffset < 0 Then
                offsetY = Player(MyIndex).YOffset + PIC_Y
            End If
        End If
        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .top = offsetY
        .Bottom = .top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = X - (TileView.Left * PIC_X)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = Y - (TileView.top * PIC_Y)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(X, Y).Layer(i).Tileset > 0 And Map.Tile(X, Y).Layer(i).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(X, Y).Layer(i).Tileset) = True
                End If
            Next
        Next
    Next
    
    For i = 1 To NumTileSets
        If tilesetInUse(i) Then
            ' load tileset
            If DDS_Tileset(i) Is Nothing Then
                Call InitDDSurf("tilesets\" & i, DDSD_Tileset(i), DDS_Tileset(i))
            End If
        Else
            ' unload tileset
            Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            Set DDS_Tileset(i) = Nothing
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltBank()
Dim i As Long, X As Long, Y As Long, itemnum As Long
Dim Amount As String
Dim sRECT As RECT, dRECT As RECT
Dim Sprite As Long, colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.picBank.Visible = True Then
        frmMain.picBank.Cls
                
        For i = 1 To MAX_BANK
            itemnum = GetBankItemNum(i)
            If itemnum > 0 And itemnum <= MAX_ITEMS Then
            
                Sprite = Item(itemnum).Pic
                
                If Sprite <= 0 Or Sprite > NumItems Then Exit Sub
                
                If DDS_Item(Sprite) Is Nothing Then
                    Call InitDDSurf("Items\" & Sprite, DDSD_Item(Sprite), DDS_Item(Sprite))
                End If
            
                With sRECT
                    .top = 0
                    .Bottom = .top + PIC_Y
                    .Left = DDSD_Item(Sprite).lWidth / 2
                    .Right = .Left + PIC_X
                End With
                
                With dRECT
                    .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    .Bottom = .top + PIC_Y
                    .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                    .Right = .Left + PIC_X
                End With
                
                Engine_BltToDC DDS_Item(Sprite), sRECT, dRECT, frmMain.picBank, False

                ' If item is a stack - draw the amount you have
                If GetBankItemValue(i) > 1 Then
                    Y = dRECT.top + 22
                    X = dRECT.Left - 4
                
                    Amount = CStr(GetBankItemValue(i))
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = QBColor(White)
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    DrawText frmMain.picBank.hDC, X, Y, ConvertCurrency(Amount), colour
                End If
            End If
        Next
    
        frmMain.picBank.Refresh
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBank", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltBankItem(ByVal X As Long, ByVal Y As Long)
Dim sRECT As RECT, dRECT As RECT
Dim itemnum As Long
Dim Sprite As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = GetBankItemNum(DragBankSlotNum)
    Sprite = Item(GetBankItemNum(DragBankSlotNum)).Pic
    
    If DDS_Item(Sprite) Is Nothing Then
        Call InitDDSurf("Items\" & Sprite, DDSD_Item(Sprite), DDS_Item(Sprite))
    End If
    
    If itemnum > 0 Then
        If itemnum <= MAX_ITEMS Then
            With sRECT
                .top = 0
                .Bottom = .top + PIC_Y
                .Left = DDSD_Item(Sprite).lWidth / 2
                .Right = .Left + PIC_X
            End With
        End If
    End If
    
    With dRECT
        .top = 2
        .Bottom = .top + PIC_Y
        .Left = 2
        .Right = .Left + PIC_X
    End With

    Engine_BltToDC DDS_Item(Sprite), sRECT, dRECT, frmMain.picTempBank
    
    With frmMain.picTempBank
        .top = Y
        .Left = X
        .Visible = True
        .ZOrder (0)
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBankItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltEvents()
Dim sRECT As DxVBLib.RECT
Dim width As Long, height As Long, i As Long, X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Map.EventCount <= 0 Then Exit Sub
    
    For i = 1 To Map.EventCount
        If Map.Events(i).pageCount <= 0 Then
                sRECT.top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            GoTo nextevent
        End If
        
        width = 32
        height = 32
    
        X = Map.Events(i).X * 32
        Y = Map.Events(i).Y * 32
        X = ConvertMapX(X)
        Y = ConvertMapY(Y)
    
        ' clipping
        If Y < 0 Then
            With sRECT
                .top = .top - Y
            End With
            Y = 0
        End If
    
        If X < 0 Then
            With sRECT
                .Left = .Left - X
            End With
            X = 0
        End If
    
        If Y + height > DDSD_BackBuffer.lHeight Then
            sRECT.Bottom = sRECT.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
        End If
    
        If X + width > DDSD_BackBuffer.lWidth Then
            sRECT.Right = sRECT.Right - (X + width - DDSD_BackBuffer.lWidth)
        End If
        
        If i > Map.EventCount Then Exit Sub
        If 1 > Map.Events(i).pageCount Then Exit Sub
    ' /clipping
        Select Case Map.Events(i).Pages(1).GraphicType
            Case 0
                sRECT.top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Case 1
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic <= NumCharacters Then
                    CharacterTimer(Map.Events(i).Pages(1).Graphic) = GetTickCount + SurfaceTimerMax
                    If DDS_Character(Map.Events(i).Pages(1).Graphic) Is Nothing Then
                        Call InitDDSurf("Characters\" & Map.Events(i).Pages(1).Graphic, DDSD_Character(Map.Events(i).Pages(1).Graphic), DDS_Character(Map.Events(i).Pages(1).Graphic))
                    End If
                    
                    sRECT.top = (Map.Events(i).Pages(1).GraphicY * (DDSD_Character(Map.Events(i).Pages(1).Graphic).lHeight / 4))
                    sRECT.Left = (Map.Events(i).Pages(1).GraphicX * (DDSD_Character(Map.Events(i).Pages(1).Graphic).lWidth / 4))
                    sRECT.Bottom = sRECT.top + 32
                    sRECT.Right = sRECT.Left + 32
                    
                    Call Engine_BltFast(X, Y, DDS_Character(Map.Events(i).Pages(1).Graphic), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    
                    sRECT.top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_SRCCOLORKEY)
                Else
                    sRECT.top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_SRCCOLORKEY)
                End If
            Case 2
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic < NumTileSets Then
                    sRECT.top = Map.Events(i).Pages(1).GraphicY * 32
                    sRECT.Left = Map.Events(i).Pages(1).GraphicX * 32
                    sRECT.Bottom = sRECT.top + 32
                    sRECT.Right = sRECT.Left + 32
                    Call Engine_BltFast(X, Y, DDS_Tileset(Map.Events(i).Pages(1).Graphic), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    
                    sRECT.top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    sRECT.top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    Call Engine_BltFast(X, Y, DDS_Event, sRECT, DDBLTFAST_SRCCOLORKEY)
                End If
        End Select
nextevent:
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHover", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorEvent_BltGraphic()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Events.picGraphicSel.Visible Then
        Select Case frmEditor_Events.cmbGraphic.ListIndex
            Case 0
                'None
                frmEditor_Events.picGraphicSel.Cls
                Exit Sub
            Case 1
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumCharacters Then
                    CharacterTimer(frmEditor_Events.scrlGraphic.Value) = GetTickCount + SurfaceTimerMax
                    If DDS_Character(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                        Call InitDDSurf("Characters\" & frmEditor_Events.scrlGraphic.Value, DDSD_Character(frmEditor_Events.scrlGraphic.Value), DDS_Character(frmEditor_Events.scrlGraphic.Value))
                    End If
                    
                    If DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth > 800 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Right = sRECT.Left + (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth - sRECT.Left)
                    Else
                        sRECT.Left = 0
                        sRECT.Right = DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth
                    End If
                    
                    If DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight > 512 Then
                        sRECT.top = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Bottom = sRECT.top + (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight - sRECT.top)
                    Else
                        sRECT.top = 0
                        sRECT.Bottom = DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight
                    End If
                    
                    With dRECT
                        .top = 0
                        .Bottom = sRECT.Bottom - sRECT.top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    Call Engine_BltToDC(DDS_Character(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphicSel)
                    
                    If GraphicSelX <= 3 And GraphicSelY <= 3 Then
                        frmEditor_Events.shpLoc.Left = GraphicSelX * (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth / 4)
                        frmEditor_Events.shpLoc.width = (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth / 4)
                        frmEditor_Events.shpLoc.top = GraphicSelY * (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight / 4)
                        frmEditor_Events.shpLoc.height = (DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight / 4)
                    Else
                        GraphicSelX = 0
                        GraphicSelY = 0
                    End If
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
            Case 2
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumTileSets Then
                    If DDS_Tileset(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                        Call InitDDSurf("Tilesets\" & frmEditor_Events.scrlGraphic.Value, DDSD_Tileset(frmEditor_Events.scrlGraphic.Value), DDS_Tileset(frmEditor_Events.scrlGraphic.Value))
                    End If
                    
                    If DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lWidth > 800 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Right = sRECT.Left + 800
                    Else
                        sRECT.Left = 0
                        sRECT.Right = DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lWidth
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value = 0
                    End If
                    
                    If DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lHeight > 512 Then
                        sRECT.top = frmEditor_Events.vScrlGraphicSel.Value
                        sRECT.Bottom = sRECT.top + 512
                    Else
                        sRECT.top = 0
                        sRECT.Bottom = DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lHeight
                        frmEditor_Events.vScrlGraphicSel.Value = 0
                    End If
                    
                    With dRECT
                        .top = 0
                        .Bottom = sRECT.Bottom - sRECT.top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    Call Engine_BltToDC(DDS_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphicSel)
         
                    'Now we draw the selection square.. tad bit harder....
                    'Stretched or not....
                    If GraphicSelX2 > 0 Or GraphicSelY2 > 0 Then
                        frmEditor_Events.shpLoc.top = (GraphicSelY * 32) - frmEditor_Events.vScrlGraphicSel.Value
                        frmEditor_Events.shpLoc.Left = (GraphicSelX * 32) - frmEditor_Events.hScrlGraphicSel.Value
                        frmEditor_Events.shpLoc.width = (GraphicSelX2 - GraphicSelX) * 32
                        frmEditor_Events.shpLoc.height = (GraphicSelY2 - GraphicSelY) * 32
                    Else
                        frmEditor_Events.shpLoc.top = (GraphicSelY * 32) - frmEditor_Events.vScrlGraphicSel.Value
                        frmEditor_Events.shpLoc.Left = (GraphicSelX * 32) - frmEditor_Events.hScrlGraphicSel.Value
                        frmEditor_Events.shpLoc.width = 32
                        frmEditor_Events.shpLoc.height = 32
                    End If
                    
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
        End Select
    Else
        Select Case tmpEvent.Pages(curPageNum).GraphicType
            Case 0
                frmEditor_Events.picGraphic.Cls
            Case 1
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumCharacters Then
                    sRECT.top = tmpEvent.Pages(curPageNum).GraphicY * (DDSD_Character(tmpEvent.Pages(curPageNum).Graphic).lHeight / 4)
                    sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * (DDSD_Character(tmpEvent.Pages(curPageNum).Graphic).lWidth / 4)
                    sRECT.Bottom = sRECT.top + (DDSD_Character(tmpEvent.Pages(curPageNum).Graphic).lHeight / 4)
                    sRECT.Right = sRECT.Left + (DDSD_Character(tmpEvent.Pages(curPageNum).Graphic).lWidth / 4)
                    With dRECT
                        dRECT.top = (193 / 2) - ((sRECT.Bottom - sRECT.top) / 2)
                        dRECT.Bottom = dRECT.top + (sRECT.Bottom - sRECT.top)
                        dRECT.Left = (121 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                        dRECT.Right = dRECT.Left + (sRECT.Right - sRECT.Left)
                    End With
                    If DDS_Character(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                        Call InitDDSurf("Characters\" & frmEditor_Events.scrlGraphic.Value, DDSD_Character(frmEditor_Events.scrlGraphic.Value), DDS_Character(frmEditor_Events.scrlGraphic.Value))
                    End If
                    Call Engine_BltToDC(DDS_Character(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphic)
                End If
            Case 2
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumTileSets Then
                    If tmpEvent.Pages(curPageNum).GraphicX2 = 0 Or tmpEvent.Pages(curPageNum).GraphicY2 = 0 Then
                        sRECT.top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.top + 32
                        sRECT.Right = sRECT.Left + 32
                        With dRECT
                            dRECT.top = (193 / 2) - ((sRECT.Bottom - sRECT.top) / 2)
                            dRECT.Bottom = dRECT.top + (sRECT.Bottom - sRECT.top)
                            dRECT.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            dRECT.Right = dRECT.Left + (sRECT.Right - sRECT.Left)
                        End With
                        If DDS_Tileset(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                            Call InitDDSurf("Tilesets\" & frmEditor_Events.scrlGraphic.Value, DDSD_Tileset(frmEditor_Events.scrlGraphic.Value), DDS_Tileset(frmEditor_Events.scrlGraphic.Value))
                        End If
                        Call Engine_BltToDC(DDS_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphic)

                    Else
                        sRECT.top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.top + ((tmpEvent.Pages(curPageNum).GraphicY2 - tmpEvent.Pages(curPageNum).GraphicY) * 32)
                        sRECT.Right = sRECT.Left + ((tmpEvent.Pages(curPageNum).GraphicX2 - tmpEvent.Pages(curPageNum).GraphicX) * 32)
                        With dRECT
                            dRECT.top = (193 / 2) - ((sRECT.Bottom - sRECT.top) / 2)
                            dRECT.Bottom = dRECT.top + (sRECT.Bottom - sRECT.top)
                            dRECT.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            dRECT.Right = dRECT.Left + (sRECT.Right - sRECT.Left)
                        End With
                        If DDS_Tileset(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
                            Call InitDDSurf("Tilesets\" & frmEditor_Events.scrlGraphic.Value, DDSD_Tileset(frmEditor_Events.scrlGraphic.Value), DDS_Tileset(frmEditor_Events.scrlGraphic.Value))
                        End If
                        Call Engine_BltToDC(DDS_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, dRECT, frmEditor_Events.picGraphic)
                    End If
                End If
        End Select
    End If
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltKey", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltEvent(id As Long)
    Dim X As Long, Y As Long, width As Long, height As Long, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, anim As Long, spritetop As Long
    If Map.MapEvents(id).Visible = 0 Then Exit Sub
    If InMapEditor Then Exit Sub
    Select Case Map.MapEvents(id).GraphicType
        Case 0
            Exit Sub
            
        Case 1
            If Map.MapEvents(id).GraphicNum <= 0 Or Map.MapEvents(id).GraphicNum > NumCharacters Then Exit Sub
            CharacterTimer(Map.MapEvents(id).GraphicNum) = GetTickCount + SurfaceTimerMax
            If DDS_Character(Map.MapEvents(id).GraphicNum) Is Nothing Then
                Call InitDDSurf("characters\" & Map.MapEvents(id).GraphicNum, DDSD_Character(Map.MapEvents(id).GraphicNum), DDS_Character(Map.MapEvents(id).GraphicNum))
            End If
            width = DDSD_Character(Map.MapEvents(id).GraphicNum).lWidth / 4
            height = DDSD_Character(Map.MapEvents(id).GraphicNum).lHeight / 4
            ' Reset frame
            If Map.MapEvents(id).Step = 3 Then
                anim = 0
            ElseIf Map.MapEvents(id).Step = 1 Then
                anim = 2
            End If
            
            Select Case Map.MapEvents(id).Dir
                Case DIR_UP
                    If (Map.MapEvents(id).YOffset > 8) Then anim = Map.MapEvents(id).Step
                Case DIR_DOWN
                    If (Map.MapEvents(id).YOffset < -8) Then anim = Map.MapEvents(id).Step
                Case DIR_LEFT
                    If (Map.MapEvents(id).XOffset > 8) Then anim = Map.MapEvents(id).Step
                Case DIR_RIGHT
                    If (Map.MapEvents(id).XOffset < -8) Then anim = Map.MapEvents(id).Step
            End Select
            
            ' Set the left
            Select Case Map.MapEvents(id).ShowDir
                Case DIR_UP
                    spritetop = 3
                Case DIR_RIGHT
                    spritetop = 2
                Case DIR_DOWN
                    spritetop = 0
                Case DIR_LEFT
                    spritetop = 1
            End Select
            
            If Map.MapEvents(id).WalkAnim = 1 Then anim = 0
            
            If Map.MapEvents(id).Moving = 0 Then anim = Map.MapEvents(id).GraphicX
            
            With sRECT
                .top = spritetop * height
                .Bottom = .top + height
                .Left = anim * width
                .Right = .Left + width
            End With
        
            ' Calculate the X
            X = Map.MapEvents(id).X * PIC_X + Map.MapEvents(id).XOffset - ((width - 32) / 2)
        
            ' Is the player's height more than 32..?
            If (height * 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                Y = Map.MapEvents(id).Y * PIC_Y + Map.MapEvents(id).YOffset - ((height) - 32)
            Else
                ' Proceed as normal
                Y = Map.MapEvents(id).Y * PIC_Y + Map.MapEvents(id).YOffset
            End If
        
            ' render the actual sprite
            Call BltSprite(Map.MapEvents(id).GraphicNum, X, Y, sRECT)
            
        Case 2
            If Map.MapEvents(id).GraphicNum < 1 Or Map.MapEvents(id).GraphicNum > NumTileSets Then Exit Sub
            
            If Map.MapEvents(id).GraphicY2 > 0 Or Map.MapEvents(id).GraphicX2 > 0 Then
                With sRECT
                    .top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .top + ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) * 32)
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + ((Map.MapEvents(id).GraphicX2 - Map.MapEvents(id).GraphicX) * 32)
                End With
            Else
                With sRECT
                    .top = Map.MapEvents(id).GraphicY * 32
                    .Bottom = .top + 32
                    .Left = Map.MapEvents(id).GraphicX * 32
                    .Right = .Left + 32
                End With
            End If
            
            X = Map.MapEvents(id).X * 32
            Y = Map.MapEvents(id).Y * 32
            
            X = X - ((sRECT.Right - sRECT.Left) / 2)
            Y = Y - (sRECT.Bottom - sRECT.top) + 32
            
            
            If DDS_Tileset(Map.MapEvents(id).GraphicNum) Is Nothing Then
                Call InitDDSurf("tilesets\" & Map.MapEvents(id).GraphicNum, DDSD_Tileset(Map.MapEvents(id).GraphicNum), DDS_Tileset(Map.MapEvents(id).GraphicNum))
            End If
            If Map.MapEvents(id).GraphicY2 > 0 Then
                Call Engine_BltFast(ConvertMapX(Map.MapEvents(id).X * 32), ConvertMapY((Map.MapEvents(id).Y - ((Map.MapEvents(id).GraphicY2 - Map.MapEvents(id).GraphicY) - 1)) * 32), DDS_Tileset(Map.MapEvents(id).GraphicNum), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                Call Engine_BltFast(ConvertMapX(Map.MapEvents(id).X * 32), ConvertMapY(Map.MapEvents(id).Y * 32), DDS_Tileset(Map.MapEvents(id).GraphicNum), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
    End Select
End Sub
