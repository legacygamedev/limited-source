Attribute VB_Name = "modDirectX"
Option Explicit

Public Function InitDirectX7() As Boolean
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
    InitDirectX7 = False
    
    ' Initialize direct draw
    Set DD = DX7.DirectDrawCreate("")
    frmDualSolace.Show
    
    With DD
        ' Indicate windows mode application
        Call .SetCooperativeLevel(frmDualSolace.hWnd, DDSCL_NORMAL)
    End With
    
        
    With DDSD_Primary
        ' Init type and get the primary surface
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    End With
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmDualSolace.picScreen.hWnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
        
    ' Initialize all surfaces
    While Not InitSurfaces
    Wend
    InitDirectX7 = True
End Function

Public Function InitSurfaces() As Boolean
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function, added gfx constants.
'****************************************************************
On Error GoTo SysSurf
InitSurfaces = False

Dim Key As DDCOLORKEY
Dim FileName As String

    ' Set path prefix
    FileName = App.Path & GFX_PATH
    
    ' Check for files existing
    If FileExist(FileName & "sprites" & GFX_EXT, True) = False Or FileExist(FileName & "tiles" & GFX_EXT, True) = False Or FileExist(FileName & "items" & GFX_EXT, True) = False Then
        Call MsgBox("You dont have the graphics files in the " & FileName & GFX_PATH & " directory!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If
        
    ' Set the key for masks
    With Key
        .low = 0
        .high = 0
    End With
    
    ' Initialize back buffer
    'DDSCAPS_VIDEOMEMORY Or
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
        .lWidth = (MAX_MAPX + 1) * PIC_X
        .lHeight = (MAX_MAPY + 1) * PIC_Y
    End With
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' Init sprite ddsd type and load the bitmap
    With DDSD_Sprite
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End With
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(FileName & "sprites" & GFX_EXT, DDSD_Sprite)
    DD_SpriteSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    ' Init tiles ddsd type and load the bitmap
    With DDSD_Tile
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End With
    Set DD_TileSurf = DD.CreateSurfaceFromFile(FileName & "tiles" & GFX_EXT, DDSD_Tile)
    DD_TileSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    ' Init items ddsd type and load the bitmap
    With DDSD_Item
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End With
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(FileName & "items" & GFX_EXT, DDSD_Item)
    DD_ItemSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    ' Init arrows ddsd type and load the bitmap
    'With DDSD_ArrowAnim
    '    .lFlags = DDSD_CAPS
    '    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    'End With
    'Set DD_ArrowAnim = DD.CreateSurfaceFromFile(App.Path & "\arrows.bmp", DDSD_ArrowAnim)
    'DD_ArrowAnim.SetColorKey DDCKEY_SRCBLT, Key
InitSurfaces = True

Exit Function

SysSurf:
Call InitSysSurf
End Function

Sub DestroyDirectX7()
    Set DX7 = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_TileSurf = Nothing
    Set DD_ItemSurf = Nothing
    'Set DD_ArrowAnim = Nothing
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

Sub InitSysSurf()
On Error GoTo SysSurfError
Dim Key As DDCOLORKEY
Dim FileName As String

    ' Set path prefix
    FileName = App.Path & GFX_PATH
    
    ' Check for files existing
    If FileExist(FileName & "sprites" & GFX_EXT, True) = False Or FileExist(FileName & "tiles" & GFX_EXT, True) = False Or FileExist(FileName & "items" & GFX_EXT, True) = False Then
        Call MsgBox("You dont have the graphics files in the " & FileName & GFX_PATH & " directory!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If
        
    ' Set the key for masks
    With Key
        .low = 0
        .high = 0
    End With
    
    ' Initialize back buffer
    'DDSCAPS_VIDEOMEMORY Or
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lWidth = (MAX_MAPX + 1) * PIC_X
        .lHeight = (MAX_MAPY + 1) * PIC_Y
    End With
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' Init sprite ddsd type and load the bitmap
    With DDSD_Sprite
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End With
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(FileName & "sprites" & GFX_EXT, DDSD_Sprite)
    DD_SpriteSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    ' Init tiles ddsd type and load the bitmap
    With DDSD_Tile
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End With
    Set DD_TileSurf = DD.CreateSurfaceFromFile(FileName & "tiles" & GFX_EXT, DDSD_Tile)
    DD_TileSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    ' Init items ddsd type and load the bitmap
    With DDSD_Item
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End With
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(FileName & "items" & GFX_EXT, DDSD_Item)
    DD_ItemSurf.SetColorKey DDCKEY_SRCBLT, Key
Exit Sub
    
SysSurfError:
MsgBox "There seems to be an error loading the graphics into memory through both your video memory and your system memory!"
Call GameDestroy
End Sub
