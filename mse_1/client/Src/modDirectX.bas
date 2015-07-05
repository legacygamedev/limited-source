Attribute VB_Name = "modDirectX"
Option Explicit

Public Sub InitDirectX()
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
    
    With DD
        ' Indicate windows mode application
        Call .SetCooperativeLevel(frmMirage.hwnd, DDSCL_NORMAL)
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
    DD_Clip.SetHWnd frmMirage.picScreen.hwnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
        
    ' Initialize all surfaces
    Call InitSurfaces
End Sub

Public Sub InitSurfaces()
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function, added gfx constants.
'****************************************************************

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
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY 'DDSCAPS_SYSTEMMEMORY
        .lWidth = (MAX_MAPX + 1) * PIC_X
        .lHeight = (MAX_MAPY + 1) * PIC_Y
    End With

    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    Set DD_LowerBuffer = DD.CreateSurface(DDSD_BackBuffer)
    Set DD_MiddleBuffer = DD.CreateSurface(DDSD_BackBuffer)
    Set DD_UpperBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    DD_BackBuffer.SetColorKey DDCKEY_SRCBLT, Key
    DD_LowerBuffer.SetColorKey DDCKEY_SRCBLT, Key
    DD_MiddleBuffer.SetColorKey DDCKEY_SRCBLT, Key
    DD_UpperBuffer.SetColorKey DDCKEY_SRCBLT, Key
    
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
End Sub

Sub DestroyDirectX()
    Set DX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_BackBuffer = Nothing
    Set DD_LowerBuffer = Nothing
    Set DD_MiddleBuffer = Nothing
    Set DD_UpperBuffer = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_TileSurf = Nothing
    Set DD_ItemSurf = Nothing
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
