Attribute VB_Name = "modDirectX"
Option Explicit

Public dx As New DirectX7

Public DD As DirectDraw7
Public DD_PrimarySurf As DirectDrawSurface7
Public DD_SpriteSurf As DirectDrawSurface7
Public DD_TileSurf(MAX_TILE_SHEETS) As DirectDrawSurface7
Public DD_TileSurf_MINI(MAX_TILE_SHEETS) As DirectDrawSurface7
Public DD_ItemSurf As DirectDrawSurface7
Public DD_BackBuffer As DirectDrawSurface7
Public DD_BackBuffer_minimap As DirectDrawSurface7
Public DD_NIGHTSURF As DirectDrawSurface7
Public DD_Clip As DirectDrawClipper


Public DDSD_Primary As DDSURFACEDESC2
Public DDSD_Sprite As DDSURFACEDESC2
Public DDSD_Tile(MAX_TILE_SHEETS) As DDSURFACEDESC2
Public DDSD_Tile_MINI(MAX_TILE_SHEETS) As DDSURFACEDESC2
Public DDSD_Item As DDSURFACEDESC2
Public DDSD_BackBuffer As DDSURFACEDESC2
Public DDSD_BackBuffer_minimap As DDSURFACEDESC2
Public DDSD_Night As DDSURFACEDESC2



Public rec As RECT
Public rec_pos As RECT


Sub InitDirectX()
Dim i As Long
    notInGame = False
    ' Initialize direct draw
    Set DD = dx.DirectDrawCreate("")
    frmMirage.Show
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hwnd, DDSCL_NORMAL)
    
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
        
    ' Initialize all surfaces
    Call InitSurfaces
    
End Sub

Sub InitSurfaces()
Dim key As DDCOLORKEY
Dim i As Long

    ' Check for files existing
    If FileExist("\data\bmp\sprites.bmp") = False Or FileExist("\data\bmp\tiles0.bmp") = False Or FileExist("\data\bmp\items.bmp") = False Then
        Call MsgBox("You dont have the graphics files in the same directory as this executable!", vbOKOnly, GAME_NAME)
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
    
    ' Initialize back buffer
    DDSD_BackBuffer_minimap.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer_minimap.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer_minimap.lWidth = 226
    DDSD_BackBuffer_minimap.lHeight = 170
    Set DD_BackBuffer_minimap = DD.CreateSurface(DDSD_BackBuffer_minimap)
    
    
    
    ' Init sprite ddsd type and load the bitmap
    DDSD_Sprite.lFlags = DDSD_CAPS
    DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\data\bmp\sprites.bmp", DDSD_Sprite)
    DD_SpriteSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init tiles ddsd type and load the bitmap
    'This will now allow for more than one tile sheet
    For i = 0 To MAX_TILE_SHEETS - 1 Step 1
        DDSD_Tile(i).lFlags = DDSD_CAPS
        DDSD_Tile(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_TileSurf(i) = DD.CreateSurfaceFromFile(App.Path & "\data\bmp\tiles" & i & ".bmp", DDSD_Tile(i))
        DD_TileSurf(i).SetColorKey DDCKEY_SRCBLT, key
        frmMirage.cmbTilePack.AddItem "Tile Pack " & i
        
'        DDSD_Tile_MINI(i).lFlags = DDSD_CAPS
'        DDSD_Tile_MINI(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
'        Set DD_TileSurf_MINI(i) = DD.CreateSurfaceFromFile(App.Path & "\data\bmp\mini\tiles" & i & ".bmp", DDSD_Tile(i))
'        DD_TileSurf_MINI(i).SetColorKey DDCKEY_SRCBLT, key
        
    Next i
    frmMirage.cmbTilePack.ListIndex = 0
    
    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(App.Path & "\data\bmp\items.bmp", DDSD_Item)
    DD_ItemSurf.SetColorKey DDCKEY_SRCBLT, key
    
    'NIGHT
    'key.low = RGB(0, 255, 0)
    'key.high = RGB(0, 255, 0)
    
    DDSD_Night.lFlags = DDSD_CAPS
    DDSD_Night.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_NIGHTSURF = DD.CreateSurfaceFromFile(App.Path & "\data\bmp\night.bmp", DDSD_Night)
    DD_NIGHTSURF.SetColorKey DDCKEY_SRCBLT, key
    
    'key.low = 0
    'key.high = 0
End Sub

Sub DestroyDirectX()
Dim i As Long
    Set dx = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    For i = 0 To MAX_TILE_SHEETS - 1
        Set DD_TileSurf(i) = Nothing
    Next i
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

