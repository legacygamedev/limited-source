Attribute VB_Name = "modDirectX"
Option Explicit

Public DX As New DirectX7
Public DD As DirectDraw7
Public DD_PrimarySurf As DirectDrawSurface7
Public DD_SpriteSurf As DirectDrawSurface7
Public DD_TileSurf As DirectDrawSurface7
Public DD_ItemSurf As DirectDrawSurface7
Public DD_BackBuffer As DirectDrawSurface7
Public DD_Clip As DirectDrawClipper

Public DDSD_Primary As DDSURFACEDESC2
Public DDSD_Sprite As DDSURFACEDESC2
Public DDSD_Tile As DDSURFACEDESC2
Public DDSD_Item As DDSURFACEDESC2
Public DDSD_BackBuffer As DDSURFACEDESC2

Public DD_HPSurf As DirectDrawSurface7
Public DD_MPSurf As DirectDrawSurface7

Public DDSD_HP As DDSURFACEDESC2
Public DDSD_MP As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX()

    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
    frmMainGame.Show
    
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMainGame.hWnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMainGame.picScreen.hWnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
        
    ' Initialize all surfaces
    Call InitSurfaces
End Sub

Sub InitSurfaces()
Dim KEY As DDCOLORKEY

    ' Check for files existing
    If FileExist("core files/graphics/Sprites.bmp") = False Or FileExist("core files/graphics/Tiles.bmp") = False Or FileExist("core files/graphics/Items.bmp") = False Then
        Call MsgBox("You dont have the graphics files in the same directory as this executable!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If
    
    ' Set the key for masks
    KEY.low = 0
    KEY.high = 0
    
    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' Init HPs ddsd type and load the bitmap
    DDSD_HP.lFlags = DDSD_CAPS
    DDSD_HP.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_HPSurf = DD.CreateSurfaceFromFile(App.Path & "\core files\interface\HP.bmp", DDSD_HP)
    DD_HPSurf.SetColorKey DDCKEY_SRCBLT, KEY
    
        ' Init MPs ddsd type and load the bitmap
    DDSD_MP.lFlags = DDSD_CAPS
    DDSD_MP.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_MPSurf = DD.CreateSurfaceFromFile(App.Path & "\core files\interface\MP.bmp", DDSD_MP)
    DD_MPSurf.SetColorKey DDCKEY_SRCBLT, KEY
    
    ' Init sprite ddsd type and load the bitmap
    DDSD_Sprite.lFlags = DDSD_CAPS
    DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\core files/graphics/Sprites.bmp", DDSD_Sprite)
    SetMaskColorFromPixel DD_SpriteSurf, 0, 0
    
    ' Init tiles ddsd type and load the bitmap
    DDSD_Tile.lFlags = DDSD_CAPS
    DDSD_Tile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_TileSurf = DD.CreateSurfaceFromFile(App.Path & "\core files/graphics/Tiles.bmp", DDSD_Tile)
    SetMaskColorFromPixel DD_TileSurf, 0, 0
    
    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(App.Path & "\core files/graphics/Items.bmp", DDSD_Item)
    SetMaskColorFromPixel DD_ItemSurf, 0, 0
End Sub

Sub DestroyDirectX()
    Set DX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_TileSurf = Nothing
    Set DD_ItemSurf = Nothing
    Set DD_HPSurf = Nothing
    Set DD_MPSurf = Nothing
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
.top = y
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

