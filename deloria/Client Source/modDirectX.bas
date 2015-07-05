Attribute VB_Name = "modDirectX"
Option Explicit

Public DX As New DirectX7
Public DD As DirectDraw7
Public DD_PrimarySurf As DirectDrawSurface7
Public DD_SpriteSurf As DirectDrawSurface7
Public DD_MonsterSurf As DirectDrawSurface7
Public DD_TileSurf As DirectDrawSurface7
Public DD_ItemSurf As DirectDrawSurface7
Public DD_EmoticonSurf As DirectDrawSurface7
Public DD_WeatherSurf As DirectDrawSurface7
Public DD_BackBuffer As DirectDrawSurface7
Public DD_Clip As DirectDrawClipper

Public DDSD_Primary As DDSURFACEDESC2
Public DDSD_Sprite As DDSURFACEDESC2
Public DDSD_Tile As DDSURFACEDESC2
Public DDSD_Item As DDSURFACEDESC2
Public DDSD_Emoticon As DDSURFACEDESC2
Public DDSD_Weather As DDSURFACEDESC2
Public DDSD_BackBuffer As DDSURFACEDESC2

Public DD_BigSpriteSurf As DirectDrawSurface7
Public DDSD_BigSprite As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

Public Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Type JOYINFO
    wXpos As Long
    wYpos As Long
    wZpos As Long
    wButtons As Long
End Type

Sub InitDirectX()

    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
    frmMirage.Show
    
    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hWnd, DDSCL_NORMAL)
    
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
        
    ' Initialize all surfaces
    Call InitSurfaces
End Sub

Sub InitSurfaces()
Dim key As DDCOLORKEY

    ' Check for files existing
    If FileExist("\GFX\sprites.bmp") = False Or FileExist("\GFX\tiles.bmp") = False Or FileExist("\GFX\items.bmp") = False Or FileExist("\GFX\bigsprites.bmp") = False Or FileExist("\GFX\emoticons.bmp") = False Or FileExist("\GFX\weather.bmp") = False Then
        Call MsgBox("Your missing some graphic files!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If
    
    ' Set the key for masks
    key.Low = 0
    key.High = 0
    
    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = 3 * ((MAX_MAPX + 1) * PIC_X)
    DDSD_BackBuffer.lHeight = 3 * ((MAX_MAPY + 1) * PIC_Y)
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' Init sprite ddsd type and load the bitmap
    DDSD_Sprite.lFlags = DDSD_CAPS
    DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\sprites.bmp", DDSD_Sprite)
    DD_SpriteSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init tiles ddsd type and load the bitmap
    DDSD_Tile.lFlags = DDSD_CAPS
    DDSD_Tile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_TileSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\tiles.bmp", DDSD_Tile)
    DD_TileSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\items.bmp", DDSD_Item)
    DD_ItemSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init big sprites ddsd type and load the bitmap
    DDSD_BigSprite.lFlags = DDSD_CAPS
    DDSD_BigSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_BigSpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\bigsprites.bmp", DDSD_BigSprite)
    DD_BigSpriteSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init emoticons ddsd type and load the bitmap
    DDSD_Emoticon.lFlags = DDSD_CAPS
    DDSD_Emoticon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_EmoticonSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\emoticons.bmp", DDSD_Emoticon)
    DD_EmoticonSurf.SetColorKey DDCKEY_SRCBLT, key
    
    ' Init weather ddsd type and load the bitmap
    DDSD_Weather.lFlags = DDSD_CAPS
    DDSD_Weather.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_WeatherSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\weather.bmp", DDSD_Weather)
    DD_WeatherSurf.SetColorKey DDCKEY_SRCBLT, key
End Sub

Sub DestroyDirectX()
    Set DX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_TileSurf = Nothing
    Set DD_ItemSurf = Nothing
    Set DD_BigSpriteSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    Set DD_WeatherSurf = Nothing
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
