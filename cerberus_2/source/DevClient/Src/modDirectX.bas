Attribute VB_Name = "modDirectX"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Public Sub InitDirectX()

    ' Initialize direct draw
    Set DD = DX.DirectDrawCreate("")
    frmCClient.Show
    
    With DD
        ' Indicate windows mode application
        Call .SetCooperativeLevel(frmCClient.hWnd, DDSCL_NORMAL)
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
    DD_Clip.SetHWnd frmCClient.picScreen.hWnd
        
    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip
        
    ' Initialize all surfaces
    Call InitSurfaces
End Sub

Public Sub InitSurfaces()
Dim Key As DDCOLORKEY
Dim FileName As String

    ' Set path prefix
    FileName = App.Path & GFX_PATH
    
    ' Check for files existing
    If FileExist(FileName & "sprites" & GFX_EXT, True) = False Or FileExist(FileName & "bigsprites" & GFX_EXT, True) = False Or FileExist(FileName & "treesprites" & GFX_EXT, True) = False Or FileExist(FileName & "tiles" & GFX_EXT, True) = False Or FileExist(FileName & "items" & GFX_EXT, True) = False Or FileExist(FileName & "Direction" & GFX_EXT, True) = False Then
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
    
    ' Init big sprite ddsd type and load the bitmap
    With DDSD_BigSprite
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End With
    Set DD_BigSpriteSurf = DD.CreateSurfaceFromFile(FileName & "bigsprites" & GFX_EXT, DDSD_BigSprite)
    DD_BigSpriteSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    ' Init tree sprite ddsd type and load the bitmap
    With DDSD_TreeSprite
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End With
    Set DD_TreeSpriteSurf = DD.CreateSurfaceFromFile(FileName & "treesprites" & GFX_EXT, DDSD_TreeSprite)
    DD_TreeSpriteSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    '' Init building sprite ddsd type and load the bitmap
    'With DDSD_BuildingSprite
        '.lFlags = DDSD_CAPS
        '.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    'End With
    'Set DD_BuildingSpriteSurf = DD.CreateSurfaceFromFile(FileName & "buildingsprites" & GFX_EXT, DDSD_BuildingSprite)
    'DD_BuildingSpriteSurf.SetColorKey DDCKEY_SRCBLT, Key
    
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
    
    ' Init spells ddsd type and load the bitmap
    With DDSD_Item
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End With
    Set DD_SpellSurf = DD.CreateSurfaceFromFile(FileName & "spells" & GFX_EXT, DDSD_Item)
    DD_SpellSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    ' Init skills ddsd type and load the bitmap
    With DDSD_Item
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End With
    Set DD_SkillSurf = DD.CreateSurfaceFromFile(FileName & "skills" & GFX_EXT, DDSD_Item)
    DD_SkillSurf.SetColorKey DDCKEY_SRCBLT, Key
    
    ' Init direction ddsd type and load the bitmap
    With DDSD_Direction
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End With
    Set DD_DirectionSurf = DD.CreateSurfaceFromFile(FileName & "Direction" & GFX_EXT, DDSD_Direction)
    DD_DirectionSurf.SetColorKey DDCKEY_SRCBLT, Key
End Sub

Sub DestroyDirectX()
    Set DX = Nothing
    Set DD = Nothing
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_BigSpriteSurf = Nothing
    Set DD_TreeSpriteSurf = Nothing
    'Set DD_BuildingSpriteSurf = Nothing
    Set DD_TileSurf = Nothing
    Set DD_ItemSurf = Nothing
    Set DD_SpellSurf = Nothing
    Set DD_SkillSurf = Nothing
    Set DD_DirectionSurf = Nothing
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
