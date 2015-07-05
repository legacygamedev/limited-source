Attribute VB_Name = "modAutotiles"
' ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ::                Autotile System by EdTheHobo (Version 1.0)                ::
' :: No engines other than the Chaos Engine have permission to use this code. ::
' ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Option Explicit

Public Sub AutotileMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim TL, TR, BL, BR, T, B, L, R As Byte
Dim x1, y1, x2, y2, picx As Long, TempEY As Long, TempEX As Long

    If InEditor Then
        x1 = Int(x / PIC_X)
        y1 = Int(y / PIC_Y)
        If frmMapEditor.optGround.Value = True Then
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Ground = EditorTileY * TilesInSheets + EditorTileX
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).GroundSet = EditorSet
            Call GroundAutotile(x1, y1)
        End If
        If frmMapEditor.optMask.Value = True Then
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask = EditorTileY * TilesInSheets + EditorTileX
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet = EditorSet
            Call MaskAutotile(x1, y1)
            If frmMapEditor.chkAutoAnim.Value > 0 Then
                TempEY = EditorTileY: TempEX = EditorTileX
                EditorTileY = EditorTileY + 6
                Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet = EditorSet
                Call AnimAutotile(x1, y1)
                EditorTileY = TempEY: EditorTileX = TempEX
            End If
        End If
        If frmMapEditor.optAnim.Value = True Then
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Anim = EditorTileY * TilesInSheets + EditorTileX
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet = EditorSet
            Call AnimAutotile(x1, y1)
            If frmMapEditor.chkAutoAnim.Value > 0 Then
                TempEY = EditorTileY: TempEX = EditorTileX
                EditorTileY = EditorTileY + 6
                Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet = EditorSet
                Call MaskAutotile(x1, y1)
                EditorTileY = TempEY: EditorTileX = TempEX
            End If
        End If
        If frmMapEditor.optMask2.Value = True Then
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2 = EditorTileY * TilesInSheets + EditorTileX
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set = EditorSet
            Call Mask2Autotile(x1, y1)
            If frmMapEditor.chkAutoAnim.Value > 0 Then
                TempEY = EditorTileY: TempEX = EditorTileX
                EditorTileY = EditorTileY + 6
                Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet = EditorSet
                Call M2AnimAutotile(x1, y1)
                EditorTileY = TempEY: EditorTileX = TempEX
            End If
        End If
        If frmMapEditor.optM2Anim.Value = True Then
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2Anim = EditorTileY * TilesInSheets + EditorTileX
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet = EditorSet
            Call M2AnimAutotile(x1, y1)
            If frmMapEditor.chkAutoAnim.Value > 0 Then
                TempEY = EditorTileY: TempEX = EditorTileX
                EditorTileY = EditorTileY + 6
                Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set = EditorSet
                Call Mask2Autotile(x1, y1)
                EditorTileY = TempEY: EditorTileX = TempEX
            End If
        End If
        If frmMapEditor.optFringe.Value = True Then
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe = EditorTileY * TilesInSheets + EditorTileX
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet = EditorSet
            Call FringeAutotile(x1, y1)
            If frmMapEditor.chkAutoAnim.Value > 0 Then
                TempEY = EditorTileY: TempEX = EditorTileX
                EditorTileY = EditorTileY + 6
                Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet = EditorSet
                Call FAnimAutotile(x1, y1)
                EditorTileY = TempEY: EditorTileX = TempEX
            End If
        End If
        If frmMapEditor.optFAnim.Value = True Then
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnim = EditorTileY * TilesInSheets + EditorTileX
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet = EditorSet
            Call FAnimAutotile(x1, y1)
            If frmMapEditor.chkAutoAnim.Value > 0 Then
                TempEY = EditorTileY: TempEX = EditorTileX
                EditorTileY = EditorTileY + 6
                Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet = EditorSet
                Call FringeAutotile(x1, y1)
                EditorTileY = TempEY: EditorTileX = TempEX
            End If
        End If
        If frmMapEditor.optFringe2.Value = True Then
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2 = EditorTileY * TilesInSheets + EditorTileX
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set = EditorSet
            Call Fringe2Autotile(x1, y1)
            If frmMapEditor.chkAutoAnim.Value > 0 Then
                TempEY = EditorTileY: TempEX = EditorTileX
                EditorTileY = EditorTileY + 6
                Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet = EditorSet
                Call F2AnimAutotile(x1, y1)
                EditorTileY = TempEY: EditorTileX = TempEX
            End If
        End If
        If frmMapEditor.optF2Anim.Value = True Then
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2Anim = EditorTileY * TilesInSheets + EditorTileX
            Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet = EditorSet
            Call F2AnimAutotile(x1, y1)
            If frmMapEditor.chkAutoAnim.Value > 0 Then
                TempEY = EditorTileY: TempEX = EditorTileX
                EditorTileY = EditorTileY + 6
                Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set = EditorSet
                Call Fringe2Autotile(x1, y1)
                EditorTileY = TempEY: EditorTileX = TempEX
            End If
        End If
    End If
End Sub

Sub GroundAutotile(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' This is the easy way to do it, you should have seen the original code D:
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).GroundSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).GroundSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
If y1 < MAX_MAPY Then
    ' Bottom-Left Tile:
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).GroundSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).GroundSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).GroundSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).GroundSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).GroundSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).GroundSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Ground = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L)
'Exit Sub
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).GroundSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).GroundSet Then Call GroundAutotile2(x1 - 1, y1)
End If
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).GroundSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).GroundSet Then Call GroundAutotile2(x1, y1 - 1)
End If
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).GroundSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).GroundSet Then Call GroundAutotile2(x1 + 1, y1)
End If
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).GroundSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).GroundSet Then Call GroundAutotile2(x1, y1 + 1)
End If
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).GroundSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).GroundSet Then Call GroundAutotile2(x1 - 1, y1 - 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).GroundSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).GroundSet Then Call GroundAutotile2(x1 + 1, y1 - 1)
    End If
End If
If x1 > 0 Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).GroundSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).GroundSet Then Call GroundAutotile2(x1 - 1, y1 + 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).GroundSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).GroundSet Then Call GroundAutotile2(x1 + 1, y1 + 1)
    End If
End If
End Sub

Sub GroundAutotile2(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).GroundSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).GroundSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
' Bottom-Left Tile:
If y1 < MAX_MAPY Then
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).GroundSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).GroundSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).GroundSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).GroundSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).GroundSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).GroundSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Ground = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L, x1, y1)
End Sub

Sub MaskAutotile(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).MaskSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).MaskSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
If y1 < MAX_MAPY Then
    ' Bottom-Left Tile:
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).MaskSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).MaskSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).MaskSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).MaskSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).MaskSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).MaskSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L)
'Exit Sub
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).MaskSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet Then Call maskAutotile2(x1 - 1, y1)
End If
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).MaskSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet Then Call maskAutotile2(x1, y1 - 1)
End If
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).MaskSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet Then Call maskAutotile2(x1 + 1, y1)
End If
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).MaskSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet Then Call maskAutotile2(x1, y1 + 1)
End If
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).MaskSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet Then Call maskAutotile2(x1 - 1, y1 - 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).MaskSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet Then Call maskAutotile2(x1 + 1, y1 - 1)
    End If
End If
If x1 > 0 Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).MaskSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet Then Call maskAutotile2(x1 - 1, y1 + 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).MaskSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).MaskSet Then Call maskAutotile2(x1 + 1, y1 + 1)
    End If
End If
End Sub

Sub maskAutotile2(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).MaskSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).MaskSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
' Bottom-Left Tile:
If y1 < MAX_MAPY Then
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).MaskSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).MaskSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).MaskSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).MaskSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).MaskSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).MaskSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L, x1, y1)
End Sub

Sub AnimAutotile(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).AnimSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).AnimSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
If y1 < MAX_MAPY Then
    ' Bottom-Left Tile:
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).AnimSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).AnimSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).AnimSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).AnimSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).AnimSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).AnimSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Anim = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L)
'Exit Sub
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet Then Call AnimAutotile2(x1 - 1, y1)
End If
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet Then Call AnimAutotile2(x1, y1 - 1)
End If
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet Then Call AnimAutotile2(x1 + 1, y1)
End If
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet Then Call AnimAutotile2(x1, y1 + 1)
End If
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet Then Call AnimAutotile2(x1 - 1, y1 - 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet Then Call AnimAutotile2(x1 + 1, y1 - 1)
    End If
End If
If x1 > 0 Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet Then Call AnimAutotile2(x1 - 1, y1 + 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).AnimSet Then Call AnimAutotile2(x1 + 1, y1 + 1)
    End If
End If
End Sub

Sub AnimAutotile2(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).AnimSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).AnimSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
' Bottom-Left Tile:
If y1 < MAX_MAPY Then
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).AnimSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).AnimSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).AnimSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).AnimSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).AnimSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).AnimSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Anim = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L, x1, y1)
End Sub

Sub Mask2Autotile(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).Mask2Set <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).Mask2Set <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
If y1 < MAX_MAPY Then
    ' Bottom-Left Tile:
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).Mask2Set <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).Mask2Set <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).Mask2Set <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).Mask2Set <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).Mask2Set <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).Mask2Set <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2 = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L)
'Exit Sub
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).Mask2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set Then Call Mask2Autotile2(x1 - 1, y1)
End If
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).Mask2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set Then Call Mask2Autotile2(x1, y1 - 1)
End If
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).Mask2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set Then Call Mask2Autotile2(x1 + 1, y1)
End If
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).Mask2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set Then Call Mask2Autotile2(x1, y1 + 1)
End If
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).Mask2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set Then Call Mask2Autotile2(x1 - 1, y1 - 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).Mask2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set Then Call Mask2Autotile2(x1 + 1, y1 - 1)
    End If
End If
If x1 > 0 Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).Mask2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set Then Call Mask2Autotile2(x1 - 1, y1 + 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).Mask2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2Set Then Call Mask2Autotile2(x1 + 1, y1 + 1)
    End If
End If
End Sub

Sub Mask2Autotile2(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).Mask2Set <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).Mask2Set <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
' Bottom-Left Tile:
If y1 < MAX_MAPY Then
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).Mask2Set <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).Mask2Set <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).Mask2Set <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).Mask2Set <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).Mask2Set <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).Mask2Set <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Mask2 = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L, x1, y1)
End Sub

Sub M2AnimAutotile(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).M2AnimSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).M2AnimSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
If y1 < MAX_MAPY Then
    ' Bottom-Left Tile:
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).M2AnimSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).M2AnimSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).M2AnimSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).M2AnimSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).M2AnimSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).M2AnimSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2Anim = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L)
'Exit Sub
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).M2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet Then Call M2AnimAutotile2(x1 - 1, y1)
End If
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).M2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet Then Call M2AnimAutotile2(x1, y1 - 1)
End If
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).M2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet Then Call M2AnimAutotile2(x1 + 1, y1)
End If
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).M2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet Then Call M2AnimAutotile2(x1, y1 + 1)
End If
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).M2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet Then Call M2AnimAutotile2(x1 - 1, y1 - 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).M2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet Then Call M2AnimAutotile2(x1 + 1, y1 - 1)
    End If
End If
If x1 > 0 Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).M2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet Then Call M2AnimAutotile2(x1 - 1, y1 + 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).M2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2AnimSet Then Call M2AnimAutotile2(x1 + 1, y1 + 1)
    End If
End If
End Sub

Sub M2AnimAutotile2(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).M2AnimSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).M2AnimSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
' Bottom-Left Tile:
If y1 < MAX_MAPY Then
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).M2AnimSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).M2AnimSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).M2AnimSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).M2AnimSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).M2AnimSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).M2AnimSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).M2Anim = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L, x1, y1)
End Sub

Sub FringeAutotile(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).FringeSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).FringeSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
If y1 < MAX_MAPY Then
    ' Bottom-Left Tile:
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).FringeSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).FringeSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).FringeSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).FringeSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).FringeSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).FringeSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L)
'Exit Sub
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).FringeSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet Then Call FringeAutotile2(x1 - 1, y1)
End If
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).FringeSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet Then Call FringeAutotile2(x1, y1 - 1)
End If
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).FringeSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet Then Call FringeAutotile2(x1 + 1, y1)
End If
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).FringeSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet Then Call FringeAutotile2(x1, y1 + 1)
End If
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).FringeSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet Then Call FringeAutotile2(x1 - 1, y1 - 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).FringeSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet Then Call FringeAutotile2(x1 + 1, y1 - 1)
    End If
End If
If x1 > 0 Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).FringeSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet Then Call FringeAutotile2(x1 - 1, y1 + 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).FringeSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FringeSet Then Call FringeAutotile2(x1 + 1, y1 + 1)
    End If
End If
End Sub

Sub FringeAutotile2(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).FringeSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).FringeSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
' Bottom-Left Tile:
If y1 < MAX_MAPY Then
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).FringeSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).FringeSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).FringeSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).FringeSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).FringeSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).FringeSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L, x1, y1)
End Sub

Sub FAnimAutotile(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).FAnimSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).FAnimSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
If y1 < MAX_MAPY Then
    ' Bottom-Left Tile:
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).FAnimSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).FAnimSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).FAnimSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).FAnimSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).FAnimSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).FAnimSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnim = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L)
'Exit Sub
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).FAnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet Then Call FAnimAutotile2(x1 - 1, y1)
End If
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).FAnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet Then Call FAnimAutotile2(x1, y1 - 1)
End If
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).FAnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet Then Call FAnimAutotile2(x1 + 1, y1)
End If
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).FAnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet Then Call FAnimAutotile2(x1, y1 + 1)
End If
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).FAnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet Then Call FAnimAutotile2(x1 - 1, y1 - 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).FAnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet Then Call FAnimAutotile2(x1 + 1, y1 - 1)
    End If
End If
If x1 > 0 Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).FAnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet Then Call FAnimAutotile2(x1 - 1, y1 + 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).FAnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnimSet Then Call FAnimAutotile2(x1 + 1, y1 + 1)
    End If
End If
End Sub

Sub FAnimAutotile2(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).FAnimSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).FAnimSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
' Bottom-Left Tile:
If y1 < MAX_MAPY Then
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).FAnimSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).FAnimSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).FAnimSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).FAnimSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).FAnimSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).FAnimSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).FAnim = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L, x1, y1)
End Sub

Sub Fringe2Autotile(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).Fringe2Set <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).Fringe2Set <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
If y1 < MAX_MAPY Then
    ' Bottom-Left Tile:
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).Fringe2Set <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).Fringe2Set <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).Fringe2Set <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).Fringe2Set <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).Fringe2Set <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).Fringe2Set <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2 = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L)
'Exit Sub
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).Fringe2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set Then Call Fringe2Autotile2(x1 - 1, y1)
End If
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).Fringe2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set Then Call Fringe2Autotile2(x1, y1 - 1)
End If
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).Fringe2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set Then Call Fringe2Autotile2(x1 + 1, y1)
End If
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).Fringe2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set Then Call Fringe2Autotile2(x1, y1 + 1)
End If
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).Fringe2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set Then Call Fringe2Autotile2(x1 - 1, y1 - 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).Fringe2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set Then Call Fringe2Autotile2(x1 + 1, y1 - 1)
    End If
End If
If x1 > 0 Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).Fringe2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set Then Call Fringe2Autotile2(x1 - 1, y1 + 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).Fringe2Set = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2Set Then Call Fringe2Autotile2(x1 + 1, y1 + 1)
    End If
End If
End Sub

Sub Fringe2Autotile2(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).Fringe2Set <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).Fringe2Set <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
' Bottom-Left Tile:
If y1 < MAX_MAPY Then
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).Fringe2Set <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).Fringe2Set <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).Fringe2Set <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).Fringe2Set <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).Fringe2Set <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).Fringe2Set <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Fringe2 = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L, x1, y1)
End Sub

Sub F2AnimAutotile(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).F2AnimSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).F2AnimSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
If y1 < MAX_MAPY Then
    ' Bottom-Left Tile:
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).F2AnimSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).F2AnimSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).F2AnimSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).F2AnimSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).F2AnimSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).F2AnimSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2Anim = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L)
'Exit Sub
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).F2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet Then Call F2AnimAutotile2(x1 - 1, y1)
End If
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).F2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet Then Call F2AnimAutotile2(x1, y1 - 1)
End If
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).F2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet Then Call F2AnimAutotile2(x1 + 1, y1)
End If
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).F2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet Then Call F2AnimAutotile2(x1, y1 + 1)
End If
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).F2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet Then Call F2AnimAutotile2(x1 - 1, y1 - 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).F2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet Then Call F2AnimAutotile2(x1 + 1, y1 - 1)
    End If
End If
If x1 > 0 Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).F2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet Then Call F2AnimAutotile2(x1 - 1, y1 + 1)
    End If
End If
If x1 < MAX_MAPX Then
    If y1 < MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).F2AnimSet = Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2AnimSet Then Call F2AnimAutotile2(x1 + 1, y1 + 1)
    End If
End If
End Sub

Sub F2AnimAutotile2(ByVal x1 As Long, ByVal y1 As Long)
Dim TL, T, TR, r, BR, B, BL, L As Byte
' First let's see if we can put tiles on the side (Make sure it doesn't go outside the map border)
If x1 = 0 Then:        TL = 1: L = 1: BL = 1
If x1 = MAX_MAPX Then: TR = 1: r = 1: BR = 1
If y1 = 0 Then:        TL = 1: T = 1: TR = 1
If y1 = MAX_MAPY Then: BL = 1: B = 1: BR = 1
' We'll fill in the corners first, then do the sides
' Top-Left Tile:
If x1 > 0 Then
    If y1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 - 1).F2AnimSet <> EditorSet Then
        Else
            TL = 1
        End If
    End If
End If
'Top-Right Tile:
If y1 > 0 Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 - 1).F2AnimSet <> EditorSet Then
        Else
            TR = 1
        End If
    End If
End If
' Bottom-Left Tile:
If y1 < MAX_MAPY Then
    If x1 > 0 Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1 + 1).F2AnimSet <> EditorSet Then
        Else
            BL = 1
        End If
    End If
End If
' Bottom-Right Tile
If y1 < MAX_MAPY Then
    If x1 < MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1 + 1).F2AnimSet <> EditorSet Then
        Else
            BR = 1
        End If
    End If
End If
' Based on the corner tiles, we can now decide what the side tiles should be.
' Left Tile:
If x1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 - 1, y1).F2AnimSet <> EditorSet Then
    Else
        L = 1
    End If
End If
' Right Tile
If x1 < MAX_MAPX Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1 + 1, y1).F2AnimSet <> EditorSet Then
    Else
        r = 1
    End If
End If
' Top Tile
If y1 > 0 Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 - 1).F2AnimSet <> EditorSet Then
    Else
        T = 1
    End If
End If
' Bottom Tile
If y1 < MAX_MAPY Then
    If Map(GetPlayerMap(MyIndex)).Tile(x1, y1 + 1).F2AnimSet <> EditorSet Then
    Else
        B = 1
    End If
End If
' And finally, we can set the middle tile.
Map(GetPlayerMap(MyIndex)).Tile(x1, y1).F2Anim = (EditorTileY * TilesInSheets + EditorTileX) + CheckMidAutotile(TL, T, TR, r, BR, B, BL, L, x1, y1)
End Sub

Function CheckMidAutotile(ByVal TL As Byte, ByVal T As Byte, ByVal TR As Byte, ByVal r As Byte, ByVal BR As Byte, ByVal B As Byte, ByVal BL As Byte, ByVal L As Byte, Optional x1 As Long = 0, Optional y1 As Long = 0) As Byte
' Returns the tile number to change the middle autotile to
' This took me, literally, hours to figure out... SO IT BETTER WORK >_<
Dim n As Long ' N = tile number
If T = 1 Then ' 1 = autotile exists at that location, 0 = something else is there
    If r = 1 Then
        If B = 1 Then
            If L = 1 Then
                If TL = 1 Then
                    If TR = 1 Then
                        If BR = 1 Then
                            If BL = 1 Then
                                n = 77
                            Else
                                n = 14
                            End If
                        Else
                            If BL = 1 Then
                                n = 4
                            Else
                                n = 18
                            End If
                        End If
                    Else
                        If BR = 1 Then
                            If BL = 1 Then
                                n = 2
                            Else
                                n = 16
                            End If
                        Else
                            If BL = 1 Then
                                n = 6
                            Else
                                n = 20
                            End If
                        End If
                    End If
                Else 'TL = 0
                    If TR = 1 Then
                        If BR = 1 Then
                            If BL = 1 Then
                                n = 1
                            Else
                                n = 15
                            End If
                        Else
                            If BL = 1 Then
                                n = 5
                            Else
                                n = 19
                            End If
                        End If
                    Else
                        If BR = 1 Then
                            If BL = 1 Then
                                n = 3
                            Else
                                n = 17
                            End If
                        Else
                            If BL = 1 Then
                                n = 7
                            Else
                                n = 21
                            End If
                        End If
                    End If
                End If
            Else 'T,B,R = 1; L = 0
            
                ' If L = 0 then we can ignore TL and BL, right? Right.
                            
                ' Reason: If there is no autotile on the left, there is no way that
                ' the middle and corner tiles can be directly connected (We need both
                ' the Top and Left tiles in order to connect to the Top-Left corner).
                
            
                If TR = 1 Then
                    If BR = 1 Then
                        n = 28
                    Else
                        n = 30
                    End If
                Else
                    If BR = 1 Then
                        n = 29
                    Else
                        n = 31
                    End If
                End If
            End If
        Else 'T, R = 1; B = 0
            ' B = 0, we can ignore BR and BL
            If L = 1 Then
                If TL = 1 Then
                    If TR = 1 Then
                        n = 46
                    Else
                        n = 48
                    End If
                Else 'TL = 0
                    If TR = 1 Then
                        n = 47
                    Else
                        n = 49
                    End If
                End If
            Else
                ' B and L = 0, we only need to look at TR
                If TR = 1 Then
                    n = 70
                Else
                    n = 71
                End If
            End If
        End If
    Else ' T = 1, R = 0
        If B = 1 Then
            If L = 1 Then
                ' R = 0, ignore TR and BR
                If TL = 1 Then
                    If BL = 1 Then
                        n = 42
                    Else
                        n = 43
                    End If
                Else
                    If BL = 1 Then
                        n = 44
                    Else
                        n = 45
                    End If
                End If
            Else
                ' L, R = 0; T, B = 1 -- Ignore all the corners :)
                ' There is only one possible tile that can go here
                n = 56
            End If
        Else 'T = 1; R, B = 0
            If L = 1 Then
                ' Only look at TL
                If TL = 1 Then
                    n = 62
                Else
                    n = 63
                End If
            Else
                ' R, B, L = 0; T = 1
                ' No corners, only one possible tile  (It's getting easy now)
                n = 74
            End If
        End If
    End If
Else ' T = 0
    If r = 1 Then
        If B = 1 Then
            If L = 1 Then
                ' Ignore TL and TR
                If BR = 1 Then
                    If BL = 1 Then
                        n = 32
                    Else
                        n = 34
                    End If
                Else
                    If BL = 1 Then
                        n = 33
                    Else
                        n = 35
                    End If
                End If
            Else ' T, L = 0; R, B = 1
                ' Only look at BR
                If BR = 1 Then
                    n = 58
                Else
                    n = 59
                End If
            End If
        Else ' T, B = 0; R = 1
            If L = 1 Then
                ' Only one possible tile
                n = 57
            Else
                ' T, B, L = 0; R = 1
                ' Again, only one possible tile
                n = 73
            End If
        End If
    Else ' R, T = 0
        If B = 1 Then
            If L = 1 Then
                ' Only look at BL
                If BL = 1 Then
                    n = 60
                Else
                    n = 61
                End If
            Else
                ' T, L, R = 0; Only one possible tile
                n = 72
            End If
        Else ' R, B, T = 0
            If L = 1 Then
                ' Only one possible tile
                n = 75
            Else
                n = 76
            End If
        End If
    End If
End If
CheckMidAutotile = n
End Function
