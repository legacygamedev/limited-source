Attribute VB_Name = "modDirectDraw7"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' -- Renders graphics                     --
' ------------------------------------------

' Master Object, leave one commented out
'Public DX7 As DirectX7 ' late binding
Public DX7 As New DirectX7 ' early binding
' ------------------------------------------

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
Public DDS_Mask As DirectDrawSurface7
Public DDS_ItemHighlight As DirectDrawSurface7
Public DDS_Sprite() As DirectDrawSurface7
Public DDS_Anim() As DirectDrawSurface7
Public DDS_SpellIcon As DirectDrawSurface7
Public DDS_Tile() As DirectDrawSurface7
Public DDS_Misc As DirectDrawSurface7
Public DDS_SpellWaitingBar As DirectDrawSurface7

Public DDSD_Item() As DDSURFACEDESC2
Public DDSD_Mask As DDSURFACEDESC2
Public DDSD_ItemHighlight As DDSURFACEDESC2
Public DDSD_Sprite() As DDSURFACEDESC2
Public DDSD_Anim() As DDSURFACEDESC2
Public DDSD_SpellIcon As DDSURFACEDESC2
Public DDSD_Tile() As DDSURFACEDESC2
Public DDSD_Misc As DDSURFACEDESC2
Public DDSD_SpellWaitingBar As DDSURFACEDESC2

' ********************
' ** Initialization **
' ********************
Public Sub InitDirectDraw()
On Error GoTo ErrorHandler
    
    ' Initialize direct draw
    Set DDS_BackBuffer = Nothing 'forces clearing of backbuffer at startup
    
    Set DD = DX7.DirectDrawCreate(vbNullString) ' empty string forces primary device

    ' dictates how we access thescreen and how other programs
    ' running at the same time will be allowed to access the screen as well.
    DD.SetCooperativeLevel frmMainGame.hWnd, DDSCL_NORMAL
    
    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMainGame.picScreen.hWnd
    
    ' Have the blits to the screen clipped to the picture box
    DDS_Primary.SetClipper DD_Clip ' method attaches a clipper object to, or deletes one from, a surface.
 
    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = (MAX_MAPX + 1) * PIC_X
        .lHeight = (MAX_MAPY + 1) * PIC_Y
    End With
    
    '  sets the backbuffer dimensions to picScreen
    frmMainGame.picScreen.Width = DDSD_BackBuffer.lWidth
    frmMainGame.picScreen.Height = DDSD_BackBuffer.lHeight
    
    ' initialize the backbuffer
    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    Exit Sub
    
ErrorHandler:
    
    Err.Raise Err.Number, , "InitDirectDraw failed (" & Err.Description & ")"
    
End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long, Optional ByRef MaskSurface As DirectDrawSurface7)
On Error GoTo ErrorHandler
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

    TmpR = Get_RECT(Y, X, 0, 0)
    
    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0
    
    If Not (MaskSurface Is Nothing) Then
        MaskSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0
    End If
    
    With TmpColorKey
        If Not (MaskSurface Is Nothing) Then
            .low = MaskSurface.GetLockedPixel(X, Y)
        Else
            .low = TheSurface.GetLockedPixel(X, Y)
        End If
        .high = .low
    End With
    
    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    
    If Not (MaskSurface Is Nothing) Then
        MaskSurface.Unlock TmpR
    End If
    
    TheSurface.Unlock TmpR
    Exit Sub
    
ErrorHandler:
    
    Err.Raise Err.Number, , "SetMaskColorFromPixel failed (" & Err.Description & ")"
    
End Sub

Public Sub InitSurfaces()
'On Error GoTo ErrorHandler
Dim i As Long

    Handle_TotalSprites
    
    For i = 0 To TOTAL_SPRITES
        InitDDSurf "Sprites\sprite" & i, DDSD_Sprite(i), DDS_Sprite(i)
    Next
    
    Handle_TotalAnimGfx
    
    For i = 0 To TOTAL_ANIMGFX
        InitDDSurf "Anims\anim" & i, DDSD_Anim(i), DDS_Anim(i)
    Next
    
    InitDDSurf "Items\colormask", DDSD_Mask, DDS_Mask
    
    For i = 0 To MAX_ITEMSETS
        InitDDSurf "Items\item" & i, DDSD_Item(i), DDS_Item(i), DDS_Mask
    Next
    
    Set DDS_Mask = Nothing
    
    InitDDSurf "icons", DDSD_SpellIcon, DDS_SpellIcon
    InitDDSurf "itemhighlight", DDSD_ItemHighlight, DDS_ItemHighlight
    InitDDSurf "spellwaiting", DDSD_SpellWaitingBar, DDS_SpellWaitingBar
    Exit Sub
    
ErrorHandler:
    
    Err.Raise Err.Number, , "InitSurfaces failed (" & Err.Description & ")"
    
End Sub

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(FileName As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7, Optional ByRef MaskSurf As DirectDrawSurface7)
On Error GoTo ErrorHandler

    ' Set path
    FileName = GFX_PATH & FileName & GFX_EXT
    
    ' check if file exists
    If Not FileExist(FileName) Then
        MsgBox "Missing graphics file: " & FileName
        DestroyGame
    End If
    
    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    
    ' select one
    'SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN ' auto determine best
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY ' system memory
    'SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY ' video memory
    
    ' init object
    Set Surf = DD.CreateSurfaceFromFile(App.Path & FileName, SurfDesc)
    
    SetMaskColorFromPixel Surf, 0, 0, MaskSurf
    Exit Sub
    
ErrorHandler:
    
    Err.Raise Err.Number, , "InitDDSurf failed (" & Err.Description & ")"
    
End Sub

Public Sub InitTileSurf(ByVal TileSet As Integer)
On Error GoTo ErrorHandler

    ' Destroy surface if it exist
    If Not DDS_Tile(TileSet) Is Nothing Then
        Set DDS_Tile(TileSet) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Tile(TileSet)), LenB(DDSD_Tile(TileSet))
    End If
    
    InitDDSurf "tiles" & TileSet, DDSD_Tile(TileSet), DDS_Tile(TileSet)
    Exit Sub
    
ErrorHandler:
    
    If Err.Number <> 91 Then Err.Raise Err.Number, , "InitTileSurf failed (" & Err.Description & ")"
    
End Sub

Public Function CheckSurfaces() As Boolean
On Error GoTo ErrorHandler

    ' Check if we need to restore surfaces
    If NeedToRestoreSurfaces Then
        DD.RestoreAllSurfaces
        InitSurfaces
    End If
    
    CheckSurfaces = True
    Exit Function
    
ErrorHandler:
Dim X As Long
Dim Y As Long
Dim i As Long
Dim TilesetLoaded() As Boolean

    ' re-initialize DirectDraw
    DestroyDirectDraw
    InitDirectDraw
    
    ReDim TilesetLoaded(0 To MAX_TILESETS)
    
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            For i = 0 To UBound(Map.Tile(X, Y).Layer)
                If Not TilesetLoaded(Map.Tile(X, Y).LayerSet(i)) Then
                    InitTileSurf Map.Tile(X, Y).LayerSet(i)
                    TilesetLoaded(Map.Tile(X, Y).LayerSet(i)) = True
                End If
            Next
        Next
    Next
    
    InitSurfaces
    
    CheckSurfaces = False
    
End Function

Private Function NeedToRestoreSurfaces() As Boolean
    NeedToRestoreSurfaces = Not (DD.TestCooperativeLevel = DD_OK)
End Function

Public Sub DestroyDirectDraw()
Dim i As Long
On Error Resume Next

    ' Get rid of all of the
    ' DirectDraw objects
    
    For i = 0 To UBound(DDS_Tile)
        Set DDS_Tile(i) = Nothing
    Next
    
    For i = 0 To UBound(DDS_Item)
        Set DDS_Item(i) = Nothing
    Next
    
    For i = 1 To TOTAL_SPRITES
        Set DDS_Sprite(i) = Nothing
    Next
    
    Set DDS_Misc = Nothing
    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    Set DD_Clip = Nothing
    Set DD = Nothing

End Sub

' **************
' ** Drawing  **
' **************

Public Function Engine_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DxVBLib.RECT, dRECT As DxVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True, Optional SkipRefresh As Boolean = False) As Boolean
On Error GoTo ErrorHandle

    If Clear Then picBox.Cls
    
    Surface.BltToDC picBox.hdc, sRECT, dRECT
    If Not SkipRefresh Then picBox.Refresh
    
    Engine_BltToDC = True
    Exit Function
    
ErrorHandle:
    
    ' returns false on error
    Engine_BltToDC = False
    
End Function

Public Sub DrawShopList()
Dim ItemNum As Long
Dim rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT
Dim LoopI As Long

    If frmMainGame.picShop.Visible Then
        frmMainGame.picShopList.Cls
        
        For ItemNum = 1 To MAX_TRADES
            If ShopTrade.TradeItem(ItemNum).GetItem > 0 Then
                
                rec = Get_RECT
                rec_pos = Get_RECT(ShopIconY + ((ShopOffsetY + PIC_Y) * ((ItemNum - 1) \ ShopIconsInRow)), ShopIconY + ((ShopOffsetX + PIC_X) * (((ItemNum - 1) Mod ShopIconsInRow))))
                
                Engine_BltToDC DDS_Item(Item(ShopTrade.TradeItem(ItemNum).GetItem).Pic), rec, rec_pos, frmMainGame.picShopList, False
                
                If ShopTrade.TradeItem(ItemNum).GetValue > 0 Then DrawText frmMainGame.picShopList.hdc, rec_pos.Left, rec_pos.Top + PIC_Y - (FONT_HEIGHT * 2), FormatNumber(ShopTrade.TradeItem(ItemNum).GetValue), FormatNumberColor(ShopTrade.TradeItem(ItemNum).GetValue)
                
            End If
        Next
        
        frmMainGame.picShopList.Refresh
    End If
    
End Sub

Public Sub DrawInventoryList()
Dim ItemNum As Long
Dim rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT
Dim LoopI As Long

    If frmMainGame.picInv.Visible Then
        frmMainGame.picInventoryList.Cls
        
        For ItemNum = 1 To MAX_INV
            If Player(MyIndex).Inv(ItemNum).Num > 0 Then
                
                For LoopI = 1 To Equipment.Equipment_Count - 1
                    If GetPlayerEquipmentSlot(MyIndex, LoopI) = ItemNum Then
                        rec = Get_RECT(, , 34, 34)
                        rec_pos = Get_RECT(ItemIconY + ((ItemOffsetY + PIC_Y) * ((ItemNum - 1) \ ItemsInRow)) - 1, ItemIconY + ((ItemOffsetX + PIC_X) * (((ItemNum - 1) Mod ItemsInRow))) - 1, 34, 34)
                        
                        Engine_BltToDC DDS_ItemHighlight, rec, rec_pos, frmMainGame.picInventoryList, False
                        Exit For
                    End If
                Next
                
                rec = Get_RECT
                rec_pos = Get_RECT(ItemIconY + ((ItemOffsetY + PIC_Y) * ((ItemNum - 1) \ ItemsInRow)), ItemIconY + ((ItemOffsetX + PIC_X) * (((ItemNum - 1) Mod ItemsInRow))))
                
                Engine_BltToDC DDS_Item(Item(Player(MyIndex).Inv(ItemNum).Num).Pic), rec, rec_pos, frmMainGame.picInventoryList, False
                
                If GetPlayerInvItemValue(MyIndex, ItemNum) > 0 And Not ItemIsEquipment(Player(MyIndex).Inv(ItemNum).Num) Then DrawText frmMainGame.picInventoryList.hdc, rec_pos.Left, rec_pos.Top + PIC_Y - (FONT_HEIGHT * 2), FormatNumber(GetPlayerInvItemValue(MyIndex, ItemNum)), FormatNumberColor(GetPlayerInvItemValue(MyIndex, ItemNum))
                
            End If
        Next
        
        frmMainGame.picInventoryList.Refresh
        
    End If
    
End Sub

Public Function FormatNumber2(ByVal Number As Currency) As String

    If InStr(1, Number, ".", vbTextCompare) Then
        FormatNumber2 = Format$(Number, "###,###,###,###,###.####")
    Else
        FormatNumber2 = Format$(Number, "###,###,###,###,###")
    End If
    
End Function

Private Function FormatNumber(ByVal Number As Currency) As String

    If Number < 1000 Then
        FormatNumber = CStr(Number)
        Exit Function
    ElseIf Number >= 1000 And Number < 10000 Then
        FormatNumber = Mid$(CStr(Number), 1, 1) & "." & Mid$(CStr(Number), 2, 1)
    ElseIf Number >= 10000 And Number < 100000 Then
        FormatNumber = Mid$(CStr(Number), 1, 2)
    ElseIf Number >= 100000 And Number < 1000000 Then
        FormatNumber = Mid$(CStr(Number), 1, 3)
    ElseIf Number >= 1000000 And Number < 10000000 Then
        FormatNumber = Mid$(CStr(Number), 1, 1) & "." & Mid$(CStr(Number), 2, 1)
    ElseIf Number >= 10000000 And Number < 100000000 Then
        FormatNumber = Mid$(CStr(Number), 1, 2)
    ElseIf Number >= 100000000 And Number < 1000000000 Then
        FormatNumber = Mid$(CStr(Number), 1, 3)
    ElseIf Number >= 1000000000 And Number < 10000000000# Then
        FormatNumber = Mid$(CStr(Number), 1, 1) & "." & Mid$(CStr(Number), 2, 1)
    ElseIf Number >= 10000000000# And Number < 100000000000# Then
        FormatNumber = Mid$(CStr(Number), 1, 2)
    ElseIf Number >= 100000000000# And Number < 1000000000000# Then
        FormatNumber = Mid$(CStr(Number), 1, 3)
    End If
    
    If Number >= 1000 And Number < 1000000 Then
        FormatNumber = FormatNumber & "k"
    ElseIf Number >= 1000000 And Number < 1000000000 Then
        FormatNumber = FormatNumber & "m"
    ElseIf Number >= 1000000000 Then
        FormatNumber = FormatNumber & "b"
    End If
    
End Function

Private Function FormatNumberColor(ByVal Number As Currency) As Long

    If Number < 1000 Then
        FormatNumberColor = ColorTable(Color.White)
    ElseIf Number >= 1000 And Number < 1000000 Then
        FormatNumberColor = ColorTable(Color.BrightGreen)
    ElseIf Number >= 1000000 And Number < 1000000000 Then
        FormatNumberColor = ColorTable(Color.Yellow)
    ElseIf Number >= 1000000000 Then
        FormatNumberColor = ColorTable(Color.BrightRed)
    End If
    
End Function

Public Function ItemIsEquipment(ByVal ItemNum As Long) As Boolean
    ItemIsEquipment = Item(ItemNum).Type > 0 And Item(ItemNum).Type <= ItemType.Shield_
End Function

Public Sub DrawSpellsWaiting()
Dim rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT
Dim SpellNum As Long

    For SpellNum = 1 To MAX_PLAYER_SPELLS
        If Player(MyIndex).Spell(SpellNum) <> 0 Then
            If Player(MyIndex).CastTimer(SpellNum) > GetTickCountNew Then
                Engine_BltToDC DDS_SpellWaitingBar, _
                               Get_RECT(, , 3), _
                               Get_RECT(, , 3, 32 * ((Player(MyIndex).CastTimer(SpellNum) - GetTickCountNew) / Spell(Player(MyIndex).Spell(SpellNum)).Timer)), _
                               frmMainGame.picSpellWaiting(SpellNum)
            End If
        End If
    Next
    
End Sub

Public Sub DrawSpellList()
Dim SpellNum As Long

    If frmMainGame.picSpells.Visible Then
        frmMainGame.picSpellList.Cls
        
        For SpellNum = 1 To MAX_PLAYER_SPELLS
            DrawSpellIcon SpellNum
        Next
        
        frmMainGame.picSpellList.Refresh
    End If
    
End Sub

Public Sub DrawSpellIcon(ByVal SpellNum As Long)
    If Player(MyIndex).Spell(SpellNum) <> 0 Then
        Engine_BltToDC DDS_SpellIcon, _
                       Get_RECT(Spell(Player(MyIndex).Spell(SpellNum)).Icon * PIC_Y), _
                       Get_RECT(IconY + ((IconOffsetY + PIC_Y) * ((SpellNum - IconsInRow) \ IconsInRow)), _
                       IconX + ((IconOffsetX + PIC_X) * (((SpellNum - IconsInRow) Mod IconsInRow)))), _
                       frmMainGame.picSpellList, False
    End If
End Sub

Public Sub DrawMapTiles()
Dim X As Long
Dim Y As Long

    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
        
            With Map.Tile(X, Y)
            
                If .Layer(Tile_Layer.Ground) > 0 Then
                    DDS_BackBuffer.BltFast MultiplyPicX(X), MultiplyPicX(Y), DDS_Tile(.LayerSet(Tile_Layer.Ground)), _
                                           Get_RECT(MultiplyPicX((.Layer(Tile_Layer.Ground) \ TILESHEET_WIDTH(.LayerSet(Tile_Layer.Ground)))), (MultiplyPicX(ModularTable(.Layer(Tile_Layer.Ground), TILESHEET_WIDTH(.LayerSet(Tile_Layer.Ground)))))), _
                                           DDBLTFAST_WAIT
                End If
                
                If MapAnim = 0 Or .Layer(Tile_Layer.Anim) <= 0 Then
                    If .Layer(Tile_Layer.Mask) > 0 Then
                        If TempTile(X, Y).DoorOpen = NO Then
                            DDS_BackBuffer.BltFast MultiplyPicX(X), MultiplyPicX(Y), DDS_Tile(.LayerSet(Tile_Layer.Mask)), _
                                                   Get_RECT(MultiplyPicX((.Layer(Tile_Layer.Mask) \ TILESHEET_WIDTH(.LayerSet(Tile_Layer.Mask)))), (MultiplyPicX(ModularTable(.Layer(Tile_Layer.Mask), TILESHEET_WIDTH(.LayerSet(Tile_Layer.Ground)))))), _
                                                   DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                        End If
                    End If
                Else
                    ' Is there an animation tile to draw?
                    If .Layer(Tile_Layer.Anim) > 0 Then
                        DDS_BackBuffer.BltFast MultiplyPicX(X), MultiplyPicX(Y), DDS_Tile(.LayerSet(Tile_Layer.Anim)), _
                                               Get_RECT(MultiplyPicX((.Layer(Tile_Layer.Anim) \ TILESHEET_WIDTH(.LayerSet(Tile_Layer.Anim)))), (MultiplyPicX(ModularTable(.Layer(Tile_Layer.Anim), TILESHEET_WIDTH(.LayerSet(Tile_Layer.Ground)))))), _
                                               DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                End If
                
            End With
            
        Next
    Next
    
End Sub

Public Sub DrawMapFringeTiles()
Dim X As Long
Dim Y As Long

    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
        
            With Map.Tile(X, Y)
                If .Layer(Tile_Layer.Fringe) > 0 Then
                    DDS_BackBuffer.BltFast MultiplyPicX(X), MultiplyPicX(Y), DDS_Tile(.LayerSet(Tile_Layer.Fringe)), _
                                           Get_RECT(MultiplyPicX((.Layer(Tile_Layer.Fringe) \ TILESHEET_WIDTH(.LayerSet(Tile_Layer.Fringe)))), (MultiplyPicX(ModularTable(.Layer(Tile_Layer.Fringe), TILESHEET_WIDTH(.LayerSet(Tile_Layer.Fringe)))))), _
                                           DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            End With
            
        Next
    Next
    
End Sub

Public Sub DrawItems()
Dim ItemNum As Long

    For ItemNum = 1 To MAX_MAP_ITEMS
    
        If MapItem(ItemNum).Num > 0 Then
            DDS_BackBuffer.BltFast MultiplyPicX(MapItem(ItemNum).X), MultiplyPicX(MapItem(ItemNum).Y), DDS_Item(Item(MapItem(ItemNum).Num).Pic), _
                                   Get_RECT, _
                                   DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            
            If MapItem(ItemNum).Anim > 0 Then
                If frmAdmin.chkHideAnimations <> 1 Then
                    Dim X As Long
                    Dim Y As Long
                    
                    If MapItem(ItemNum).AnimTimer < GetTickCountNew Then
                        If MapItem(ItemNum).AnimFrame > (DDSD_Anim(Animation(MapItem(ItemNum).Anim).Pic).lWidth \ Animation(MapItem(ItemNum).Anim).Width) - 1 Then MapItem(ItemNum).AnimFrame = -1
                        MapItem(ItemNum).AnimFrame = MapItem(ItemNum).AnimFrame + 1
                        MapItem(ItemNum).AnimTimer = GetTickCountNew + Animation(MapItem(ItemNum).Anim).Delay
                    End If
                    
                    X = MultiplyPicX(MapItem(ItemNum).X)
                    Y = MultiplyPicX(MapItem(ItemNum).Y)
                    
                    If Animation(MapItem(ItemNum).Anim).Width <> PIC_X Then
                        X = X - ((Animation(MapItem(ItemNum).Anim).Width - PIC_X) \ 2)
                    End If
                    
                    If Animation(MapItem(ItemNum).Anim).Height <> PIC_Y Then
                        Y = Y - ((Animation(MapItem(ItemNum).Anim).Height - PIC_Y) \ 2)
                    End If
                    
                    DrawAnim Get_RECT(, MapItem(ItemNum).AnimFrame * Animation(MapItem(ItemNum).Anim).Width, Animation(MapItem(ItemNum).Anim).Width, Animation(MapItem(ItemNum).Anim).Height), _
                             X, Y, MapItem(ItemNum).Anim
                    
                End If
            End If
            
        End If
        
    Next
    
End Sub

Public Sub DrawAnimations()
Dim i As Long
Dim X As Long
Dim Y As Long

    For i = 1 To UBound(Animations)
    
        If Animations(i).Active Then
            If Animations(i).Picture < TOTAL_ANIMGFX Then
                
                If Animations(i).Timer < GetTickCountNew Then
                    If Animations(i).Frame > (DDSD_Anim(i).lWidth \ Animation(Animations(i).Key).Width) - 1 Then
                        ZeroMemory ByVal VarPtr(Animations(i)), LenB(Animations(i))
                        GoTo Skip
                    End If
                    Animations(i).Frame = Animations(i).Frame + 1
                    Animations(i).Timer = GetTickCountNew + Animations(i).DelayTime
                End If
                
                X = Animations(i).X
                Y = Animations(i).Y
                
                If Animation(Animations(i).Key).Width <> PIC_X Then
                    X = X - ((Animation(Animations(i).Key).Width - PIC_X) \ 2)
                End If
                
                If Animation(Animations(i).Key).Height <> PIC_Y Then
                    Y = Y - ((Animation(Animations(i).Key).Height - PIC_Y) \ 2)
                End If
                
                DrawAnim Get_RECT(, Animations(i).Frame * Animation(Animations(i).Key).Width, Animation(Animations(i).Key).Width, Animation(Animations(i).Key).Height), _
                         X, Y, Animations(i).Key
                
            End If
        End If
Skip:
    Next
    
End Sub

Public Sub Draw_WithY()
Dim Y As Long

    ' Handle the walking and attacking animation for the NPCs
    Handle_NpcAnims
    
    ' Handle the walking and attacking animation for the players
    Handle_PlayerAnims
    
    For Y = 0 To MAX_MAPY
        DrawPlayer Y
        DrawNpc Y
    Next
    
End Sub

Public Sub DrawAnim(ByRef rec As DxVBLib.RECT, ByVal X As Long, ByVal Y As Long, ByVal AnimNum As Long)

    If Y < 0 Then
        rec.Top = rec.Top - Y
        Y = 0
    End If
    
    If X < 0 Then
        rec.Left = rec.Left - X
        X = 0
    End If
    
    If X + Animation(AnimNum).Width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Animation(AnimNum).Width - DDSD_BackBuffer.lWidth)
    End If
    
    If Y + Animation(AnimNum).Height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + Animation(AnimNum).Height - DDSD_BackBuffer.lHeight)
    End If
    
    DDS_BackBuffer.BltFast X, Y, DDS_Anim(Animation(AnimNum).Pic), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
End Sub

Public Sub DrawSprite(ByRef rec As DxVBLib.RECT, ByVal X As Long, ByVal Y As Long, ByVal SpriteNum As Long)

    If Y < 0 Then
        rec.Top = rec.Top - Y
        Y = 0
    End If
    
    If X < 0 Then
        rec.Left = rec.Left - X
        X = 0
    End If
    
    If X + Sprite_Size(SpriteNum).SizeX > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + Sprite_Size(SpriteNum).SizeX - DDSD_BackBuffer.lWidth)
    End If
    
    DDS_BackBuffer.BltFast X, Y, DDS_Sprite(SpriteNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
End Sub

Public Sub DrawPlayer(ByVal YPos As Long)
Dim Anim As Byte
Dim X As Long
Dim Y As Long
Dim Index As Long

    For Index = 1 To MAX_PLAYERS
        If IsPlaying(Index) Then
            If GetPlayerMap(Index) = GetPlayerMap(MyIndex) Then
                If GetPlayerY(Index) = YPos Then
                
                    If GetPlayerSprite(Index) < 0 Or GetPlayerSprite(Index) > TOTAL_SPRITES Then GoTo Skipper
                    
                    Anim = GameConfig.StandFrame
                    
                    With Player(Index)
                        If .Moving Then
                            If .WalkAnim > UBound(GameConfig.WalkFrame) Then .WalkAnim = UBound(GameConfig.WalkFrame)
                            Anim = GameConfig.WalkFrame(.WalkAnim)
                        End If
                        
                        If .Attacking Then
                            If .WalkAnim > UBound(GameConfig.AttackFrame) Then .WalkAnim = UBound(GameConfig.AttackFrame)
                            Anim = GameConfig.AttackFrame(.WalkAnim)
                        End If
                        
                        If .AttackTimer + 1000 < GetTickCountNew Then
                            .Attacking = 0
                            .AttackTimer = 0
                        End If
                        
                        X = GetPlayerX(Index) * PIC_X + .XOffset
                        Y = GetPlayerY(Index) * PIC_Y + .YOffset - Sprite_Offset
                    End With
                    
                    If Sprite_Size(GetPlayerSprite(Index)).SizeX <> PIC_X Then
                        X = X - ((Sprite_Size(GetPlayerSprite(Index)).SizeX - PIC_X) \ 2)
                    End If
                    
                    If Sprite_Size(GetPlayerSprite(Index)).SizeY <> PIC_Y Then
                        Y = Y - (Sprite_Size(GetPlayerSprite(Index)).SizeY - PIC_Y)
                    End If
                    
                    DrawSprite Get_RECT(, (Direction_Anim(GetPlayerDir(Index)) * (Total_SpriteFrames) + Anim) * Sprite_Size(GetPlayerSprite(Index)).SizeX, Sprite_Size(GetPlayerSprite(Index)).SizeX, Sprite_Size(GetPlayerSprite(Index)).SizeY), _
                               X, Y, GetPlayerSprite(Index)
                    
                End If
Skipper:
            End If
        End If
    Next
    
End Sub

Public Sub DrawPlayerBars()
Dim Anim As Byte
Dim X As Long
Dim Y As Long
Dim Index As Long

    For Index = 1 To MAX_PLAYERS
        If IsPlaying(Index) Then
            If GetPlayerMap(Index) = GetPlayerMap(MyIndex) Then
                
                If GetPlayerSprite(Index) < 0 Or GetPlayerSprite(Index) > TOTAL_SPRITES Then GoTo Skipper
                If GetPlayerVital(Index, HP) >= GetPlayerMaxVital(Index, HP) Then GoTo Skipper
                
                With Player(Index)
                    X = MultiplyPicX(GetPlayerX(Index)) + .XOffset
                    Y = MultiplyPicX(GetPlayerY(Index)) + .YOffset - Sprite_Offset
                End With
                
                If Sprite_Size(GetPlayerSprite(Index)).SizeX <> PIC_X Then
                    X = X - ((Sprite_Size(GetPlayerSprite(Index)).SizeX - PIC_X) \ 2)
                End If
                
                If Sprite_Size(GetPlayerSprite(Index)).SizeY <> PIC_Y Then
                    Y = Y - (Sprite_Size(GetPlayerSprite(Index)).SizeY - PIC_Y)
                End If
                
                Y = Y + Sprite_Size(GetPlayerSprite(Index)).SizeY + 2
                
                DDS_BackBuffer.SetFillColor RGB(0, 0, 0)
                DDS_BackBuffer.DrawBox X, Y, X + Sprite_Size(GetPlayerSprite(Index)).SizeX, Y + 4
                
                DDS_BackBuffer.SetFillColor RGB(255, 0, 0)
                DDS_BackBuffer.DrawBox X, Y, X + Sprite_Size(GetPlayerSprite(Index)).SizeX, Y + 4
                
                DDS_BackBuffer.SetFillColor RGB(0, 255, 0)
                DDS_BackBuffer.DrawBox X, Y, X + (Sprite_Size(GetPlayerSprite(Index)).SizeX * (GetPlayerVital(Index, HP) / GetPlayerMaxVital(Index, HP))), Y + 4
Skipper:
            End If
        End If
    Next
    
End Sub

Public Sub DrawNpc(ByVal YPos As Long)
Dim MapNpcNum As Long
Dim Anim As Byte
Dim X As Long
Dim Y As Long

    For MapNpcNum = 1 To UBound(MapNpc)
        If GetMapNpcY(MapNpcNum) = YPos Then
            
            If MapNpc(MapNpcNum).Num < 1 Or GetMapNpcSprite(MapNpcNum) < 0 Or GetMapNpcSprite(MapNpcNum) > TOTAL_SPRITES Then GoTo Skipper
            
            Anim = GameConfig.StandFrame
            
            With MapNpc(MapNpcNum)
                If .Attacking Or .Moving Then Anim = .WalkAnim
                
                If .AttackTimer + 1000 < GetTickCountNew Then
                    .Attacking = 0
                    .AttackTimer = 0
                End If
                
                X = MultiplyPicX(GetMapNpcX(MapNpcNum)) + .XOffset
                Y = MultiplyPicX(GetMapNpcY(MapNpcNum)) + .YOffset - Sprite_Offset
            End With
            
            If Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeX <> PIC_X Then
                X = X - ((Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeX - PIC_X) \ 2)
            End If
            
            If Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeY <> PIC_Y Then
                Y = Y - (Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeY - PIC_Y)
            End If
            
            DrawSprite Get_RECT(, (Direction_Anim(GetMapNpcDir(MapNpcNum)) * (Total_SpriteFrames) + Anim) * Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeX, Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeX, Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeY), _
                       X, Y, GetMapNpcSprite(MapNpcNum)
            
        End If
Skipper:
    Next
        
End Sub

Public Sub DrawNpcBars()
Dim MapNpcNum As Long
Dim Anim As Byte
Dim X As Long
Dim Y As Long

    For MapNpcNum = 1 To UBound(MapNpc)
        If MapNpc(MapNpcNum).Num < 1 Or GetMapNpcSprite(MapNpcNum) < 0 Or GetMapNpcSprite(MapNpcNum) > TOTAL_SPRITES Then GoTo Skipper
        If MapNpc(MapNpcNum).Vital(Vitals.HP) < 1 Or MapNpc(MapNpcNum).Vital(Vitals.HP) >= Npc(MapNpc(MapNpcNum).Num).HP Then GoTo Skipper
        
        With MapNpc(MapNpcNum)
            X = MultiplyPicX(GetMapNpcX(MapNpcNum)) + .XOffset
            Y = MultiplyPicX(GetMapNpcY(MapNpcNum)) + .YOffset - Sprite_Offset
        End With
        
        If Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeX <> PIC_X Then
            X = X - ((Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeX - PIC_X) \ 2)
        End If
        
        If Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeY <> PIC_Y Then
            Y = Y - (Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeY - PIC_Y)
        End If
        
        Y = Y + Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeY + 2
        
        DDS_BackBuffer.SetFillColor RGB(0, 0, 0)
        DDS_BackBuffer.DrawBox X, Y, X + Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeX, Y + 4
        
        DDS_BackBuffer.SetFillColor RGB(255, 0, 0)
        DDS_BackBuffer.DrawBox X, Y, X + Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeX, Y + 4
        
        DDS_BackBuffer.SetFillColor RGB(0, 255, 0)
        DDS_BackBuffer.DrawBox X, Y, X + (Sprite_Size(GetMapNpcSprite(MapNpcNum)).SizeX * (MapNpc(MapNpcNum).Vital(Vitals.HP) / Npc(MapNpc(MapNpcNum).Num).HP)), Y + 4
Skipper:
    Next
    
End Sub

' ******************
' ** Game Editors **
' ******************

Public Sub BltMapEditor()
Dim Height As Long
Dim Width As Long
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    Height = DDSD_Tile(frmMainGame.scrlTileSet).lHeight
    Width = DDSD_Tile(frmMainGame.scrlTileSet).lWidth
    
    dRECT = Get_RECT(, , Width, Height)
    
    frmMainGame.picBackSelect.Height = Height
    frmMainGame.picBackSelect.Width = Width
    
    Engine_BltToDC DDS_Tile(frmMainGame.scrlTileSet), sRECT, dRECT, frmMainGame.picBackSelect
    
End Sub

Public Sub BltTileOutline()
    DDS_BackBuffer.BltFast MultiplyPicX(CurX), MultiplyPicX(CurY), DDS_Misc, Get_RECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub

'Public Sub MapItemEditorBltItem()
'Dim sRECT As DxVBLib.RECT

'    If LenB(Trim$(Item(frmMapItem.scrlItem.Value).Name)) < 1 Then
'        frmMapItem.picPreview.Cls
'        Exit Sub
'    End If
    
'    sRECT = Get_RECT(Item(frmMapItem.scrlItem.Value).Pic * PIC_Y)
'    Engine_BltToDC DDS_Item, sRECT, Get_RECT, frmMapItem.picPreview
    
'End Sub

'Public Sub KeyItemEditorBltItem()
'Dim sRECT As DxVBLib.RECT

'    If LenB(Trim$(Item(frmMapKey.scrlItem.Value).Name)) < 1 Then
'        frmMapKey.picPreview.Cls
'        Exit Sub
'    End If
    
'    sRECT = Get_RECT(Item(frmMapKey.scrlItem.Value).Pic * PIC_Y)
'    Engine_BltToDC DDS_Item, sRECT, Get_RECT, frmMapKey.picPreview
    
'End Sub

Public Sub ItemEditorBltItem()

    Engine_BltToDC DDS_Item(frmItemEditor.scrlPic.Value), Get_RECT, Get_RECT, frmItemEditor.picPic
    
End Sub

Public Sub NpcEditorBltSprite()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    frmNpcEditor.picSprite.Width = Sprite_Size(frmNpcEditor.scrlSprite.Value).SizeX * Screen.TwipsPerPixelX
    frmNpcEditor.picSprite.Height = Sprite_Size(frmNpcEditor.scrlSprite.Value).SizeY * Screen.TwipsPerPixelY
    
    sRECT = Get_RECT(, (Direction_Anim(E_Direction.Down_) * Total_SpriteFrames) * Sprite_Size(frmNpcEditor.scrlSprite.Value).SizeX, Sprite_Size(frmNpcEditor.scrlSprite.Value).SizeX, Sprite_Size(frmNpcEditor.scrlSprite.Value).SizeY)
    dRECT = Get_RECT(, , Sprite_Size(frmNpcEditor.scrlSprite.Value).SizeX, Sprite_Size(frmNpcEditor.scrlSprite.Value).SizeY)
    
    Engine_BltToDC DDS_Sprite(frmNpcEditor.scrlSprite.Value), sRECT, dRECT, frmNpcEditor.picSprite
    
End Sub

Public Sub AnimEditorDrawPic()
Dim sRECT As DxVBLib.RECT
Dim dRECT As DxVBLib.RECT

    frmAnimEditor.picPic.Width = Val(frmAnimEditor.txtSizeX) * Screen.TwipsPerPixelX
    frmAnimEditor.picPic.Height = Val(frmAnimEditor.txtSizeY) * Screen.TwipsPerPixelY
    
    sRECT = Get_RECT(, Val(frmAnimEditor.txtSizeX) * AnimEditorAnim, Val(frmAnimEditor.txtSizeX), Val(frmAnimEditor.txtSizeY))
    dRECT = Get_RECT(, , Val(frmAnimEditor.txtSizeX), Val(frmAnimEditor.txtSizeY))
    
    Engine_BltToDC DDS_Anim(frmAnimEditor.scrlPic.Value), sRECT, dRECT, frmAnimEditor.picPic
    
End Sub

Public Sub SpellEditorBltIcon()

    Engine_BltToDC DDS_SpellIcon, _
                   Get_RECT(frmSpellEditor.scrlIcon.Value * SpellIconHeight), _
                   Get_RECT, _
                   frmSpellEditor.picIcon
    
End Sub

Public Sub DrawSelectedCharacter(ByVal Index As Long, ByVal Frame As Long)

    If CharIsThere(Index) Then
        If frmChars.picChar(Index).Width <> Sprite_Size(Char_Sprite(Index)).SizeX Then frmChars.picChar(Index).Width = Sprite_Size(Char_Sprite(Index)).SizeX
        If frmChars.picChar(Index).Height <> Sprite_Size(Char_Sprite(Index)).SizeY Then frmChars.picChar(Index).Height = Sprite_Size(Char_Sprite(Index)).SizeY
        
        Engine_BltToDC DDS_Sprite(Char_Sprite(Index)), _
                       Get_RECT(0, (Direction_Anim(E_Direction.Down_) * (Total_SpriteFrames) + Frame) * Sprite_Size(Char_Sprite(Index)).SizeX, Sprite_Size(Char_Sprite(Index)).SizeX, Sprite_Size(Char_Sprite(Index)).SizeY), _
                       Get_RECT(0, 0, Sprite_Size(Char_Sprite(Index)).SizeX, Sprite_Size(Char_Sprite(Index)).SizeY), _
                       frmChars.picChar(Index)
    End If
    
End Sub

Public Function Get_RECT(Optional ByVal TopVal As Long = 0, Optional ByVal LeftVal As Long = 0, Optional ByVal Width As Long = PIC_X, Optional ByVal Height As Long = PIC_Y) As DxVBLib.RECT
    With Get_RECT
        .Top = TopVal
        .Bottom = .Top + Height
        .Left = LeftVal
        .Right = .Left + Width
    End With
End Function

Private Sub DrawMapAttributes()
Dim X As Long
Dim Y As Long

    If frmMainGame.optAttribs.Value Or SettingSpawn Then
        For X = 0 To MAX_MAPX
            For Y = 0 To MAX_MAPY
                With Map.Tile(X, Y)
                    Select Case .Type
                    
                        Case Tile_Type.Blocked_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "B", ColorTable(Color.BrightRed)
                            
                        Case Tile_Type.Warp_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "W", ColorTable(Color.White)
                            
                        Case Tile_Type.Item_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "I", ColorTable(Color.White)
                            
                        Case Tile_Type.NpcAvoid_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "N", ColorTable(Color.White)
                            
                        Case Tile_Type.Key_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "K", ColorTable(Color.White)
                            
                        Case Tile_Type.KeyOpen_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "O", ColorTable(Color.White)
                            
                        Case Tile_Type.Shop_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "SH", ColorTable(Color.White)
                            
                        Case Tile_Type.Sign_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "SI", ColorTable(Color.White)
                            
                        Case Tile_Type.Guild_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "G", ColorTable(Color.White)
                            
                        Case Tile_Type.Heal_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "H", ColorTable(Color.White)
                            
                        Case Tile_Type.Damage_
                            DrawText TexthDC, ((MultiplyPicX(X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((Y * PIC_Y) - FONT_HEIGHT) + (PIC_Y \ 2), "DOT", ColorTable(Color.White)
                            
                    End Select
                End With
            Next
        Next
        For X = 1 To UBound(MapSpawn.Npc)
            If MapSpawn.Npc(X).X <> -1 Then
                DrawText TexthDC, ((MultiplyPicX(MapSpawn.Npc(X).X)) - (FONT_WIDTH \ 2)) + (PIC_X \ 2), ((MultiplyPicX(MapSpawn.Npc(X).Y)) - FONT_HEIGHT) + (PIC_Y \ 2), "N" & X, ColorTable(Color.White)
            End If
        Next
    End If
    
End Sub

Public Sub Render_Graphics()
Dim rec As DxVBLib.RECT
Dim rec_pos As DxVBLib.RECT

    If frmMainGame.WindowState = vbMinimized Then Exit Sub
    
    If Not CheckSurfaces Then Exit Sub
    
    If Not GettingMap Then
    
        ' Clear the back buffer before we draw on it
        ' or we can say this is "refreshing" the screen
        DDS_BackBuffer.BltColorFill rec, 0
        
        ' Draw the bottom layers
        DrawMapTiles
        
        ' Draw the items
        DrawItems
        
        If Not (GetPlayerAccess(MyIndex) > 0 And frmAdmin.chkHideSprites.Value = 1) Then
            ' Draw the NPC and player sprites
            ' based on their Y position
            Draw_WithY
        End If
        
        DrawPlayerBars
        DrawNpcBars
        
        ' Draw the fringe layer
        DrawMapFringeTiles
        
        If Not (GetPlayerAccess(MyIndex) > 0 And frmAdmin.chkHideAnimations.Value = 1) Then
            ' Draw the spell animations
            DrawAnimations
        End If
        
        If InEditor Then BltTileOutline
        
        If (GetPlayerAccess(MyIndex) > 0 And frmAdmin.chkHideAll.Value = 1) Then GoTo SkipDrawText:
        
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DDS_BackBuffer.GetDC
        
        If InEditor Then DrawMapAttributes
        
        If PingEnabled Then DrawText TexthDC, 2, frmMainGame.picScreen.Height - 17, "Ping: " & CurPing, ColorTable(Color.Yellow)
        
        ' draw FPS
        If BFPS Then DrawText TexthDC, ((frmMainGame.picScreen.Width) - (Len("FPS: " & GameFPS)) * FONT_WIDTH) - 1, 0, "FPS: " & GameFPS, ColorTable(Color.Yellow)
        
        ' draw cursor, player X and Y locations
        If BLoc Then
            DrawText TexthDC, 1, 0, "Cur X: " & CurX & " Y: " & CurY, ColorTable(Color.Yellow)
            DrawText TexthDC, 1, 15, "Loc X: " & GetPlayerX(MyIndex) & " Y: " & GetPlayerY(MyIndex), ColorTable(Color.Yellow)
            DrawText TexthDC, 1, 30, "Map #" & GetPlayerMap(MyIndex), ColorTable(Color.Yellow)
        End If
        
        ' drawing the player guild names
        If ShowPNames Then DrawPlayerGuildNames
        
        ' draw npc and player names
        If ShowNNames Then DrawNpcNames
        If ShowPNames Then DrawPlayerNames
        
        If Not (GetPlayerAccess(MyIndex) > 0 And frmAdmin.chkHideMap.Value = 1) Then
            ' Draw map name
            DrawText TexthDC, DrawMapNameX, DrawMapNameY, Map.Name, DrawMapNameColor
        End If
    Else
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DDS_BackBuffer.GetDC
        
        ' Check if we are getting a map, and if we are tell them so
        DrawText TexthDC, 50, 50, "Receiving Map...", ColorTable(Color.BrightCyan)
        
    End If
    
    ' Release DC
    DDS_BackBuffer.ReleaseDC TexthDC
    
SkipDrawText:
    
    ' Get the rect to blit to
    DX7.GetWindowRect frmMainGame.picScreen.hWnd, rec_pos
    
    ' Blit the backbuffer
    DDS_Primary.Blt rec_pos, DDS_BackBuffer, rec, DDBLT_WAIT
    
End Sub

