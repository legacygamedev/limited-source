Attribute VB_Name = "modEditors"
Option Explicit

Public Sub MapEditorInit()
Dim rec As RECT
Dim rec_pos As RECT

    SaveMap = Map
    InEditor = True
    
    frmMapEditor.Visible = True
    frmMapEditor.Top = frmMainGame.Top
    frmMapEditor.Left = frmMainGame.Left + frmMainGame.Width

    frmMapEditor.scrlPicture.Max = Int(DDSD_Tile.lHeight / PIC_Y) - 7
    
    ' Set the max for NPC info
    frmMapEditor.scrlMob.Max = MAX_MOBS
    
    ' Set the width of the mapeditor
    frmMapEditor.picBack.Width = DDSD_Tile.lWidth

    ' Set the width of the form
    frmMapEditor.scrlPicture.Left = frmMapEditor.picBack.Width + frmMapEditor.scrlPicture.Width
    frmMapEditor.Width = (frmMapEditor.scrlPicture.Left + frmMapEditor.scrlPicture.Width) * 16
    
    With rec
        .Top = 0
        .Bottom = frmMapEditor.picBack.Height
        .Left = 0
        .Right = frmMapEditor.picBack.Width
    End With
   
    With rec_pos
        If frmMapEditor.scrlPicture.Value = 0 Then
            .Top = 0
        Else
            .Top = (frmMapEditor.scrlPicture.Value * PIC_Y) * 1
        End If
        .Left = 0
        .Bottom = .Top + (frmMapEditor.picBack.Height)
        .Right = frmMapEditor.picBack.Width
    End With
    
    DD_TileSurf.BltToDC frmMapEditor.picBack.hdc, rec_pos, rec
    
    frmMapEditor.picBack.Refresh
    
    'frmMainGame.Width = 15435
End Sub

Public Sub MapEditorMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Dim Tile As Long
Dim X2 As Long
Dim Y2 As Long

    If InEditor Then
        X = TileView.Left + ((X + Camera.Left) \ PIC_X)
        Y = TileView.Top + ((Y + Camera.Top) \ PIC_Y)
        
        If Not IsValidMapPoint(X, Y) Then Exit Sub
        
        If Button = 1 Then
            If frmMapEditor.optLayers.Value = True Then
                If frmMapEditor.chkRandomTile Then
                    Tile = RandomTile(Int(Rnd * 4))
                    With Map.Tile(X, Y)
                        If frmMapEditor.optGround.Value = True Then .Ground = Tile
                        If frmMapEditor.optMask.Value = True Then .Mask = Tile
                        If frmMapEditor.optAnim.Value = True Then .Anim = Tile
                        If frmMapEditor.optMask2.Value = True Then .Mask2 = Tile
                        If frmMapEditor.optM2anim.Value = True Then .M2Anim = Tile
                        If frmMapEditor.optFringe.Value = True Then .Fringe = Tile
                        If frmMapEditor.optFAnim.Value = True Then .FAnim = Tile
                        If frmMapEditor.optFringe2.Value = True Then .Fringe2 = Tile
                        If frmMapEditor.optF2anim.Value = True Then .F2Anim = Tile
                    End With
                Else
                    For X2 = 0 To (frmMapEditor.shpSelection.Width / PIC_X) - 1
                        For Y2 = 0 To (frmMapEditor.shpSelection.Height / PIC_Y) - 1
                            If IsValidMapPoint(X + X2, Y + Y2) Then
                                Tile = (EditorTileY + Y2) * TILE_WIDTH + (EditorTileX + X2)
                                With Map.Tile(X + X2, Y + Y2)
                                    If frmMapEditor.optGround.Value = True Then .Ground = Tile
                                    If frmMapEditor.optMask.Value = True Then .Mask = Tile
                                    If frmMapEditor.optAnim.Value = True Then .Anim = Tile
                                    If frmMapEditor.optMask2.Value = True Then .Mask2 = Tile
                                    If frmMapEditor.optM2anim.Value = True Then .M2Anim = Tile
                                    If frmMapEditor.optFringe.Value = True Then .Fringe = Tile
                                    If frmMapEditor.optFAnim.Value = True Then .FAnim = Tile
                                    If frmMapEditor.optFringe2.Value = True Then .Fringe2 = Tile
                                    If frmMapEditor.optF2anim.Value = True Then .F2Anim = Tile
                                End With
                            End If
                        Next
                    Next
                End If
            ElseIf frmMapEditor.optAttribs.Value Then
                With Map.Tile(X, Y)
                    If frmMapEditor.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                    If frmMapEditor.optWarp.Value = True Then
                        .Type = TILE_TYPE_WARP
                        .Data1 = EditorWarpMap
                        .Data2 = EditorWarpX
                        .Data3 = EditorWarpY
                    End If
                    If frmMapEditor.optItem.Value = True Then
                        .Type = TILE_TYPE_ITEM
                        .Data1 = ItemEditorNum
                        .Data2 = ItemEditorValue
                        .Data3 = ItemEditorBlocked
                    End If
                    If frmMapEditor.optNpcAvoid.Value = True Then
                        .Type = TILE_TYPE_NPCAVOID
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                    End If
                    If frmMapEditor.optKey.Value = True Then
                        .Type = TILE_TYPE_KEY
                        .Data1 = KeyEditorNum
                        .Data2 = KeyEditorTake
                        .Data3 = 0
                    End If
                    If frmMapEditor.optKeyOpen.Value = True Then
                        .Type = TILE_TYPE_KEYOPEN
                        .Data1 = KeyOpenEditorX
                        .Data2 = KeyOpenEditorY
                        .Data3 = KeyOpenPressure
                    End If
                End With
            ElseIf frmMapEditor.optNpcs.Value Then
                With Map.Tile(X, Y)
                    ' First make sure that there are npcs in this mob group
                    If Map.Mobs(frmMapEditor.scrlMob.Value).NpcCount > 0 Then
                        .Type = TILE_TYPE_MOBSPAWN
                        .Data1 = frmMapEditor.scrlMob.Value         ' What mob num to spawn on this tile
                        .Data2 = frmMapEditor.cmbDir.ListIndex - 1  ' What direcition
                        .Data3 = 0
                    End If
                End With
            End If
        End If
        
        If Button = 2 Then
            If frmMapEditor.optLayers.Value = True Then
                With Map.Tile(X, Y)
                    If frmMapEditor.optGround.Value = True Then .Ground = 0
                    If frmMapEditor.optMask.Value = True Then .Mask = 0
                    If frmMapEditor.optAnim.Value = True Then .Anim = 0
                    If frmMapEditor.optMask2.Value = True Then .Mask2 = 0
                    If frmMapEditor.optM2anim.Value = True Then .M2Anim = 0
                    If frmMapEditor.optFringe.Value = True Then .Fringe = 0
                    If frmMapEditor.optFAnim.Value = True Then .FAnim = 0
                    If frmMapEditor.optFringe2.Value = True Then .Fringe2 = 0
                    If frmMapEditor.optF2anim.Value = True Then .F2Anim = 0
                End With
            Else
                With Map.Tile(X, Y)
                    .Type = 0
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End With
            End If
        End If
    End If
End Sub

Public Sub MapEditorChooseTile(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rec As RECT
Dim rec_pos As RECT
Dim i As Long
    
    If Button = vbLeftButton Then
        EditorTileX = Int(X / PIC_X)
        EditorTileY = Int(Y / PIC_Y) + frmMapEditor.scrlPicture.Value
            
        If (Shift And vbShiftMask) = False Then
            frmMapEditor.shpSelection.Left = Int(X / PIC_X) * PIC_X
            frmMapEditor.shpSelection.Top = Int(Y / PIC_Y) * PIC_Y
            frmMapEditor.shpSelection.Width = PIC_X
            frmMapEditor.shpSelection.Height = PIC_Y
        Else
            frmMapEditor.shpSelection.Width = Clamp((EditorTileX + 1 - (frmMapEditor.shpSelection.Left / PIC_X)) * PIC_X, PIC_X, frmMapEditor.picBack.Width - frmMapEditor.shpSelection.Left)
            frmMapEditor.shpSelection.Height = Clamp(((EditorTileY + 1 - frmMapEditor.scrlPicture.Value) - (frmMapEditor.shpSelection.Top / PIC_Y)) * PIC_Y, PIC_Y, frmMapEditor.picBack.Height - frmMapEditor.shpSelection.Top)
            EditorTileX = Int((frmMapEditor.shpSelection.Left + PIC_X) / PIC_X) - 1
            EditorTileY = Int((frmMapEditor.shpSelection.Top + PIC_Y) / PIC_Y) - 1 + frmMapEditor.scrlPicture.Value
        End If
            
        ' Random tile section
        If frmMapEditor.chkRandomTile Then RandomTile(RandomTileSelected) = EditorTileY * TILE_WIDTH + EditorTileX
           
        With rec
            .Top = EditorTileY * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = EditorTileX * PIC_X
            .Right = .Left + PIC_X
        End With
        
        With rec_pos
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
        
        For i = 0 To 3
            DD_TileSurf.BltToDC frmMapEditor.picRandomTile(i).hdc, LookUpTileRec(RandomTile(i)), rec_pos
            frmMapEditor.picRandomTile(i).Refresh
        Next
    End If
End Sub

Public Sub MapEditorTileScroll()
Dim rec As RECT
Dim rec_pos As RECT

    With rec
        .Top = 0
        .Bottom = frmMapEditor.picBack.Height
        .Left = 0
        .Right = frmMapEditor.picBack.Width
    End With

    With rec_pos
        .Top = frmMapEditor.scrlPicture.Value * PIC_Y
        .Left = 0
        .Bottom = .Top + frmMapEditor.picBack.Height
        .Right = frmMapEditor.picBack.Width
    End With

    DD_TileSurf.BltToDC frmMapEditor.picBack.hdc, rec_pos, rec
    frmMapEditor.picBack.Refresh
End Sub

Public Sub MapEditorSend()
    SendMap
    MapEditorCancel
End Sub

Public Sub MapEditorCancel()
    Map = SaveMap
    
    UpdateMapNpcCount
    
    InEditor = False
    
    frmMapEditor.Visible = False
End Sub

Public Sub MapEditorFillLayer()
Dim YesNo As Long
Dim X As Long
Dim Y As Long
Dim Tile As Long

    YesNo = MsgBox("Are you sure you wish to fill this layer?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Tile = EditorTileY * TILE_WIDTH + EditorTileX
                With Map.Tile(X, Y)
                    If frmMapEditor.optGround.Value = True Then .Ground = Tile
                    If frmMapEditor.optMask.Value = True Then .Mask = Tile
                    If frmMapEditor.optAnim.Value = True Then .Anim = Tile
                    If frmMapEditor.optMask2.Value = True Then .Mask2 = Tile
                    If frmMapEditor.optM2anim.Value = True Then .M2Anim = Tile
                    If frmMapEditor.optFringe.Value = True Then .Fringe = Tile
                    If frmMapEditor.optFAnim.Value = True Then .FAnim = Tile
                    If frmMapEditor.optFringe2.Value = True Then .Fringe2 = Tile
                    If frmMapEditor.optF2anim.Value = True Then .F2Anim = Tile
                End With
            Next
        Next
    End If
End Sub

Public Sub MapEditorClearLayer()
Dim YesNo As Long, X As Long, Y As Long

    YesNo = MsgBox("Are you sure you wish to clear this layer?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    If frmMapEditor.optGround.Value = True Then .Ground = 0
                    If frmMapEditor.optMask.Value = True Then .Mask = 0
                    If frmMapEditor.optAnim.Value = True Then .Anim = 0
                    If frmMapEditor.optMask2.Value = True Then .Mask2 = 0
                    If frmMapEditor.optM2anim.Value = True Then .M2Anim = 0
                    If frmMapEditor.optFringe.Value = True Then .Fringe = 0
                    If frmMapEditor.optFAnim.Value = True Then .FAnim = 0
                    If frmMapEditor.optFringe2.Value = True Then .Fringe2 = 0
                    If frmMapEditor.optF2anim.Value = True Then .F2Anim = 0
                End With
            Next
        Next
    End If
End Sub

Public Sub MapEditorClearAttribs()
Dim YesNo As Long, X As Long, Y As Long

    YesNo = MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Type = 0
            Next
        Next
    End If
End Sub

Public Sub ItemEditor()
Dim i As Long
    
    InItemsEditor = True
        
    frmIndex.Show
    frmIndex.lstIndex.Clear
    
    ' Add the names
    For i = 1 To MAX_ITEMS
        frmIndex.lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
    Next
    
    frmIndex.lstIndex.ListIndex = 0
End Sub

Public Sub ItemEditorInit()
Dim i As Long

    frmItemEditor.scrlPic.Max = (DDSD_Item.lHeight \ PIC_Y) - 1
    
    ' set the scrl to 0 for the initial value
    frmItemEditor.scrlAnim.Value = 0
    
    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).Name)
    frmItemEditor.scrlPic.Value = Item(EditorIndex).Pic
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    
    frmItemEditor.scrlLevel.Value = Item(EditorIndex).LevelReq
    
    frmItemEditor.scrlStackMax.Max = MAX_INTEGER
    
    ' Item Bound
    For i = 0 To ItemBind.ItemBind_Count
        frmItemEditor.cmbBind.AddItem BindName(i)
    Next
    frmItemEditor.cmbBind.ListIndex = Item(EditorIndex).Bound
    
    ' Equipment
    For i = 1 To Slots.Slot_Count
        frmItemEditor.cmbEquipmentType.AddItem EquipmentName(i)
    Next
    frmItemEditor.cmbEquipmentType.ListIndex = 0
    
    For i = 1 To Stats.Stat_Count
        ' Stat Reqs
        Load frmItemEditor.lblStatName(i)
        Load frmItemEditor.scrlStat(i)
        Load frmItemEditor.lblStat(i)
        
        frmItemEditor.lblStatName(i).Visible = True
        frmItemEditor.scrlStat(i).Visible = True
        frmItemEditor.lblStat(i).Visible = True
        frmItemEditor.lblStatName(i).Top = frmItemEditor.lblStatName(i).Top + (frmItemEditor.lblStatName(i).Height * (i - 1))
        frmItemEditor.scrlStat(i).Top = frmItemEditor.scrlStat(i).Top + (frmItemEditor.scrlStat(i).Height * (i - 1))
        frmItemEditor.lblStat(i).Top = frmItemEditor.lblStat(i).Top + (frmItemEditor.lblStat(i).Height * (i - 1))
        
        frmItemEditor.lblStatName(i).Caption = StatName(i)
        frmItemEditor.scrlStat(i).Value = Item(EditorIndex).StatReq(i)
        
        ' Mod Stats
        Load frmItemEditor.lblModStatName(i)
        Load frmItemEditor.scrlModStat(i)
        Load frmItemEditor.txtModStat(i)
        
        frmItemEditor.lblModStatName(i).Visible = True
        frmItemEditor.scrlModStat(i).Visible = True
        frmItemEditor.txtModStat(i).Visible = True
        frmItemEditor.lblModStatName(i).Top = frmItemEditor.txtModStat(i).Top + (frmItemEditor.txtModStat(i).Height * (i - 1))
        frmItemEditor.scrlModStat(i).Top = frmItemEditor.txtModStat(i).Top + (frmItemEditor.txtModStat(i).Height * (i - 1))
        frmItemEditor.txtModStat(i).Top = frmItemEditor.txtModStat(i).Top + (frmItemEditor.txtModStat(i).Height * (i - 1))
        
        frmItemEditor.lblModStatName(i).Caption = StatName(i)
        frmItemEditor.scrlModStat(i).Value = Item(EditorIndex).ModStat(i)
    Next
    
    For i = 1 To Vitals.Vital_Count
        Load frmItemEditor.lblModVitalName(i)
        Load frmItemEditor.scrlModVital(i)
        Load frmItemEditor.txtModVital(i)
        
        frmItemEditor.lblModVitalName(i).Visible = True
        frmItemEditor.scrlModVital(i).Visible = True
        frmItemEditor.txtModVital(i).Visible = True
        frmItemEditor.lblModVitalName(i).Top = frmItemEditor.txtModVital(i).Top + (frmItemEditor.txtModVital(i).Height * (i - 1))
        frmItemEditor.scrlModVital(i).Top = frmItemEditor.txtModVital(i).Top + (frmItemEditor.txtModVital(i).Height * (i - 1))
        frmItemEditor.txtModVital(i).Top = frmItemEditor.txtModVital(i).Top + (frmItemEditor.txtModVital(i).Height * (i - 1))
        
        frmItemEditor.lblModVitalName(i).Caption = VitalName(i)
        frmItemEditor.scrlModVital(i).Value = Item(EditorIndex).ModVital(i)
    Next
    
    For i = 0 To MAX_CLASSES
        If i > 0 Then Load frmItemEditor.chkClass(i)
        frmItemEditor.chkClass(i).Visible = True
        frmItemEditor.chkClass(i).Top = frmItemEditor.chkClass(i).Top + (frmItemEditor.chkClass(i).Height * i)
        
        frmItemEditor.chkClass(i).Caption = Trim$(Class(i).Name)
        ' If the flag is true, set the checkbox
        If Item(EditorIndex).ClassReq And (2 ^ i) Then
            frmItemEditor.chkClass(i).Value = 1
        Else
            frmItemEditor.chkClass(i).Value = 0
        End If
    Next
    
    ' Class Req Frame
    frmItemEditor.frmClassReq.Height = frmItemEditor.frmClassReq.Height + (MAX_CLASSES * frmItemEditor.chkClass(0).Height)
    
    ' Stat Req Frame
    frmItemEditor.frmStatReq.Top = frmItemEditor.frmClassReq.Top + frmItemEditor.frmClassReq.Height + 10
    frmItemEditor.frmStatReq.Height = frmItemEditor.lblStat(1).Height + ((Stats.Stat_Count + 1) * frmItemEditor.lblStat(1).Height)
    
    ' Mod Vital Frame
    frmItemEditor.fraModVitals.Top = frmItemEditor.frmStatReq.Top + frmItemEditor.frmStatReq.Height + 10
    frmItemEditor.fraModVitals.Height = frmItemEditor.fraModVitals.Height + ((Vitals.Vital_Count - 1) * frmItemEditor.txtModVital(1).Height)
    
    ' Mod Stat Frame
    frmItemEditor.fraModStat.Top = frmItemEditor.fraModVitals.Top + frmItemEditor.fraModVitals.Height + 10
    frmItemEditor.fraModStat.Height = frmItemEditor.txtModStat(1).Height + ((Stats.Stat_Count + 1) * frmItemEditor.txtModStat(1).Height)
    
    ' Check if we need to extend the actual form
    If frmItemEditor.fraModStat.Top + frmItemEditor.fraModStat.Height > (frmItemEditor.Height - frmItemEditor.ScaleHeight) Then
        frmItemEditor.Height = (frmItemEditor.Height - frmItemEditor.ScaleHeight) + (frmItemEditor.fraModStat.Top + frmItemEditor.fraModStat.Height) + 125
    End If
    
    frmItemEditor.fraModStat.Visible = False
    frmItemEditor.fraModVitals.Visible = False
    frmItemEditor.fraSpell.Visible = False
    frmItemEditor.fraStack.Visible = False
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_EQUIPMENT) Then
        frmItemEditor.fraModStat.Visible = True
        frmItemEditor.fraModVitals.Visible = True
        frmItemEditor.fraEquipmentType.Visible = True
        frmItemEditor.cmbEquipmentType.ListIndex = Item(EditorIndex).Data1 - 1
        If frmItemEditor.chkStack Then
            frmItemEditor.chkStack.Value = 0
            frmItemEditor.scrlStackMax.Value = 0
        End If
        If frmItemEditor.cmbEquipmentType.ListIndex + 1 = Slots.Weapon Then
            frmItemEditor.frmAnim.Visible = True
            frmItemEditor.scrlAnim.Value = Item(EditorIndex).Data2
        End If
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_POTION) Then
        frmItemEditor.fraModVitals.Visible = True
        frmItemEditor.fraStack.Visible = True
        frmItemEditor.chkStack.Value = Item(EditorIndex).Stack
        frmItemEditor.scrlStackMax.Value = Item(EditorIndex).StackMax
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_KEY) Or (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_NONE) Then
        frmItemEditor.fraStack.Visible = True
        frmItemEditor.chkStack.Value = Item(EditorIndex).Stack
        frmItemEditor.scrlStackMax.Value = Item(EditorIndex).StackMax
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
    End If
    
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    Item(EditorIndex).Name = frmItemEditor.txtName.Text
    Item(EditorIndex).Pic = frmItemEditor.scrlPic.Value
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
    
    Item(EditorIndex).LevelReq = frmItemEditor.scrlLevel.Value
    
    ' Get the item bind
    Item(EditorIndex).Bound = frmItemEditor.cmbBind.ListIndex
    ' Check if BOE, error out
    If Item(EditorIndex).Bound = ItemBind.BindOnEquip Then
        If Item(EditorIndex).Type <> ITEM_TYPE_EQUIPMENT Then
            MsgBox BindName(BindOnEquip) & " can only be set for equipment.", vbOKOnly
            Exit Sub
        End If
    End If
    
    For i = 1 To Stats.Stat_Count
        Item(EditorIndex).StatReq(i) = frmItemEditor.scrlStat(i).Value
        Item(EditorIndex).ModStat(i) = frmItemEditor.scrlModStat(i).Value
    Next
    
    For i = 1 To Vitals.Vital_Count
        Item(EditorIndex).ModVital(i) = frmItemEditor.scrlModVital(i).Value
    Next
    
    For i = 0 To MAX_CLASSES
        If frmItemEditor.chkClass(i).Value Then
            Item(EditorIndex).ClassReq = Item(EditorIndex).ClassReq Or (2 ^ i)
        Else
            Item(EditorIndex).ClassReq = Item(EditorIndex).ClassReq And Not (2 ^ i)
        End If
    Next
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_EQUIPMENT) Then
        Item(EditorIndex).Data1 = frmItemEditor.cmbEquipmentType.ListIndex + 1
        Item(EditorIndex).Data2 = frmItemEditor.scrlAnim.Value
        Item(EditorIndex).Stack = 0
        Item(EditorIndex).StackMax = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_POTION) Then
        Item(EditorIndex).Stack = frmItemEditor.chkStack.Value
        Item(EditorIndex).StackMax = frmItemEditor.scrlStackMax.Value
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_KEY) Or (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_NONE) Then
        Item(EditorIndex).Stack = frmItemEditor.chkStack.Value
        Item(EditorIndex).StackMax = frmItemEditor.scrlStackMax.Value
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
    End If
'
    SendSaveItem EditorIndex
    InItemsEditor = False
    Unload frmItemEditor
    
    ItemEditor
    
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmItemEditor
    
    ItemEditor
End Sub

Public Sub EmoticonEditorInit()
    
    frmEmoticonEditor.scrlEmoticon.Max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.Max = (DDSD_Emoticon.lHeight \ PIC_Y) - 1
    
    frmEmoticonEditor.scrlEmoticon.Value = Emoticons(EditorIndex).Pic
    
    frmEmoticonEditor.txtCommand.Text = Trim$(Emoticons(EditorIndex).Command)
    
    frmEmoticonEditor.Show vbModal
End Sub

Public Sub EmoticonEditorOk()

    Emoticons(EditorIndex).Pic = frmEmoticonEditor.scrlEmoticon.Value
    
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex).Command = ""
    End If
    
    Call SendSaveEmoticon(EditorIndex)
    Call EmoticonEditorCancel
End Sub

Public Sub EmoticonEditorCancel()
    InEmoticonEditor = False
    Unload frmEmoticonEditor
End Sub

Public Sub NpcEditor()
Dim i As Long

    InNpcEditor = True
        
    frmIndex.Show
    frmIndex.lstIndex.Clear
    
    ' Add the names
    For i = 1 To MAX_NPCS
        frmIndex.lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
    Next
    
    frmIndex.lstIndex.ListIndex = 0
End Sub

Public Sub NpcEditorInit()
Dim X As Long
    
    frmNpcEditor.scrlSprite.Max = (DDSD_Sprite.lHeight \ PIC_Y) - 1
    
    frmNpcEditor.txtName.Text = Trim$(Npc(EditorIndex).Name)
    frmNpcEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = Str$(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    
    frmNpcEditor.txtChance.Text = Npc(EditorIndex).Drop(1).Chance
    frmNpcEditor.scrlNum.Max = MAX_ITEMS
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).Drop(1).Item
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).Drop(1).ItemValue
    
    frmNpcEditor.txtChance2.Text = Npc(EditorIndex).Drop(2).Chance
    frmNpcEditor.scrlNum2.Max = MAX_ITEMS
    frmNpcEditor.scrlNum2.Value = Npc(EditorIndex).Drop(2).Item
    frmNpcEditor.scrlValue2.Value = Npc(EditorIndex).Drop(2).ItemValue
    
    frmNpcEditor.txtChance3.Text = Npc(EditorIndex).Drop(3).Chance
    frmNpcEditor.scrlNum3.Max = MAX_ITEMS
    frmNpcEditor.scrlNum3.Value = Npc(EditorIndex).Drop(3).Item
    frmNpcEditor.scrlValue3.Value = Npc(EditorIndex).Drop(3).ItemValue
    
    frmNpcEditor.txtChance4.Text = Npc(EditorIndex).Drop(4).Chance
    frmNpcEditor.scrlNum4.Max = MAX_ITEMS
    frmNpcEditor.scrlNum4.Value = Npc(EditorIndex).Drop(4).Item
    frmNpcEditor.scrlValue4.Value = Npc(EditorIndex).Drop(4).ItemValue
    
    For X = 1 To Stats.Stat_Count
        frmNpcEditor.lblStatName(X - 1) = StatName(X)
        frmNpcEditor.scrlStat(X - 1) = Npc(EditorIndex).Stat(X)
    Next
    
    frmNpcEditor.txtHP.Text = Npc(EditorIndex).MaxHP
    frmNpcEditor.txtExp.Text = Npc(EditorIndex).MaxEXP
    
    frmNpcEditor.scrlLevel.Value = Npc(EditorIndex).Level
   
    If Npc(EditorIndex).MovementSpeed = 0 Then Npc(EditorIndex).MovementSpeed = 1
    frmNpcEditor.scrlMovementSpeed.Value = Npc(EditorIndex).MovementSpeed
    
    frmNpcEditor.cmbMovementFrequency.ListIndex = Npc(EditorIndex).MovementFrequency - 1
    
    frmNpcEditor.cmbShop.AddItem "No Shop"
    For X = 1 To MAX_SHOPS
        frmNpcEditor.cmbShop.AddItem X & ": " & Trim$(Shop(X).Name)
    Next
    If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
        frmNpcEditor.cmbShop.ListIndex = Npc(EditorIndex).Stat(Stats.Wisdom)
    End If
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    Npc(EditorIndex).Name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.Value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    
    Npc(EditorIndex).Drop(1).Chance = Val(frmNpcEditor.txtChance.Text)
    Npc(EditorIndex).Drop(1).Item = frmNpcEditor.scrlNum.Value
    Npc(EditorIndex).Drop(1).ItemValue = frmNpcEditor.scrlValue.Value
    
    Npc(EditorIndex).Drop(2).Chance = Val(frmNpcEditor.txtChance2.Text)
    Npc(EditorIndex).Drop(2).Item = frmNpcEditor.scrlNum2.Value
    Npc(EditorIndex).Drop(2).ItemValue = frmNpcEditor.scrlValue2.Value
    
    Npc(EditorIndex).Drop(3).Chance = Val(frmNpcEditor.txtChance3.Text)
    Npc(EditorIndex).Drop(3).Item = frmNpcEditor.scrlNum3.Value
    Npc(EditorIndex).Drop(3).ItemValue = frmNpcEditor.scrlValue3.Value
    
    Npc(EditorIndex).Drop(4).Chance = Val(frmNpcEditor.txtChance4.Text)
    Npc(EditorIndex).Drop(4).Item = frmNpcEditor.scrlNum4.Value
    Npc(EditorIndex).Drop(4).ItemValue = frmNpcEditor.scrlValue4.Value
    
    For i = 1 To Stats.Stat_Count
        Npc(EditorIndex).Stat(i) = frmNpcEditor.scrlStat(i - 1).Value
    Next
    
    
    Npc(EditorIndex).MaxHP = frmNpcEditor.txtHP.Text
    Npc(EditorIndex).MaxEXP = frmNpcEditor.txtExp.Text
    
    Npc(EditorIndex).Level = frmNpcEditor.scrlLevel.Value
    
    Npc(EditorIndex).MovementSpeed = frmNpcEditor.scrlMovementSpeed.Value
    
    Npc(EditorIndex).MovementFrequency = frmNpcEditor.cmbMovementFrequency.ListIndex + 1
    
    If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
        Npc(EditorIndex).Stat(Stats.Wisdom) = frmNpcEditor.cmbShop.ListIndex
    End If
    
    Call SendSaveNpc(EditorIndex)
    InNpcEditor = False
    Unload frmNpcEditor
    
    NpcEditor
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmNpcEditor
    
    NpcEditor
End Sub

Public Sub ShopEditor()
Dim i As Long

    InShopEditor = True
    
    frmIndex.Show
    frmIndex.lstIndex.Clear
    
    ' Add the names
    For i = 1 To MAX_SHOPS
        frmIndex.lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
    Next
    
    frmIndex.lstIndex.ListIndex = 0
End Sub

Public Sub ShopEditorInit()
Dim i As Long

    frmShopEditor.fraShop.Visible = False
    frmShopEditor.fraInn.Visible = False
    
    frmShopEditor.txtName.Text = Trim$(Shop(EditorIndex).Name)
    frmShopEditor.txtJoinSay.Text = Trim$(Shop(EditorIndex).JoinSay)
    frmShopEditor.cmbType.ListIndex = Shop(EditorIndex).Type
    
    If Shop(EditorIndex).Type = SHOP_TYPE_SHOP Then
        frmShopEditor.fraShop.Visible = True
    
        frmShopEditor.cmbItemGive.Clear
        frmShopEditor.cmbItemGive.AddItem "None"
        frmShopEditor.cmbItemGet.Clear
        frmShopEditor.cmbItemGet.AddItem "None"
        For i = 1 To MAX_ITEMS
            frmShopEditor.cmbItemGive.AddItem i & ": " & Trim$(Item(i).Name)
            frmShopEditor.cmbItemGet.AddItem i & ": " & Trim$(Item(i).Name)
        Next
        frmShopEditor.cmbItemGive.ListIndex = 0
        frmShopEditor.cmbItemGet.ListIndex = 0
    ElseIf Shop(EditorIndex).Type = SHOP_TYPE_INN Then
        frmShopEditor.fraInn.Visible = True
    
        frmShopEditor.txtMap.Text = Shop(EditorIndex).BindPoint.Map
        frmShopEditor.scrlX.Value = Shop(EditorIndex).BindPoint.X
        frmShopEditor.scrlY.Value = Shop(EditorIndex).BindPoint.Y
    End If
    
    Call UpdateShopTrade
    
    frmShopEditor.Show vbModal
End Sub

Public Sub UpdateShopTrade()
Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long
    
    frmShopEditor.lstTradeItem.Clear
    For i = 1 To MAX_TRADES
        GetItem = Shop(EditorIndex).TradeItem(i).GetItem
        GetValue = Shop(EditorIndex).TradeItem(i).GetValue
        GiveItem = Shop(EditorIndex).TradeItem(i).GiveItem
        GiveValue = Shop(EditorIndex).TradeItem(i).GiveValue
        
        If GetItem > 0 And GiveItem > 0 Then
            frmShopEditor.lstTradeItem.AddItem i & ": " & GiveValue & " " & Trim$(Item(GiveItem).Name) & " for " & GetValue & " " & Trim$(Item(GetItem).Name)
        Else
            frmShopEditor.lstTradeItem.AddItem "Empty Trade Slot"
        End If
    Next
    frmShopEditor.lstTradeItem.ListIndex = 0
End Sub

Public Sub ShopEditorOk()

    Shop(EditorIndex).Name = frmShopEditor.txtName.Text
    Shop(EditorIndex).JoinSay = frmShopEditor.txtJoinSay.Text
    Shop(EditorIndex).Type = frmShopEditor.cmbType.ListIndex
    
    Shop(EditorIndex).BindPoint.Map = Val(frmShopEditor.txtMap.Text)
    Shop(EditorIndex).BindPoint.X = frmShopEditor.scrlX.Value
    Shop(EditorIndex).BindPoint.Y = frmShopEditor.scrlY.Value
    
    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmShopEditor
    
    ShopEditor
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmShopEditor
    
    ShopEditor
End Sub

'//////////////////
'// Spell Editor //
'//////////////////
Public Sub SpellEditor()
Dim i As Long

    InSpellEditor = True
        
    frmIndex.Show
    frmIndex.lstIndex.Clear
    
    ' Add the names
    For i = 1 To MAX_SPELLS
        frmIndex.lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
    Next
    
    frmIndex.lstIndex.ListIndex = 0
End Sub

Public Sub SpellEditorInit()
Dim i As Long
      
    ' set the scrl to 0 for the initial value
    frmSpellEditor.scrlAnim.Value = 0
    
    For i = 0 To MAX_CLASSES
        If i > 0 Then Load frmSpellEditor.chkClass(i)
        frmSpellEditor.chkClass(i).Visible = True
        frmSpellEditor.chkClass(i).Top = frmSpellEditor.chkClass(i).Top + (frmSpellEditor.chkClass(i).Height * i)
        
        frmSpellEditor.chkClass(i).Caption = Trim$(Class(i).Name)
        ' If the flag is true, set the checkbox
        If Spell(EditorIndex).ClassReq And (2 ^ i) Then
            frmSpellEditor.chkClass(i).Value = 1
        Else
            frmSpellEditor.chkClass(i).Value = 0
        End If
    Next
    
    For i = 1 To Vitals.Vital_Count
        ' Req Vitals
        Load frmSpellEditor.lblVitalReqName(i)
        Load frmSpellEditor.scrlVitalReq(i)
        Load frmSpellEditor.txtVitalReq(i)
        
        frmSpellEditor.lblVitalReqName(i).Visible = True
        frmSpellEditor.scrlVitalReq(i).Visible = True
        frmSpellEditor.txtVitalReq(i).Visible = True
        frmSpellEditor.lblVitalReqName(i).Top = frmSpellEditor.txtVitalReq(i).Top + (frmSpellEditor.txtVitalReq(i).Height * (i - 1))
        frmSpellEditor.scrlVitalReq(i).Top = frmSpellEditor.txtVitalReq(i).Top + (frmSpellEditor.txtVitalReq(i).Height * (i - 1))
        frmSpellEditor.txtVitalReq(i).Top = frmSpellEditor.txtVitalReq(i).Top + (frmSpellEditor.txtVitalReq(i).Height * (i - 1))
        
        frmSpellEditor.lblVitalReqName(i).Caption = VitalName(i)
        frmSpellEditor.scrlVitalReq(i).Value = Spell(EditorIndex).VitalReq(i)
        
        ' Mod Vitals
        Load frmSpellEditor.lblModVitalName(i)
        Load frmSpellEditor.scrlModVital(i)
        Load frmSpellEditor.txtModVital(i)
        
        frmSpellEditor.lblModVitalName(i).Visible = True
        frmSpellEditor.scrlModVital(i).Visible = True
        frmSpellEditor.txtModVital(i).Visible = True
        frmSpellEditor.lblModVitalName(i).Top = frmSpellEditor.txtModVital(i).Top + (frmSpellEditor.txtModVital(i).Height * (i - 1))
        frmSpellEditor.scrlModVital(i).Top = frmSpellEditor.txtModVital(i).Top + (frmSpellEditor.txtModVital(i).Height * (i - 1))
        frmSpellEditor.txtModVital(i).Top = frmSpellEditor.txtModVital(i).Top + (frmSpellEditor.txtModVital(i).Height * (i - 1))
        
        frmSpellEditor.lblModVitalName(i).Caption = VitalName(i)
        frmSpellEditor.scrlModVital(i).Value = Spell(EditorIndex).ModVital(i)
    Next
    
    For i = 1 To Stats.Stat_Count
        ' Mod Stats
        Load frmSpellEditor.lblModStatName(i)
        Load frmSpellEditor.scrlModStat(i)
        Load frmSpellEditor.txtModStat(i)
        
        frmSpellEditor.lblModStatName(i).Visible = True
        frmSpellEditor.scrlModStat(i).Visible = True
        frmSpellEditor.txtModStat(i).Visible = True
        frmSpellEditor.lblModStatName(i).Top = frmSpellEditor.txtModStat(i).Top + (frmSpellEditor.txtModStat(i).Height * (i - 1))
        frmSpellEditor.scrlModStat(i).Top = frmSpellEditor.txtModStat(i).Top + (frmSpellEditor.txtModStat(i).Height * (i - 1))
        frmSpellEditor.txtModStat(i).Top = frmSpellEditor.txtModStat(i).Top + (frmSpellEditor.txtModStat(i).Height * (i - 1))
        
        frmSpellEditor.lblModStatName(i).Caption = StatName(i)
        frmSpellEditor.scrlModStat(i).Value = Spell(EditorIndex).ModStat(i)
    Next
    
    ' Class Req Frame
    frmSpellEditor.frmClassReq.Height = frmSpellEditor.frmClassReq.Height + (MAX_CLASSES * frmSpellEditor.chkClass(0).Height)
    
    ' Req Vital Frame
    frmSpellEditor.fraVitalReq.Top = frmSpellEditor.frmClassReq.Top + frmSpellEditor.frmClassReq.Height + 10
    frmSpellEditor.fraVitalReq.Height = frmSpellEditor.txtVitalReq(1).Height + ((Vitals.Vital_Count + 1) * frmSpellEditor.txtVitalReq(1).Height)
    
    ' Mod Vital Frame
    frmSpellEditor.fraModVitals.Height = frmSpellEditor.fraModVitals.Height + ((Vitals.Vital_Count - 1) * frmSpellEditor.txtModVital(1).Height)
    
    ' Mod Stat Frame
    frmSpellEditor.fraModStat.Top = frmSpellEditor.fraModVitals.Top + frmSpellEditor.fraModVitals.Height + 10
    frmSpellEditor.fraModStat.Height = frmSpellEditor.txtModStat(1).Height + ((Stats.Stat_Count + 1) * frmSpellEditor.txtModStat(1).Height)
    
    ' Check if we need to extend the actual form
    If frmSpellEditor.fraModStat.Top + frmSpellEditor.fraModStat.Height > (frmSpellEditor.Height - frmSpellEditor.ScaleHeight) Then
        frmSpellEditor.Height = (frmSpellEditor.Height - frmSpellEditor.ScaleHeight) + (frmSpellEditor.fraModStat.Top + frmSpellEditor.fraModStat.Height) + 125
    End If
    
    frmSpellEditor.scrlAnim.Value = Spell(EditorIndex).Animation

    frmSpellEditor.scrlCastTime.Value = Spell(EditorIndex).CastTime
    frmSpellEditor.scrlCooldown.Value = Spell(EditorIndex).Cooldown
    frmSpellEditor.scrlRange.Value = Spell(EditorIndex).Range
    frmSpellEditor.scrlTickCount.Value = Spell(EditorIndex).TickCount
    frmSpellEditor.scrlTickUpdate.Value = Spell(EditorIndex).TickUpdate
    
    For i = 0 To Targets.Target_Count - 1
        frmSpellEditor.chkTargets(i).Caption = TargetName(2 ^ i)
        ' If the flag is true, set the checkbox
        If Spell(EditorIndex).TargetFlags And (2 ^ i) Then frmSpellEditor.chkTargets(i).Value = 1
    Next
    
    ' Load in some values
    frmSpellEditor.txtName.Text = Trim$(Spell(EditorIndex).Name)
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    frmSpellEditor.scrlLevel.Value = Spell(EditorIndex).LevelReq
    
    ' Different things for different types of spells
    Select Case Spell(EditorIndex).Type
        ' Check if the spell is revive to change the min/max on the vitals
        Case SPELL_TYPE_REVIVE
            For i = 1 To Vitals.Vital_Count
                frmSpellEditor.scrlModVital(i).Min = 0
                frmSpellEditor.scrlModVital(i).Max = 100
                frmSpellEditor.scrlModVital(i).Value = Spell(EditorIndex).ModVital(i)
            Next
    End Select
    
    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
Dim i As Long

    Spell(EditorIndex).Name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevel.Value
    
    For i = 0 To MAX_CLASSES
        If frmSpellEditor.chkClass(i).Value Then
            Spell(EditorIndex).ClassReq = Spell(EditorIndex).ClassReq Or (2 ^ i)
        Else
            Spell(EditorIndex).ClassReq = Spell(EditorIndex).ClassReq And Not (2 ^ i)
        End If
    Next
    
    For i = 1 To Vitals.Vital_Count
        Spell(EditorIndex).VitalReq(i) = frmSpellEditor.scrlVitalReq(i).Value
        Spell(EditorIndex).ModVital(i) = frmSpellEditor.scrlModVital(i).Value
    Next
    
    For i = 1 To Stats.Stat_Count
        Spell(EditorIndex).ModStat(i) = frmSpellEditor.scrlModStat(i).Value
    Next
    
    Spell(EditorIndex).Animation = frmSpellEditor.scrlAnim.Value
    
    Spell(EditorIndex).CastTime = frmSpellEditor.scrlCastTime.Value
    Spell(EditorIndex).Cooldown = frmSpellEditor.scrlCooldown.Value
    Spell(EditorIndex).Range = frmSpellEditor.scrlRange.Value
    Spell(EditorIndex).TickCount = frmSpellEditor.scrlTickCount.Value
    Spell(EditorIndex).TickUpdate = frmSpellEditor.scrlTickUpdate.Value
    
    For i = 0 To Targets.Target_Count - 1
        If frmSpellEditor.chkTargets(i).Value Then
            Spell(EditorIndex).TargetFlags = Spell(EditorIndex).TargetFlags Or (2 ^ i)
        Else
            Spell(EditorIndex).TargetFlags = Spell(EditorIndex).TargetFlags And Not (2 ^ i)
        End If
    Next
    
    SendSaveSpell EditorIndex
    InSpellEditor = False
    Unload frmSpellEditor
    
    SpellEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
    
    SpellEditor
End Sub

'//////////////////////
'// Animation Editor //
'//////////////////////
Public Sub AnimationEditorInit()

    frmAnimationEditor.txtName.Text = Trim$(Animation(EditorIndex).Name)
    
    frmAnimationEditor.scrlSprite.Value = Animation(EditorIndex).Animation
    frmAnimationEditor.scrlSpeed.Value = Animation(EditorIndex).AnimationSpeed
    frmAnimationEditor.scrlFrames.Value = Animation(EditorIndex).AnimationFrames
    
    If Animation(EditorIndex).AnimationSize = 1 Then
        frmAnimationEditor.opt32.Value = True
        frmAnimationEditor.scrlSprite.Max = (DDSD_Animation.lHeight \ PIC_Y) - 1
    ElseIf Animation(EditorIndex).AnimationSize = 2 Then
        frmAnimationEditor.opt64.Value = True
        frmAnimationEditor.scrlSprite.Max = (DDSD_Animation2.lHeight \ PIC_Y) - 1
    End If
    
    If Animation(EditorIndex).AnimationLayer = 0 Then
        frmAnimationEditor.optBelow.Value = True
        frmAnimationEditor.optAbove.Value = False
    ElseIf Animation(EditorIndex).AnimationLayer = 1 Then
        frmAnimationEditor.optBelow.Value = False
        frmAnimationEditor.optAbove.Value = True
    End If
        
    frmAnimationEditor.Show vbModal
End Sub

Public Sub AnimationEditorOk()

    Animation(EditorIndex).Name = frmAnimationEditor.txtName.Text
    Animation(EditorIndex).Animation = frmAnimationEditor.scrlSprite.Value
    Animation(EditorIndex).AnimationSpeed = frmAnimationEditor.scrlSpeed
    Animation(EditorIndex).AnimationFrames = frmAnimationEditor.scrlFrames
    
    If frmAnimationEditor.opt32.Value = True Then
        Animation(EditorIndex).AnimationSize = 1
    ElseIf frmAnimationEditor.opt64.Value = True Then
        Animation(EditorIndex).AnimationSize = 2
    End If
    
    If frmAnimationEditor.optBelow.Value = True Then
        Animation(EditorIndex).AnimationLayer = 0
    ElseIf frmAnimationEditor.optAbove.Value = True Then
        Animation(EditorIndex).AnimationLayer = 1
    End If
    
    SendSaveAnimation EditorIndex
    InAnimationEditor = False
    Unload frmAnimationEditor
End Sub

Public Sub AnimationEditorCancel()
    InAnimationEditor = False
    Unload frmAnimationEditor
End Sub

Public Sub SetTextBox(ByRef txtBox As TextBox, ByRef scrlBar As HScrollBar)
Dim Temp As String
    Temp = txtBox.Text
    If Not IsNumeric(Temp) Then
        Temp = scrlBar.Value
    End If

    If Val(Temp) > scrlBar.Max Then txtBox.Text = scrlBar.Max
    If Val(Temp) < scrlBar.Min Then txtBox.Text = scrlBar.Min
    scrlBar.Value = Val(txtBox.Text)
End Sub
