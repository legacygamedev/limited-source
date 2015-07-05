Attribute VB_Name = "modGameEditors"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' ////////////////
' // Map Editor //
' ////////////////

Public Sub MapEditorInit()
    InMapEditor = True
    frmMirage.picMapEditor.Visible = True
     
    Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    
    frmMirage.lblTileset = Map.TileSet
    frmMirage.scrlTileSet = Map.TileSet
     
    frmMirage.Width = 14175
     
    Call BltMapEditor
    Call BltMapEditorTilePreview
    
    frmMirage.scrlPicture.Max = (frmMirage.picBackSelect.Height \ PIC_Y) - (frmMirage.picBack.Height \ PIC_Y)

End Sub

Public Sub MapEditorMouseDown(Button As Integer)
    If Not isInBounds Then Exit Sub

    If Button = vbLeftButton Then
        If frmMirage.optLayers.Value Then
            
            With Map.Tile(CurX, CurY)
                If frmMirage.optGround.Value Then .Ground = EditorTileY * TILESHEET_WIDTH + EditorTileX
                If frmMirage.optMask.Value Then .Mask = EditorTileY * TILESHEET_WIDTH + EditorTileX
                If frmMirage.optMask2.Value Then .Mask2 = EditorTileY * TILESHEET_WIDTH + EditorTileX
                If frmMirage.optAnim.Value Then .Anim = EditorTileY * TILESHEET_WIDTH + EditorTileX
                If frmMirage.optFringe.Value Then .Fringe = EditorTileY * TILESHEET_WIDTH + EditorTileX
                If frmMirage.optFringe2.Value Then .Fringe2 = EditorTileY * TILESHEET_WIDTH + EditorTileX
            End With
        Else
            With Map.Tile(CurX, CurY)
                If frmMirage.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                If frmMirage.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                End If
                If frmMirage.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                End If
                If frmMirage.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
                If frmMirage.optKey.Value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                End If
                If frmMirage.optKeyOpen.Value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                End If
            End With
        End If
    End If
    
    If Button = vbRightButton Then
        If frmMirage.optLayers.Value Then
            With Map.Tile(CurX, CurY)
                If frmMirage.optGround.Value Then .Ground = 0
                If frmMirage.optMask.Value Then .Mask = 0
                If frmMirage.optMask2.Value Then .Mask2 = 0
                If frmMirage.optAnim.Value Then .Anim = 0
                If frmMirage.optFringe.Value Then .Fringe = 0
                If frmMirage.optFringe2.Value Then .Fringe2 = 0
            End With
        Else
            With Map.Tile(CurX, CurY)
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With
        End If
    End If

End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
        
        frmMirage.shpSelected.Top = EditorTileY * PIC_Y
        frmMirage.shpSelected.Left = EditorTileX * PIC_Y
        
        Call BltMapEditorTilePreview
    End If
End Sub

Public Sub MapEditorTileScroll()
    frmMirage.picBackSelect.Top = (frmMirage.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub MapEditorSend()
    Call SendMap
    Call MapEditorCancel
End Sub

Public Sub MapEditorCancel()
    Call LoadMap(GetPlayerMap(MyIndex))
    InMapEditor = False
    frmMirage.picMapEditor.Visible = False
    Set DDS_Misc = Nothing
    
    frmMirage.Width = 10080
End Sub

Public Sub MapEditorClearLayer()
Dim X As Long
Dim Y As Long

    ' Ground layer
    If frmMirage.optGround.Value Then
        If MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Ground = 0
                Next
            Next
        End If
    End If

    ' Mask layer
    If frmMirage.optMask.Value Then
        If MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Mask = 0
                Next
            Next
        End If
    End If

    ' Animation layer
    If frmMirage.optAnim.Value Then
        If MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Anim = 0
                Next
            Next
        End If
    End If
    
    ' Mask layer
    If frmMirage.optMask2.Value Then
        If MsgBox("Are you sure you wish to clear the mask2 layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Mask2 = 0
                Next
            Next
        End If
    End If

    ' Fringe layer
    If frmMirage.optFringe.Value Then
        If MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Fringe = 0
                Next
            Next
        End If
    End If
    
    ' Fringe layer
    If frmMirage.optFringe2.Value Then
        If MsgBox("Are you sure you wish to clear the fringe2 layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Fringe2 = 0
                Next
            Next
        End If
    End If
    
End Sub

Public Sub MapEditorFillLayer()
Dim X As Long
Dim Y As Long

    ' Ground layer
    If frmMirage.optGround.Value Then
        If MsgBox("Are you sure you wish to fill the ground layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Ground = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If

    ' Mask layer
    If frmMirage.optMask.Value Then
        If MsgBox("Are you sure you wish to fill the mask layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Mask = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If

    ' Animation layer
    If frmMirage.optAnim.Value Then
        If MsgBox("Are you sure you wish to fill the animation layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Anim = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If
    
    ' Mask layer
    If frmMirage.optMask2.Value Then
        If MsgBox("Are you sure you wish to fill the mask2 layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Mask2 = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If

    ' Fringe layer
    If frmMirage.optFringe.Value Then
        If MsgBox("Are you sure you wish to fill the fringe layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Fringe = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If
    
    ' Fringe layer
    If frmMirage.optFringe2.Value Then
        If MsgBox("Are you sure you wish to fill the fringe2 layer?", vbYesNo, GAME_NAME) = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).Fringe2 = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If
    End If
    
End Sub

Public Sub MapEditorClearAttribs()
Dim X As Long
Dim Y As Long
    
    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Type = 0
            Next
        Next
    End If
End Sub

Public Sub MapEditorLeaveMap()
     If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If
End Sub

' /////////////////
' // Item Editor //
' /////////////////

Public Sub ItemEditorInit()
  
    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).Name)
    frmItemEditor.scrlPic.Value = Item(EditorIndex).Pic
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraSpell.Visible = False
    End If

    Call ItemEditorBltItem
    
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).Name = frmItemEditor.txtName.Text
    Item(EditorIndex).Pic = frmItemEditor.scrlPic.Value
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) Then
        If (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
            Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
            Item(EditorIndex).Data3 = 0
        End If
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) Then
        If (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
            Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.Value
            Item(EditorIndex).Data2 = 0
            Item(EditorIndex).Data3 = 0
        End If
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    Call SendSaveItem(EditorIndex)
    Editor = 0
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    Editor = 0
    Unload frmItemEditor
End Sub

' ////////////////
' // Npc Editor //
' ////////////////

Public Sub NpcEditorInit()
    frmNpcEditor.txtName.Text = Trim$(Npc(EditorIndex).Name)
    frmNpcEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    frmNpcEditor.txtChance.Text = CStr(Npc(EditorIndex).DropChance)
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).DropItem
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).DropItemValue
    frmNpcEditor.scrlStrength.Value = Npc(EditorIndex).Stat(Stats.Strength)
    frmNpcEditor.scrlDefense.Value = Npc(EditorIndex).Stat(Stats.Defense)
    frmNpcEditor.scrlSpeed.Value = Npc(EditorIndex).Stat(Stats.SPEED)
    frmNpcEditor.scrlMagic.Value = Npc(EditorIndex).Stat(Stats.Magic)

    Call NpcEditorBltSprite
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).Name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.Value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    Npc(EditorIndex).DropChance = Val(frmNpcEditor.txtChance.Text)
    Npc(EditorIndex).DropItem = frmNpcEditor.scrlNum.Value
    Npc(EditorIndex).DropItemValue = frmNpcEditor.scrlValue.Value
    Npc(EditorIndex).Stat(Stats.Strength) = frmNpcEditor.scrlStrength.Value
    Npc(EditorIndex).Stat(Stats.Defense) = frmNpcEditor.scrlDefense.Value
    Npc(EditorIndex).Stat(Stats.SPEED) = frmNpcEditor.scrlSpeed.Value
    Npc(EditorIndex).Stat(Stats.Magic) = frmNpcEditor.scrlMagic.Value
    
    Call SendSaveNpc(EditorIndex)
    Editor = 0
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    Editor = 0
    Unload frmNpcEditor
End Sub

' /////////////////
' // Shop Editor //
' /////////////////

Public Sub ShopEditorInit()
Dim i As Long

    frmShopEditor.txtName.Text = Trim$(Shop(EditorIndex).Name)
    frmShopEditor.txtJoinSay.Text = Trim$(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.Text = Trim$(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    
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
    
    Call UpdateShopTrade
    
    frmShopEditor.Show vbModal
End Sub

Public Sub UpdateShopTrade()
Dim i As Long

Dim GetItem As Long
Dim GetValue As Long
Dim GiveItem As Long
Dim GiveValue As Long
    
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
    Shop(EditorIndex).LeaveSay = frmShopEditor.txtLeaveSay.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.Value
    
    Call SendSaveShop(EditorIndex)
   Editor = 0
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    Editor = 0
    Unload frmShopEditor
End Sub

' //////////////////
' // Spell Editor //
' //////////////////

Public Sub SpellEditorInit()
Dim i As Long

    frmSpellEditor.cmbClassReq.AddItem "All Classes"
    For i = 1 To Max_Classes
        frmSpellEditor.cmbClassReq.AddItem Trim$(Class(i).Name)
    Next
    
    frmSpellEditor.txtName.Text = Trim$(Spell(EditorIndex).Name)
    frmSpellEditor.scrlPic = Spell(EditorIndex).Pic
    frmSpellEditor.scrlMPReq.Value = Spell(EditorIndex).MPReq
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.Value = Spell(EditorIndex).LevelReq
        
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    If Spell(EditorIndex).Type <> SPELL_TYPE_GIVEITEM Then
        frmSpellEditor.fraVitals.Visible = True
        frmSpellEditor.fraGiveItem.Visible = False
        frmSpellEditor.scrlVitalMod.Value = Spell(EditorIndex).Data1
    Else
        frmSpellEditor.fraVitals.Visible = False
        frmSpellEditor.fraGiveItem.Visible = True
        frmSpellEditor.scrlItemNum.Value = Spell(EditorIndex).Data1
        frmSpellEditor.scrlItemValue.Value = Spell(EditorIndex).Data2
    End If
        
    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).Name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).Pic = frmSpellEditor.scrlPic.Value
    Spell(EditorIndex).MPReq = frmSpellEditor.scrlMPReq
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.Value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    
    If Spell(EditorIndex).Type <> SPELL_TYPE_GIVEITEM Then
        Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.Value
    Else
        Spell(EditorIndex).Data1 = frmSpellEditor.scrlItemNum.Value
        Spell(EditorIndex).Data2 = frmSpellEditor.scrlItemValue.Value
    End If
    
    Spell(EditorIndex).Data3 = 0
    
    Call SendSaveSpell(EditorIndex)
    Editor = 0
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    Editor = 0
    Unload frmSpellEditor
End Sub

