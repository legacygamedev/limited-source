Attribute VB_Name = "modGameEditors"
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************

Public Editor As Byte ' which game editor is used
Public EditorIndex As Long ' index number

' selecting tiles
Public EditorTileX As Long
Public EditorTileY As Long

' map attribute data
Public EditorData1 As Long
Public EditorData2 As Long
Public EditorData3 As Long
Public ReCalcTiles As Boolean

' ////////////////
' // Map Editor //
' ////////////////

Public Sub MapEditorInit()
On Error GoTo ErrorHandle
    
    ReCalcTiles = True
    
    Editor = EDITOR_MAP
    
    EditorTileX = 0
    EditorTileY = 0
    
    EditorData1 = 0
    EditorData2 = 0
    EditorData3 = 0
    
    With frmMainGame
        .Width = 14175
        .picMapEditor.Visible = True
        .lblTileset = map(5).TileSet
        .scrlTileSet = map(5).TileSet
        .scrlPicture.Max = (.picBackSelect.Height \ PIC_Y) - (.picBack.Height \ PIC_Y)
    End With
     
    Call InitDDSurf("misc", DDS_Misc)
     
    Call BltMapEditor
    Call BltMapEditorTilePreview
    
    Exit Sub
    
ErrorHandle:

    Select Case Err
    
        Case 380
            map(5).TileSet = 1
            MapEditorInit
            
    End Select
    
End Sub

Public Sub MapEditorMouseDown(Button As Integer)
    If Not isInBounds Then Exit Sub
    
    Select Case Button
    
        Case vbLeftButton
        
        If frmMainGame.optLayers.Value Then
            
            map(5).Tile(CurX, CurY).Num(frmMainGame.HScroll1.Value) = EditorTileY * TILESHEET_WIDTH + EditorTileX
        
            Call CalcTilePositions
            
        Else
        
            With map(5).Tile(CurX, CurY)
            
                ' clear data
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            
                If frmMainGame.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                
                If frmMainGame.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorData1
                    .Data2 = EditorData2
                    .Data3 = EditorData3
                End If
                If frmMainGame.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = EditorData1
                    .Data2 = EditorData2
                    .Data3 = 0
                End If
                If frmMainGame.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
                If frmMainGame.optKey.Value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = EditorData1
                    .Data2 = EditorData2
                    .Data3 = 0
                End If
                If frmMainGame.optKeyOpen.Value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = EditorData1
                    .Data2 = EditorData2
                    .Data3 = 0
                End If
                If frmMainGame.OptHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
                If frmMainGame.optDMG.Value Then
                    .Type = TILE_TYPE_KILL
                    .Data1 = EditorData1
                    .Data2 = 0
                    .Data3 = 0
                End If
                If frmMainGame.optSpr.Value Then
                    .Type = TILE_TYPE_SPRITE
                    .Data1 = EditorData1
                    .Data2 = 0
                    .Data3 = 0
                End If
            End With
        End If
    
    Case vbRightButton
    
        If frmMainGame.optLayers.Value Then
            
            map(5).Tile(CurX, CurY).Num(frmMainGame.HScroll1.Value) = 0
        
            Call CalcTilePositions
    
        Else
            With map(5).Tile(CurX, CurY)
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With
        End If

    End Select

End Sub

Public Sub MapEditorChooseTile(ByVal x As Single, ByVal y As Single)

    EditorTileX = x \ PIC_X
    EditorTileY = y \ PIC_Y
    
    With frmMainGame
        .shpSelected.Top = EditorTileY * PIC_Y
        .shpSelected.Left = EditorTileX * PIC_Y
    End With
    
    Call BltMapEditorTilePreview
    
End Sub

Public Sub MapEditorTileScroll()
    With frmMainGame
        .picBackSelect.Top = (.scrlPicture.Value * PIC_Y) * -1
    End With
End Sub

Public Sub MapEditorSend()
    Call SendMap
    Call MapEditorCancel
End Sub

Public Sub MapEditorCancel()
    Editor = EDITOR_NONE
        
    With frmMainGame
        .Width = 10080
        .picMapEditor.Visible = False
    End With

    Call LoadMaps(GetPlayerMap(MyIndex))
    Call InitMapData
    
    Call DD_ClearBuffer(DDS_Misc)
End Sub

Public Sub MapEditorClearLayer()
Dim x As Long
Dim y As Long

        If MsgBox("Are you sure you wish to clear the " & frmMainGame.lblLayer.Caption & " layer?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    map(5).Tile(x, y).Num(frmMainGame.HScroll1.Value) = 0
                Next
            Next
        End If
    
    Call CalcTilePositions
    
End Sub

Public Sub MapEditorFillLayer()
Dim x As Long
Dim y As Long

        If MsgBox("Are you sure you wish to fill the " & frmMainGame.lblLayer.Caption & " layer?", vbYesNo, GAME_NAME) = vbYes Then
            For x = 0 To MAX_MAPX
                For y = 0 To MAX_MAPY
                    map(5).Tile(x, y).Num(frmMainGame.HScroll1.Value) = EditorTileY * TILESHEET_WIDTH + EditorTileX
                Next
            Next
        End If

    
    Call CalcTilePositions
    
End Sub

Public Sub MapEditorClearAttribs()
Dim x As Long
Dim y As Long
    
    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME) = vbYes Then
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                map(5).Tile(x, y).Type = 0
            Next
        Next
    End If
End Sub

Public Sub MapEditorLeaveMap()
     If Editor = EDITOR_MAP Then
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
    
    With frmItemEditor
        .txtName.Text = Trim$(Item(EditorIndex).Name)
        .scrlPic.Value = Item(EditorIndex).Pic
        .cmbType.ListIndex = Item(EditorIndex).Type
        
        If (.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            .fraEquipment.Visible = True
            .scrlDurability.Value = Item(EditorIndex).Data1
            .scrlStrength.Value = Item(EditorIndex).Data2
        Else
            .fraEquipment.Visible = False
        End If
        
        If (.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
            .fraVitals.Visible = True
            .scrlVitalMod.Value = Item(EditorIndex).Data1
        Else
            .fraVitals.Visible = False
        End If
        
        If (.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            .fraSpell.Visible = True
            .scrlSpell.Value = Item(EditorIndex).Data1
        Else
            .fraSpell.Visible = False
        End If
        
        Call ItemEditorBltItem
        
        .Show vbModal
    End With
End Sub

Public Sub ItemEditorOk()

    With Item(EditorIndex)
        .Data1 = 0
        .Data2 = 0
        .Data3 = 0
    
        .Name = frmItemEditor.txtName.Text
        .Pic = frmItemEditor.scrlPic.Value
        .Type = frmItemEditor.cmbType.ListIndex

        If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) Then
            If (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
                .Data1 = frmItemEditor.scrlDurability.Value
                .Data2 = frmItemEditor.scrlStrength.Value
            End If
        End If
        
        If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) Then
            If (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
                .Data1 = frmItemEditor.scrlVitalMod.Value
            End If
        End If
        
        If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            .Data1 = frmItemEditor.scrlSpell.Value
        End If
    End With
    
    Call SendSaveItem(EditorIndex)
    
    Editor = EDITOR_NONE
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    Editor = EDITOR_NONE
    Unload frmItemEditor
End Sub

' ////////////////
' // Npc Editor //
' ////////////////

Public Sub NpcEditorInit()
    With frmNpcEditor
        .txtName.Text = Trim$(Npc(EditorIndex).Name)
        .txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
        .scrlSprite.Value = Npc(EditorIndex).Sprite
        .txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
        .cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
        .scrlRange.Value = Npc(EditorIndex).Range
        .txtChance.Text = CStr(Npc(EditorIndex).DropChance)
        .scrlNum.Value = Npc(EditorIndex).DropItem
        .scrlValue.Value = Npc(EditorIndex).DropItemValue
        .scrlStrength.Value = Npc(EditorIndex).Stat(Stats.Strength)
        .scrlDefense.Value = Npc(EditorIndex).Stat(Stats.Defense)
        .scrlSpeed.Value = Npc(EditorIndex).Stat(Stats.Speed)
        .scrlMagic.Value = Npc(EditorIndex).Stat(Stats.Magic)

        .Show vbModal
    End With
    
    Call NpcEditorBltSprite
        
End Sub

Public Sub NpcEditorOk()
    With Npc(EditorIndex)
        .Name = frmNpcEditor.txtName.Text
        .AttackSay = frmNpcEditor.txtAttackSay.Text
        .Sprite = frmNpcEditor.scrlSprite.Value
        .SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
        .Behavior = frmNpcEditor.cmbBehavior.ListIndex
        .Range = frmNpcEditor.scrlRange.Value
        .DropChance = Val(frmNpcEditor.txtChance.Text)
        .DropItem = frmNpcEditor.scrlNum.Value
        .DropItemValue = frmNpcEditor.scrlValue.Value
        .Stat(Stats.Strength) = frmNpcEditor.scrlStrength.Value
        .Stat(Stats.Defense) = frmNpcEditor.scrlDefense.Value
        .Stat(Stats.Speed) = frmNpcEditor.scrlSpeed.Value
        .Stat(Stats.Magic) = frmNpcEditor.scrlMagic.Value
    End With
        
    Call SendSaveNpc(EditorIndex)
    Editor = EDITOR_NONE
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    Editor = EDITOR_NONE
    Unload frmNpcEditor
End Sub

' /////////////////
' // Shop Editor //
' /////////////////

Public Sub ShopEditorInit()
Dim i As Long
    With frmShopEditor
        .txtName.Text = Trim$(Shop(EditorIndex).Name)
        .txtJoinSay.Text = Trim$(Shop(EditorIndex).JoinSay)
        .txtLeaveSay.Text = Trim$(Shop(EditorIndex).LeaveSay)
        .chkFixesItems.Value = Shop(EditorIndex).FixesItems
        
        .cmbItemGive.Clear
        .cmbItemGive.AddItem "None"
        .cmbItemGet.Clear
        .cmbItemGet.AddItem "None"
        
        For i = 1 To MAX_ITEMS
            .cmbItemGive.AddItem i & ": " & Trim$(Item(i).Name)
            .cmbItemGet.AddItem i & ": " & Trim$(Item(i).Name)
        Next
        
        .cmbItemGive.ListIndex = 0
        .cmbItemGet.ListIndex = 0
        
        .Show vbModal
    End With
    
    Call UpdateShopTrade
    
End Sub

Public Sub UpdateShopTrade()
Dim i As Long

Dim GetItem As Long
Dim GetValue As Long
Dim GiveItem As Long
Dim GiveValue As Long
    
    frmShopEditor.lstTradeItem.Clear
    
    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            GetItem = .GetItem
            GetValue = .GetValue
            GiveItem = .GiveItem
            GiveValue = .GiveValue
        End With
        
        If GetItem > 0 And GiveItem > 0 Then
            frmShopEditor.lstTradeItem.AddItem i & ": " & GiveValue & " " & Trim$(Item(GiveItem).Name) & " for " & GetValue & " " & Trim$(Item(GetItem).Name)
        Else
            frmShopEditor.lstTradeItem.AddItem "Empty Trade Slot"
        End If
    Next
    
    frmShopEditor.lstTradeItem.ListIndex = 0
End Sub

Public Sub ShopEditorOk()
    With Shop(EditorIndex)
        .Name = frmShopEditor.txtName.Text
        .JoinSay = frmShopEditor.txtJoinSay.Text
        .LeaveSay = frmShopEditor.txtLeaveSay.Text
        .FixesItems = frmShopEditor.chkFixesItems.Value
    End With
    
    Call SendSaveShop(EditorIndex)
    Editor = EDITOR_NONE
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    Editor = EDITOR_NONE
    Unload frmShopEditor
End Sub

' //////////////////
' // Spell Editor //
' //////////////////

Public Sub SpellEditorInit()
Dim i As Long

    With frmSpellEditor
        .cmbClassReq.AddItem "All Classes"
        For i = 1 To Max_Classes
            .cmbClassReq.AddItem Trim$(Class(i).Name)
        Next
        
        .txtName.Text = Trim$(Spell(EditorIndex).Name)
        .scrlPic = Spell(EditorIndex).Pic
        .scrlMPReq.Value = Spell(EditorIndex).MPReq
        .cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
        .scrlLevelReq.Value = Spell(EditorIndex).LevelReq
            
        .cmbType.ListIndex = Spell(EditorIndex).Type
        If Spell(EditorIndex).Type <> SPELL_TYPE_GIVEITEM Then
            .fraVitals.Visible = True
            .fraGiveItem.Visible = False
            .scrlVitalMod.Value = Spell(EditorIndex).Data1
        Else
            .fraVitals.Visible = False
            .fraGiveItem.Visible = True
            .scrlItemNum.Value = Spell(EditorIndex).Data1
            .scrlItemValue.Value = Spell(EditorIndex).Data2
        End If
            
        .Show vbModal
    End With
End Sub

Public Sub SpellEditorOk()
    With Spell(EditorIndex)
        .Name = frmSpellEditor.txtName.Text
        .Pic = frmSpellEditor.scrlPic.Value
        .MPReq = frmSpellEditor.scrlMPReq
        .ClassReq = frmSpellEditor.cmbClassReq.ListIndex
        .LevelReq = frmSpellEditor.scrlLevelReq.Value
        .Type = frmSpellEditor.cmbType.ListIndex
        
        If .Type <> SPELL_TYPE_GIVEITEM Then
            .Data1 = frmSpellEditor.scrlVitalMod.Value
        Else
            .Data1 = frmSpellEditor.scrlItemNum.Value
            .Data2 = frmSpellEditor.scrlItemValue.Value
        End If
        
        .Data3 = 0
    End With
    
    Call SendSaveSpell(EditorIndex)
    Editor = EDITOR_NONE
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    Editor = EDITOR_NONE
    Unload frmSpellEditor
End Sub

