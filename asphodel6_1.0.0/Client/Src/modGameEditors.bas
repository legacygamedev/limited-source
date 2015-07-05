Attribute VB_Name = "modGameEditors"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

' ////////////////
' // Map Editor //
' ////////////////

Public Sub MapEditorInit()

    InitDDSurf "misc", DDSD_Misc, DDS_Misc
    
    frmMainGame.lblTileset.Caption = "0"
    frmMainGame.scrlTileSet.Value = 0
    frmMainGame.Width = MenuWidth_withEditor
    
    frmMainGame.Left = frmMainGame.Left - ((MenuWidth_withEditor - Default_MenuWidth) / 2)
    
    frmMainGame.scrlTileSet_Change
    
    BltMapEditor
    
    InEditor = True
    frmMainGame.picMapEditor.Visible = True
    
    DirectMusic_StopMidi
    
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer)
Dim i As Long
Dim LoopI As Long

    If Not isInBounds Then Exit Sub
    If frmMapProperties.Visible Then Exit Sub
    
    If Button = vbLeftButton Then
    
        If frmMainGame.optLayers.Value Then
            Dim X As Long
            Dim Y As Long
            
            For X = CurX To CurX + (EditorTileX2 - EditorTileX)
                For Y = CurY To CurY + (EditorTileY2 - EditorTileY)
                    If X <= MAX_MAPX Then
                        If Y <= MAX_MAPY Then
                            With Map.Tile(X, Y)
                                For i = 0 To UBound(.Layer)
                                    If frmMainGame.optLayer(i).Value Then
                                        .LayerSet(i) = frmMainGame.scrlTileSet.Value
                                        .Layer(i) = (EditorTileY + ((Y - CurY))) * TILESHEET_WIDTH(.LayerSet(i)) + (EditorTileX + ((X - CurX)))
                                        Exit For
                                    End If
                                Next
                            End With
                        End If
                    End If
                Next
            Next
        Else
            With Map.Tile(CurX, CurY)
                If MapAttribType > 0 Then
                    For LoopI = 1 To UBound(MapSpawn.Npc)
                        If MapSpawn.Npc(LoopI).X <> -1 Then
                            If MapSpawn.Npc(LoopI).X = CurX And MapSpawn.Npc(LoopI).Y = CurY Then
                                Exit For
                            End If
                        End If
                    Next
                    If LoopI = UBound(MapSpawn.Npc) + 1 Then
                        .Type = MapAttribType
                        .Data1 = MapAttribData(1)
                        .Data2 = MapAttribData(2)
                        .Data3 = MapAttribData(3)
                    Else
                        AddText "You cannot place an attribute on an NPC spawn spot! Clear the NPC spawn spot first!", Color.BrightRed
                        Exit Sub
                    End If
                End If
            End With
        End If
    End If
    
    If Button = vbRightButton Then
        If frmMainGame.optLayers.Value Then
            For i = 0 To UBound(Map.Tile(CurX, CurY).Layer)
                If frmMainGame.optLayer(i).Value Then
                    Map.Tile(CurX, CurY).Layer(i) = 0
                    Map.Tile(CurX, CurY).LayerSet(i) = 0
                    Exit For
                End If
            Next
        Else
            With Map.Tile(CurX, CurY)
                .Type = Tile_Type.None_
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With
        End If
    End If
    
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single, Shift As Integer)

    If Button = vbLeftButton Then
        If Shift > 0 Then
            GoTo Skippy
        End If
    End If
    
    If Button = vbLeftButton Then
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
        EditorTileX2 = X \ PIC_X
        EditorTileY2 = Y \ PIC_Y
        
        frmMainGame.shpSelected.Top = EditorTileY * PIC_Y
        frmMainGame.shpSelected.Left = EditorTileX * PIC_Y
        frmMainGame.shpSelected.Width = PIC_X
        frmMainGame.shpSelected.Height = PIC_Y
        
        frmMainGame.cmdFill.Enabled = True
    Else
Skippy:
        EditorTileX2 = X \ PIC_X
        EditorTileY2 = Y \ PIC_Y
        
        If EditorTileX2 < EditorTileX Then EditorTileX2 = EditorTileX
        If EditorTileY2 < EditorTileY Then EditorTileY2 = EditorTileY
        
        frmMainGame.shpSelected.Width = (EditorTileX2 * PIC_Y) - (EditorTileX * PIC_Y) + PIC_X
        frmMainGame.shpSelected.Height = (EditorTileY2 * PIC_Y) - (EditorTileY * PIC_Y) + PIC_Y
        
        If EditorTileX2 > EditorTileX Or EditorTileY2 > EditorTileY Then frmMainGame.cmdFill.Enabled = False Else frmMainGame.cmdFill.Enabled = True
    End If
    
End Sub

Public Sub MapEditorTileScroll()
    frmMainGame.picBackSelect.Top = -(frmMainGame.scrlPicture.Value * PIC_Y)
End Sub

Public Sub MapEditorTileScrollRight()
    frmMainGame.picBackSelect.Left = -(frmMainGame.scrlRight.Value * PIC_Y)
End Sub

Public Sub MapEditorSend()
    SendMap
    MapEditorCancel
End Sub

Public Sub MapEditorCancel()

    LoadMap GetPlayerMap(MyIndex)
    InEditor = False
    frmMainGame.picMapEditor.Visible = False
    Set DDS_Misc = Nothing
    
    frmMainGame.Width = Default_MenuWidth
    frmMainGame.Left = frmMainGame.Left + ((MenuWidth_withEditor - Default_MenuWidth) / 2)
    
    If frmAttrib.Visible Then Unload frmAttrib: ClearMapAttribs
    
End Sub

Public Sub MapEditorClearLayer()
Dim X As Long
Dim Y As Long
Dim i As Long

    For i = 0 To UBound(Map.Tile(0, 0).Layer)
        If frmMainGame.optLayer(i).Value Then
            If MsgBox("Are you sure you wish to clear the map of layer " & i & "?", vbYesNo, Game_Name) = vbYes Then
                For X = 0 To MAX_MAPX
                    For Y = 0 To MAX_MAPY
                        Map.Tile(X, Y).Layer(i) = 0
                        Map.Tile(X, Y).LayerSet(i) = 0
                    Next
                Next
                Exit Sub
            End If
        End If
    Next
    
End Sub

Public Sub MapEditorFillLayer()
Dim X As Long
Dim Y As Long
Dim i As Long

    For i = 0 To UBound(Map.Tile(0, 0).Layer)
        If frmMainGame.optLayer(i).Value Then
            If MsgBox("Are you sure you wish to fill the map with layer " & i & "?", vbYesNo, Game_Name) = vbYes Then
                For X = 0 To MAX_MAPX
                    For Y = 0 To MAX_MAPY
                        Map.Tile(X, Y).Layer(i) = EditorTileY * TILESHEET_WIDTH(frmMainGame.scrlTileSet) + EditorTileX
                        Map.Tile(X, Y).LayerSet(i) = frmMainGame.scrlTileSet
                    Next
                Next
                Exit Sub
            End If
        End If
    Next
    
End Sub

Public Sub MapEditorFillAttribs()
Dim X As Long
Dim Y As Long

    ' Block attribute
    If MapAttribType > 0 Then
        If MsgBox("Are you sure you wish to fill the map with attribute " & MapAttribType & "?", vbYesNo, Game_Name) = vbYes Then
            For X = 0 To MAX_MAPX
                For Y = 0 To MAX_MAPY
                    Map.Tile(X, Y).Type = MapAttribType
                    Map.Tile(X, Y).Data1 = MapAttribData(1)
                    Map.Tile(X, Y).Data2 = MapAttribData(2)
                    Map.Tile(X, Y).Data3 = MapAttribData(3)
                Next
            Next
            Exit Sub
        End If
    End If
    
End Sub

Public Sub MapEditorClearAttribs()
Dim X As Long
Dim Y As Long
    
    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, Game_Name) = vbYes Then
        For X = 0 To MAX_MAPX
            For Y = 0 To MAX_MAPY
                Map.Tile(X, Y).Type = 0
                Map.Tile(X, Y).Data1 = 0
                Map.Tile(X, Y).Data2 = 0
                Map.Tile(X, Y).Data3 = 0
            Next
        Next
    End If
    
End Sub

Public Sub MapEditorLeaveMap()
     If InEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            MapEditorSend
        Else
            MapEditorCancel
        End If
    End If
End Sub

' /////////////////
' // Item Editor //
' /////////////////

Public Sub ItemEditorInit()
Dim LoopI As Long

    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).Name)
    frmItemEditor.scrlPic.Value = Item(EditorIndex).Pic
    
    If Item(EditorIndex).Type > frmItemEditor.cmbType.ListCount Then Item(EditorIndex).Type = 0
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    
    If Item(EditorIndex).CostAmount > frmItemEditor.scrlAmount.Max Then Item(EditorIndex).CostAmount = frmItemEditor.scrlAmount.Max
    If Item(EditorIndex).CostItem > MAX_ITEMS Then Item(EditorIndex).CostItem = MAX_ITEMS
    
    frmItemEditor.scrlAmount.Value = Item(EditorIndex).CostAmount
    frmItemEditor.scrlWorthItem.Value = Item(EditorIndex).CostItem
    
    For LoopI = 1 To UBound(Class)
        frmItemEditor.cmbClass.AddItem Trim$(Class(LoopI).Name)
    Next
    
    frmItemEditor.scrlAnim.Value = Item(EditorIndex).Anim
    If Item(EditorIndex).Required(Item_Requires.Class_) > frmItemEditor.cmbClass.ListCount Then Item(EditorIndex).Required(Item_Requires.Class_) = frmItemEditor.cmbClass.ListCount
    frmItemEditor.cmbClass.ListIndex = Item(EditorIndex).Required(Item_Requires.Class_)
    
    If Item(EditorIndex).Required(Item_Requires.Level_) < 1 Then Item(EditorIndex).Required(Item_Requires.Level_) = 1
    frmItemEditor.scrlLevel.Value = Item(EditorIndex).Required(Item_Requires.Level_)
    
    If Item(EditorIndex).Required(Item_Requires.Access_) > 4 Then Item(EditorIndex).Required(Item_Requires.Access_) = 4
    frmItemEditor.scrlAccess.Value = Item(EditorIndex).Required(Item_Requires.Access_)
    
    For LoopI = 0 To 3
        frmItemEditor.scrlRequires(LoopI).Value = Item(EditorIndex).Required(LoopI)
    Next
    
    For LoopI = 1 To Stats.Stat_Count - 1
        frmItemEditor.scrlBonusStat(LoopI).Value = Item(EditorIndex).BuffStats(LoopI)
    Next
    
    For LoopI = 1 To Vitals.Vital_Count - 1
        frmItemEditor.scrlBonusVital(LoopI).Value = Item(EditorIndex).BuffVitals(LoopI)
    Next
    
    If (frmItemEditor.cmbType.ListIndex >= ItemType.Weapon_) And (frmItemEditor.cmbType.ListIndex <= ItemType.Shield_) Then
        frmItemEditor.fraEquipment.Visible = True
        If Item(EditorIndex).Durability < 0 Then Item(EditorIndex).Durability = 0
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Durability
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ItemType.Potion) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlVitalMod2.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlVitalMod3.Value = Item(EditorIndex).Data3
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ItemType.Spell_) Then
        frmItemEditor.fraSpell.Visible = True
        If Item(EditorIndex).Data1 < 1 Then Item(EditorIndex).Data1 = 1
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
        frmItemEditor.lblSpellName.Caption = Item(EditorIndex).Name
    Else
        frmItemEditor.fraSpell.Visible = False
    End If
    
    frmItemEditor.scrlPic.Max = MAX_ITEMSETS
    
    ItemEditorBltItem
    
    frmItemEditor.Show vbModal
    
End Sub

Public Sub ItemEditorOk()
Dim LoopI As Long

    Item(EditorIndex).Name = frmItemEditor.txtName.Text
    Item(EditorIndex).Pic = frmItemEditor.scrlPic.Value
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
    
    Item(EditorIndex).CostItem = frmItemEditor.scrlWorthItem.Value
    Item(EditorIndex).CostAmount = frmItemEditor.scrlAmount.Value
    
    Item(EditorIndex).Required(Item_Requires.Class_) = frmItemEditor.cmbClass.ListIndex
    Item(EditorIndex).Required(Item_Requires.Level_) = frmItemEditor.scrlLevel.Value
    Item(EditorIndex).Required(Item_Requires.Access_) = frmItemEditor.scrlAccess.Value
    
    Item(EditorIndex).Anim = frmItemEditor.scrlAnim.Value
    
    For LoopI = 0 To 3
        Item(EditorIndex).Required(LoopI) = frmItemEditor.scrlRequires(LoopI).Value
    Next
    
    If (frmItemEditor.cmbType.ListIndex >= ItemType.Weapon_) Then
        If (frmItemEditor.cmbType.ListIndex <= ItemType.Shield_) Then
            Item(EditorIndex).Durability = frmItemEditor.scrlDurability.Value
            Item(EditorIndex).Data1 = 0
            Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
            Item(EditorIndex).Data3 = 0
            For LoopI = 1 To Stats.Stat_Count - 1
                Item(EditorIndex).BuffStats(LoopI) = frmItemEditor.scrlBonusStat(LoopI).Value
            Next
            For LoopI = 1 To Vitals.Vital_Count - 1
                Item(EditorIndex).BuffVitals(LoopI) = frmItemEditor.scrlBonusVital(LoopI).Value
            Next
        End If
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ItemType.Potion) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlVitalMod2.Value
        Item(EditorIndex).Data3 = frmItemEditor.scrlVitalMod3.Value
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ItemType.Spell_) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    SendSaveItem EditorIndex
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
Dim i As Long
Dim ii As Long
Dim SoundName As String

    frmNpcEditor.txtName.Text = Trim$(Npc(EditorIndex).Name)
    frmNpcEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    If Npc(EditorIndex).DropChance < 1 Then Npc(EditorIndex).DropChance = 1
    frmNpcEditor.scrlChance.Value = Npc(EditorIndex).DropChance
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).DropItem
    
    frmNpcEditor.chkGivesGuild.Value = Npc(EditorIndex).GivesGuild
    
    frmNpcEditor.flSound.Path = App.Path & SOUND_PATH
    frmNpcEditor.lstSoundTypes.ListIndex = 0
    frmNpcEditor.flSound.ListIndex = -1
    
    For ii = 0 To UBound(Npc(EditorIndex).Sound)
        Select Case ii
            Case NpcSound.Attack_
                SoundName = "Attack "
            Case NpcSound.Spawn_
                SoundName = "On Spawn "
            Case NpcSound.Death_
                SoundName = "On Death "
        End Select
        If LenB(Trim$(Npc(EditorIndex).Sound(ii))) > 0 Then
            If ii = NpcSound.Attack_ Then frmNpcEditor.lblCurrentSound.Caption = SoundName & "Sound: " & Trim$(Npc(EditorIndex).Sound(ii)) & SOUND_EXT
            Select Case ii
                Case NpcSound.Attack_
                    EditorNpcAttackSound = Trim$(Npc(EditorIndex).Sound(ii))
                Case NpcSound.Spawn_
                    EditorNpcSpawnSound = Trim$(Npc(EditorIndex).Sound(ii))
                Case NpcSound.Death_
                    EditorNpcDeathSound = Trim$(Npc(EditorIndex).Sound(ii))
            End Select
            If ii = NpcSound.Attack_ Then
                If frmNpcEditor.flSound.ListCount > 0 Then
                    For i = 0 To frmNpcEditor.flSound.ListCount
                        If frmNpcEditor.flSound.List(i) = Trim$(Npc(EditorIndex).Sound(ii)) & SOUND_EXT Then
                            frmNpcEditor.flSound.Selected(i) = True
                            frmNpcEditor.flSound.ListIndex = i
                        End If
                    Next
                End If
            End If
        Else
            If ii = NpcSound.Attack_ Then frmNpcEditor.lblCurrentSound.Caption = SoundName & "Sound: None"
            Select Case ii
                Case NpcSound.Attack_
                    EditorNpcAttackSound = vbNullString
                Case NpcSound.Spawn_
                    EditorNpcSpawnSound = vbNullString
                Case NpcSound.Death_
                    EditorNpcDeathSound = vbNullString
            End Select
        End If
    Next
    
    For i = 0 To UBound(Npc(EditorIndex).Reflection)
        If Npc(EditorIndex).Reflection(i) > 0 Then
            frmNpcEditor.chkReflection(i).Value = 1
            frmNpcEditor.scrlReflection(i).Value = Npc(EditorIndex).Reflection(i)
        End If
    Next
    
    If frmNpcEditor.scrlNum.Value < 1 Then
        frmNpcEditor.lblItemName.Caption = "None"
    Else
        frmNpcEditor.lblItemName.Caption = Item(frmNpcEditor.scrlNum.Value).Name
    End If
    
    If Npc(EditorIndex).HP < 1 Then Npc(EditorIndex).HP = 1
    frmNpcEditor.scrlHP.Value = Npc(EditorIndex).HP
    frmNpcEditor.scrlExp.Value = Npc(EditorIndex).Experience
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).DropItemValue
    frmNpcEditor.scrlStrength.Value = Npc(EditorIndex).Stat(Stats.Strength)
    frmNpcEditor.scrlDefense.Value = Npc(EditorIndex).Stat(Stats.Defense)
    frmNpcEditor.scrlSpeed.Value = Npc(EditorIndex).Stat(Stats.Speed)
    frmNpcEditor.scrlMagic.Value = Npc(EditorIndex).Stat(Stats.Magic)
    
    frmNpcEditor.scrlSprite.Max = TOTAL_SPRITES
    
    NpcEditorBltSprite
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
Dim LoopI As Long

    With Npc(EditorIndex)
        .HP = frmNpcEditor.scrlHP.Value
        .Experience = frmNpcEditor.scrlExp.Value
        .Name = frmNpcEditor.txtName.Text
        .AttackSay = frmNpcEditor.txtAttackSay.Text
        .Sprite = frmNpcEditor.scrlSprite.Value
        .SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
        .Behavior = frmNpcEditor.cmbBehavior.ListIndex
        .Range = frmNpcEditor.scrlRange.Value
        .DropChance = frmNpcEditor.scrlChance.Value
        .DropItem = frmNpcEditor.scrlNum.Value
        .DropItemValue = frmNpcEditor.scrlValue.Value
        .Stat(Stats.Strength) = frmNpcEditor.scrlStrength.Value
        .Stat(Stats.Defense) = frmNpcEditor.scrlDefense.Value
        .Stat(Stats.Speed) = frmNpcEditor.scrlSpeed.Value
        .Stat(Stats.Magic) = frmNpcEditor.scrlMagic.Value
        .Sound(NpcSound.Attack_) = EditorNpcAttackSound
        .Sound(NpcSound.Spawn_) = EditorNpcSpawnSound
        .Sound(NpcSound.Death_) = EditorNpcDeathSound
        .GivesGuild = frmNpcEditor.chkGivesGuild.Value
        For LoopI = 0 To UBound(.Reflection)
            If frmNpcEditor.chkReflection(LoopI).Value = 1 Then
                .Reflection(LoopI) = frmNpcEditor.scrlReflection(LoopI).Value
            Else
                .Reflection(LoopI) = 0
            End If
        Next
    End With
    
    SendSaveNpc EditorIndex
    
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
        
        UpdateShopTrade
        
        .Show vbModal
    End With
    
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
    
    SendSaveShop EditorIndex
    Editor = 0
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    Editor = 0
    Unload frmShopEditor
End Sub

' //////////////////
' // Anim Editor  //
' //////////////////

Public Sub AnimEditorInit()
Dim i As Long

    frmAnimEditor.txtName.Text = Trim$(Animation(EditorIndex).Name)
    
    If Animation(EditorIndex).Height < 1 Then Animation(EditorIndex).Height = 32
    frmAnimEditor.txtSizeY.Text = Animation(EditorIndex).Height
    If Animation(EditorIndex).Width < 1 Then Animation(EditorIndex).Width = 32
    frmAnimEditor.txtSizeX.Text = Animation(EditorIndex).Width
    If Animation(EditorIndex).Delay < 1 Then Animation(EditorIndex).Delay = 100
    frmAnimEditor.scrlDelay.Value = Animation(EditorIndex).Delay
    If Animation(EditorIndex).Pic < 1 Then Animation(EditorIndex).Pic = 1
    frmAnimEditor.scrlPic.Value = Animation(EditorIndex).Pic
    
    frmAnimEditor.Show vbModal
    
End Sub

Public Sub AnimEditorOk()
Dim LoopI As Long

    With Animation(EditorIndex)
        .Name = frmAnimEditor.txtName.Text
        .Delay = frmAnimEditor.scrlDelay.Value
        .Height = Val(frmAnimEditor.txtSizeY.Text)
        .Width = Val(frmAnimEditor.txtSizeX.Text)
        .Pic = frmAnimEditor.scrlPic.Value
    End With
    
    SendSaveAnim EditorIndex
    AnimEditorCancel
    
End Sub

Public Sub AnimEditorCancel()
    Editor = 0
    Unload frmAnimEditor
End Sub

' //////////////////
' // Sign Editor  //
' //////////////////

Public Sub SignEditorInit()
Dim i As Long

    frmSignEditor.txtName.Text = Trim$(Sign(EditorIndex).Name)
    
    frmSignEditor.lstSections.Clear
    
    For i = 0 To UBound(Sign(EditorIndex).Section)
        frmSignEditor.lstSections.AddItem "Section " & i
        SignSection(i) = Sign(EditorIndex).Section(i)
    Next
    
    frmSignEditor.lstSections.ListIndex = 0
    
    frmSignEditor.txtSign.Text = SignSection(0)
    
    frmSignEditor.Show vbModal
    
End Sub

Public Sub SignEditorOk()
Dim LoopI As Long

    Sign(EditorIndex).Name = frmSignEditor.txtName.Text
    
    For LoopI = 0 To UBound(Sign(EditorIndex).Section)
        Sign(EditorIndex).Section(LoopI) = SignSection(LoopI)
    Next
    
    SendSaveSign EditorIndex
    SignEditorCancel
    
End Sub

Public Sub SignEditorCancel()
    Editor = 0
    Unload frmSignEditor
End Sub

' //////////////////
' // Spell Editor //
' //////////////////

Public Sub SpellEditorInit()
Dim i As Long

    If Spell(EditorIndex).Timer < 1 Then Spell(EditorIndex).Timer = 1000
    frmSpellEditor.scrlSpeed.Value = Spell(EditorIndex).Timer
    
    frmSpellEditor.flSound.Path = App.Path & SOUND_PATH
    frmSpellEditor.flSound.ListIndex = -1
    
    If LenB(Trim$(Spell(EditorIndex).CastSound)) > 0 Then
        frmSpellEditor.lblCurrentSound.Caption = "Cast Sound: " & Trim$(Spell(EditorIndex).CastSound) & SOUND_EXT
        EditorSpellSound = Trim$(Spell(EditorIndex).CastSound)
        If frmSpellEditor.flSound.ListCount > 0 Then
            For i = 0 To frmSpellEditor.flSound.ListCount
                If frmSpellEditor.flSound.List(i) = Trim$(Spell(EditorIndex).CastSound) & SOUND_EXT Then
                    frmSpellEditor.flSound.Selected(i) = True
                    frmSpellEditor.flSound.ListIndex = i
                End If
            Next
        End If
    Else
        frmSpellEditor.lblCurrentSound.Caption = "Cast Sound: None"
    End If
    
    frmSpellEditor.txtName.Text = Trim$(Spell(EditorIndex).Name)
    frmSpellEditor.scrlMPReq.Value = Spell(EditorIndex).MPReq
    If Spell(EditorIndex).Anim < 1 Then Spell(EditorIndex).Anim = 1
    frmSpellEditor.scrlPic.Value = Spell(EditorIndex).Anim
    frmSpellEditor.scrlRange.Value = Spell(EditorIndex).Range
    frmSpellEditor.scrlIcon.Value = Spell(EditorIndex).Icon
    frmSpellEditor.chkAOE.Value = Spell(EditorIndex).AOE
    
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    If Spell(EditorIndex).Type <> Spell_Type.GiveItem Then
        frmSpellEditor.fraVitals.Visible = True
        frmSpellEditor.fraGiveItem.Visible = False
        frmSpellEditor.scrlVitalMod.Value = Spell(EditorIndex).Data1
    Else
        frmSpellEditor.fraVitals.Visible = False
        frmSpellEditor.fraGiveItem.Visible = True
        frmSpellEditor.scrlItemNum.Value = Spell(EditorIndex).Data1
        frmSpellEditor.scrlItemValue.Value = Spell(EditorIndex).Data2
    End If
    
    frmSpellEditor.scrlPic.Max = MAX_ANIMS
    frmSpellEditor.scrlIcon.Max = (DDSD_SpellIcon.lHeight \ 32) - 1
    
    SpellEditorBltIcon
    
    frmSpellEditor.Show vbModal
    
End Sub

Public Sub SpellEditorOk()

    With Spell(EditorIndex)
        .Name = frmSpellEditor.txtName.Text
        .CastSound = EditorSpellSound
        .Timer = frmSpellEditor.scrlSpeed.Value
        .MPReq = frmSpellEditor.scrlMPReq.Value
        .Type = frmSpellEditor.cmbType.ListIndex
        .Anim = frmSpellEditor.scrlPic.Value
        .Range = frmSpellEditor.scrlRange.Value
        .Icon = frmSpellEditor.scrlIcon.Value
        .AOE = frmSpellEditor.chkAOE.Value
        If .Type <> Spell_Type.GiveItem Then
            .Data1 = frmSpellEditor.scrlVitalMod.Value
        Else
            .Data1 = frmSpellEditor.scrlItemNum.Value
            .Data2 = frmSpellEditor.scrlItemValue.Value
        End If
        .Data3 = 0
    End With
    
    SendSaveSpell EditorIndex
    Editor = 0
    Unload frmSpellEditor
    
End Sub

Public Sub SpellEditorCancel()
    Editor = 0
    Unload frmSpellEditor
End Sub

