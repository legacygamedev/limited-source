Attribute VB_Name = "modGameEditors"
Option Explicit
Public cpEvent As EventRec
Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public EventList() As EventListRec

' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
Dim i As Long
Dim smusic() As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set the width
    frmEditor_Map.Width = 7425
    
    ' we're in the map editor
    InMapEditor = True
    
    ' show the form
    frmEditor_Map.Visible = True
    
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.max = NumTileSets
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.value = 1
    
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.value).Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.value).Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    MapEditorTileScroll
    
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"
    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).name
    Next
    
    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorProperties()
Dim x As Long
Dim y As Long
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    For i = 1 To UBound(musicCache)
        frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
    Next
    frmEditor_MapProperties.cmbSound.Clear
    frmEditor_MapProperties.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_MapProperties.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_MapProperties
        .txtName.text = Trim$(Map.name)
        
        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For i = 0 To .lstMusic.ListCount
                If .lstMusic.List(i) = Trim$(Map.Music) Then
                    .lstMusic.ListIndex = i
                End If
            Next
        End If
        
        If .cmbSound.ListCount >= 0 Then
            .cmbSound.ListIndex = 0
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Map.BGS) Then
                    .cmbSound.ListIndex = i
                End If
            Next
        End If
        
        ' rest of it
        .txtUp.text = CStr(Map.Up)
        .txtDown.text = CStr(Map.Down)
        .txtLeft.text = CStr(Map.Left)
        .txtRight.text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral
        .txtBootMap.text = CStr(Map.BootMap)
        .txtBootX.text = CStr(Map.BootX)
        .txtBootY.text = CStr(Map.BootY)
        
        .CmbWeather.ListIndex = Map.Weather
        .scrlWeatherIntensity.value = Map.WeatherIntensity
        
        .ScrlFog.value = Map.Fog
        .ScrlFogSpeed.value = Map.FogSpeed
        .scrlFogOpacity.value = Map.FogOpacity
        
        .ScrlR.value = Map.Red
        .ScrlG.value = Map.Green
        .ScrlB.value = Map.Blue
        .scrlA.value = Map.Alpha

        ' show the map npcs
        .lstNpcs.Clear
        For x = 1 To MAX_MAP_NPCS
            If Map.Npc(x) > 0 Then
            .lstNpcs.AddItem x & ": " & Trim$(Npc(Map.Npc(x)).name)
            Else
                .lstNpcs.AddItem x & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
        
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For x = 1 To MAX_NPCS
            .cmbNpc.AddItem x & ": " & Trim$(Npc(x).name)
        Next
        
        ' set the combo box properly
        Dim tmpString() As String
        Dim npcNum As Long
        tmpString = Split(.lstNpcs.List(.lstNpcs.ListIndex))
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.Npc(npcNum)
    
        ' show the current map
        .lblMap.Caption = "Current map: " & GetPlayerMap(MyIndex)
        .txtMaxX.text = Map.MaxX
        .txtMaxY.text = Map.MaxY
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorProperties", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSetTile(ByVal x As Long, ByVal y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False, Optional ByVal theAutotile As Byte = 0)
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If theAutotile > 0 Then
        With Map.Tile(x, y)
            ' set layer
            .Layer(CurLayer).x = EditorTileX
            .Layer(CurLayer).y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
            .Autotile(CurLayer) = theAutotile
            CacheRenderState x, y, CurLayer
        End With
        ' do a re-init so we can see our changes
        initAutotiles
        Exit Sub
    End If

    If Not multitile Then ' single
        With Map.Tile(x, y)
            ' set layer
            .Layer(CurLayer).x = EditorTileX
            .Layer(CurLayer).y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
            .Autotile(CurLayer) = 0
            CacheRenderState x, y, CurLayer
        End With
    Else ' multitile
        y2 = 0 ' starting tile for y axis
        For y = CurY To CurY + EditorTileHeight - 1
            x2 = 0 ' re-set x count every y loop
            For x = CurX To CurX + EditorTileWidth - 1
                If x >= 0 And x <= Map.MaxX Then
                    If y >= 0 And y <= Map.MaxY Then
                        With Map.Tile(x, y)
                            .Layer(CurLayer).x = EditorTileX + x2
                            .Layer(CurLayer).y = EditorTileY + y2
                            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
                            .Autotile(CurLayer) = 0
                            CacheRenderState x, y, CurLayer
                        End With
                    End If
                End If
                x2 = x2 + 1
            Next
            y2 = y2 + 1
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal x As Long, ByVal y As Long, Optional ByVal movedMouse As Boolean = True)
Dim i As Long
Dim CurLayer As Long
Dim tmpDir As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).value Then
            CurLayer = i
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.value Then
            ' no autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.value
            Else ' multi tile!
                If frmEditor_Map.scrlAutotile.value = 0 Then
                    MapEditorSetTile CurX, CurY, CurLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.value
                End If
            End If
        ElseIf frmEditor_Map.optEvent.value Then
            If frmEditor_Events.Visible = False Then
                AddEvent CurX, CurY
            End If
        ElseIf frmEditor_Map.optAttribs.value Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = ""
                End If
                ' item spawn
                If frmEditor_Map.optItem.value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' key
                If frmEditor_Map.optKey.value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' key open
                If frmEditor_Map.optKeyOpen.value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' resource
                If frmEditor_Map.optResource.value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' door
                If frmEditor_Map.optDoor.value Then
                    .Type = TILE_TYPE_DOOR
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = ""
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' shop
                If frmEditor_Map.optShop.value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' bank
                If frmEditor_Map.optBank.value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' heal
                If frmEditor_Map.optHeal.value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' trap
                If frmEditor_Map.optTrap.value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' slide
                If frmEditor_Map.optSlide.value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' sound
                If frmEditor_Map.optSound.value Then
                    .Type = TILE_TYPE_SOUND
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = MapEditorSound
                End If
            End With
        ElseIf frmEditor_Map.optBlock.value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            x = x - ((x \ 32) * 32)
            y = y - ((y \ 32) * 32)
            ' see if it hits an arrow
            For i = 1 To 4
                If x >= DirArrowX(i) And x <= DirArrowX(i) + 8 Then
                    If y >= DirArrowY(i) And y <= DirArrowY(i) + 8 Then
                        ' flip the value.
                        setDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(i))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.value Then
            With Map.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).x = 0
                .Layer(CurLayer).y = 0
                .Layer(CurLayer).Tileset = 0
                If .Autotile(CurLayer) > 0 Then
                    .Autotile(CurLayer) = 0
                    ' do a re-init so we can see our changes
                    initAutotiles
                End If
                CacheRenderState x, y, CurLayer
            End With
        ElseIf frmEditor_Map.optEvent.value Then
            Call DeleteEvent(CurX, CurY)
        ElseIf frmEditor_Map.optAttribs.value Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With

        End If
    End If

    CacheResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorChooseTile(Button As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = x \ PIC_X
        EditorTileY = y \ PIC_Y
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorDrag(Button As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        x = (x \ PIC_X) + 1
        y = (y \ PIC_Y) + 1
        ' check it's not out of bounds
        If x < 0 Then x = 0
        If x > Tex_Tileset(frmEditor_Map.scrlTileSet.value).Width / PIC_X Then x = Tex_Tileset(frmEditor_Map.scrlTileSet.value).Width / PIC_X
        If y < 0 Then y = 0
        If y > Tex_Tileset(frmEditor_Map.scrlTileSet.value).Height / PIC_Y Then y = Tex_Tileset(frmEditor_Map.scrlTileSet.value).Height / PIC_Y
        ' find out what to set the width + height of map editor to
        If x > EditorTileX Then ' drag right
            EditorTileWidth = x - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If y > EditorTileY Then ' drag down
            EditorTileHeight = y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorTileScroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' horizontal scrolling
    If Tex_Tileset(frmEditor_Map.scrlTileSet.value).Width < frmEditor_Map.picBack.Width Then
        frmEditor_Map.scrlPictureX.Enabled = False
    Else
        frmEditor_Map.scrlPictureX.Enabled = True
    End If
    
    ' vertical scrolling
    If Tex_Tileset(frmEditor_Map.scrlTileSet.value).Height < frmEditor_Map.picBack.Height Then
        frmEditor_Map.scrlPictureY.Enabled = False
    Else
        frmEditor_Map.scrlPictureY.Enabled = True
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorTileScroll", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSend()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendMap
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSend", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorCancel()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong 1
    SendData buffer.ToArray()
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearLayer()
Dim i As Long
Dim x As Long
Dim y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).value Then
            CurLayer = i
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For x = 0 To Map.MaxX
            For y = 0 To Map.MaxY
                Map.Tile(x, y).Layer(CurLayer).x = 0
                Map.Tile(x, y).Layer(CurLayer).y = 0
                Map.Tile(x, y).Layer(CurLayer).Tileset = 0
                CacheRenderState x, y, CurLayer
            Next
        Next
        
        initAutotiles
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorFillLayer()
Dim i As Long
Dim x As Long
Dim y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).value Then
            CurLayer = i
            Exit For
        End If
    Next

    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For x = 0 To Map.MaxX
            For y = 0 To Map.MaxY
                Map.Tile(x, y).Layer(CurLayer).x = EditorTileX
                Map.Tile(x, y).Layer(CurLayer).y = EditorTileY
                Map.Tile(x, y).Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
                Map.Tile(x, y).Autotile(CurLayer) = frmEditor_Map.scrlAutotile.value
                CacheRenderState x, y, CurLayer
            Next
        Next
        
        ' now cache the positions
        initAutotiles
    End If
    

    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearAttribs()
Dim x As Long
Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, Options.Game_Name) = vbYes Then

        For x = 0 To Map.MaxX
            For y = 0 To Map.MaxY
                Map.Tile(x, y).Type = 0
            Next
        Next

    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearAttribs", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorLeaveMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorLeaveMap", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Item.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Item(EditorIndex)
        frmEditor_Item.txtName.text = Trim$(.name)
        If .Pic > frmEditor_Item.scrlPic.max Then .Pic = 0
        frmEditor_Item.scrlPic.value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.value = .Animation
        frmEditor_Item.txtDesc.text = Trim$(.Desc)
        
        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.List(i) = Trim$(.sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlDamage.value = .Data2
            frmEditor_Item.cmbTool.ListIndex = .Data3

            If .speed < 100 Then .speed = 100
            frmEditor_Item.scrlSpeed.value = .speed
            
            ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).value = .Add_Stat(i)
            Next
            
            frmEditor_Item.scrlPaperdoll = .Paperdoll
        Else
            frmEditor_Item.fraEquipment.Visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.Visible = True
            frmEditor_Item.scrlAddHp.value = .AddHP
            frmEditor_Item.scrlAddMP.value = .AddMP
            frmEditor_Item.scrlAddExp.value = .AddEXP
            frmEditor_Item.scrlCastSpell.value = .CastSpell
            frmEditor_Item.chkInstant.value = .instaCast
        Else
            frmEditor_Item.fraVitals.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.value = .Data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.value = .AccessReq
        frmEditor_Item.scrlLevelReq.value = .LevelReq
        
        ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(i).value = .Stat_Req(i)
        Next
        
        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.scrlPrice.value = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.value = .Rarity
         
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With
    
    Item_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    
    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Item()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Animation.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.text = Trim$(.name)
        
        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.List(i) = Trim$(.sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
        
        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).value = .Sprite(i)
            frmEditor_Animation.scrlFrameCount(i).value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).value = .LoopCount(i)
            
            If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).value = 45
            End If
            
        Next
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With
    
    Animation_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If
    Next
    
    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Animation()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_NPC.Visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_NPC
        .txtName.text = Trim$(Npc(EditorIndex).name)
        .txtAttackSay.text = Trim$(Npc(EditorIndex).AttackSay)
        If Npc(EditorIndex).Sprite < 0 Or Npc(EditorIndex).Sprite > .scrlSprite.max Then Npc(EditorIndex).Sprite = 0
        .scrlSprite.value = Npc(EditorIndex).Sprite
        .txtSpawnSecs.text = CStr(Npc(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = Npc(EditorIndex).Behaviour
        .scrlRange.value = Npc(EditorIndex).Range
        .txtChance.text = CStr(Npc(EditorIndex).DropChance)
        .scrlNum.value = Npc(EditorIndex).DropItem
        .scrlValue.value = Npc(EditorIndex).DropItemValue
        .txtHP.text = Npc(EditorIndex).HP
        .txtEXP.text = Npc(EditorIndex).EXP
        .txtLevel.text = Npc(EditorIndex).Level
        .txtDamage.text = Npc(EditorIndex).Damage
        .scrlAnimation.value = Npc(EditorIndex).Animation
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Npc(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).value = Npc(EditorIndex).Stat(i)
        Next
    End With
    
    NPC_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If
    Next
    
    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_NPC()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
Dim i As Long
Dim SoundSet As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Resource.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Resource
        .scrlExhaustedPic.max = NumResources
        .scrlNormalPic.max = NumResources
        .scrlAnimation.max = MAX_ANIMATIONS
        
        .txtName.text = Trim$(Resource(EditorIndex).name)
        .txtMessage.text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.value = Resource(EditorIndex).ItemReward
        .scrlTool.value = Resource(EditorIndex).ToolRequired
        .scrlHealth.value = Resource(EditorIndex).health
        .scrlRespawn.value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.value = Resource(EditorIndex).Animation
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Resource(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Resource_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If
    Next
    
    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Resource()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Shop.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    
    frmEditor_Shop.txtName.text = Trim$(Shop(EditorIndex).name)
    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.value = 100
    End If
    
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).name)
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    
    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Shop()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Spell.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.max = MAX_ANIMATIONS
        .scrlAnim.max = MAX_ANIMATIONS
        .scrlAOE.max = MAX_BYTE
        .scrlRange.max = MAX_BYTE
        .scrlMap.max = MAX_MAPS
        
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).name)
        Next
        
        If Spell(EditorIndex).ClassReq > -1 And Spell(EditorIndex).ClassReq <= Max_Classes Then
            .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        End If
        
        ' set values
        .txtName.text = Trim$(Spell(EditorIndex).name)
        .txtDesc.text = Trim$(Spell(EditorIndex).Desc)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.value = Spell(EditorIndex).MPCost
        .scrlLevel.value = Spell(EditorIndex).LevelReq
        .scrlAccess.value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.value = Spell(EditorIndex).CastTime
        .scrlCool.value = Spell(EditorIndex).CDTime
        .scrlIcon.value = Spell(EditorIndex).Icon
        .scrlMap.value = Spell(EditorIndex).Map
        .scrlX.value = Spell(EditorIndex).x
        .scrlY.value = Spell(EditorIndex).y
        .scrlDir.value = Spell(EditorIndex).Dir
        .scrlVital.value = Spell(EditorIndex).Vital
        .scrlDuration.value = Spell(EditorIndex).Duration
        .scrlInterval.value = Spell(EditorIndex).Interval
        .scrlRange.value = Spell(EditorIndex).Range
        If Spell(EditorIndex).IsAoE Then
            .chkAOE.value = 1
        Else
            .chkAOE.value = 0
        End If
        .scrlAOE.value = Spell(EditorIndex).AoE
        .scrlAnimCast.value = Spell(EditorIndex).CastAnim
        .scrlAnim.value = Spell(EditorIndex).SpellAnim
        .scrlStun.value = Spell(EditorIndex).StunDuration
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Spell(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Spell_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If
    Next
    
    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Spell()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAttributeDialogue()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Map.fraNpcSpawn.Visible = False
    frmEditor_Map.fraResource.Visible = False
    frmEditor_Map.fraMapItem.Visible = False
    frmEditor_Map.fraMapKey.Visible = False
    frmEditor_Map.fraKeyOpen.Visible = False
    frmEditor_Map.fraMapWarp.Visible = False
    frmEditor_Map.fraShop.Visible = False
    frmEditor_Map.fraSoundEffect.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAttributeDialogue", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


'Event Editor Stuffz Also includes event functions from the map editor (copy/paste/delete)

Sub CopyEvent_Map(x As Long, y As Long)
Dim count As Long, i As Long
    count = Map.EventCount
    If count = 0 Then Exit Sub
    
    For i = 1 To count
        If Map.Events(i).x = x And Map.Events(i).y = y Then
            ' copy it
            'CopyMemory ByVal VarPtr(cpEvent), ByVal VarPtr(Map.Events(i)), LenB(Map.Events(i))
            cpEvent = Map.Events(i)
            ' exit
            Exit Sub
        End If
    Next
End Sub

Sub PasteEvent_Map(x As Long, y As Long)
Dim count As Long, i As Long, eventNum As Long
    count = Map.EventCount
    
    If count > 0 Then
        For i = 1 To count
            If Map.Events(i).x = x And Map.Events(i).y = y Then
                ' already an event - paste over it
                eventNum = i
            End If
        Next
    End If
    
    ' couldn't find one - create one
    If eventNum = 0 Then
        ' increment count
        AddEvent x, y, True
        eventNum = count + 1
    End If
    
    ' copy it
    'CopyMemory ByVal VarPtr(Map.Events(eventNum)), ByVal VarPtr(cpEvent), LenB(cpEvent)
    Map.Events(eventNum) = cpEvent
    
    ' set position
    Map.Events(eventNum).x = x
    Map.Events(eventNum).y = y
End Sub

Sub DeleteEvent(x As Long, y As Long)
Dim count As Long, i As Long, lowIndex As Long
    If Not InMapEditor Then Exit Sub
    If frmEditor_Events.Visible = True Then Exit Sub
    count = Map.EventCount
    For i = 1 To count
        If Map.Events(i).x = x And Map.Events(i).y = y Then
            ' delete it
            ClearEvent i
            lowIndex = i
            Exit For
        End If
    Next
    
    ' not found anything
    If lowIndex = 0 Then Exit Sub
    
    ' move everything down an index
    For i = lowIndex To count - 1
        CopyEvent i + 1, i
    Next
    ' delete the last index
    ClearEvent count
    ' set the new count
    Map.EventCount = count - 1
End Sub

Sub AddEvent(x As Long, y As Long, Optional ByVal cancelLoad As Boolean = False)
Dim count As Long, pageCount As Long, i As Long
    count = Map.EventCount + 1
    ' make sure there's not already an event
    If count - 1 > 0 Then
        For i = 1 To count - 1
            If Map.Events(i).x = x And Map.Events(i).y = y Then
                ' already an event - edit it
                If Not cancelLoad Then EventEditorInit i
                Exit Sub
            End If
        Next
    End If
    ' increment count
    Map.EventCount = count
    ReDim Preserve Map.Events(0 To count)
    ' set the new event
    Map.Events(count).x = x
    Map.Events(count).y = y
    ' give it a new page
    pageCount = Map.Events(count).pageCount + 1
    Map.Events(count).pageCount = pageCount
    ReDim Preserve Map.Events(count).Pages(pageCount)
    ' load the editor
    If Not cancelLoad Then EventEditorInit count
End Sub

Sub ClearEvent(eventNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map.Events(eventNum)), LenB(Map.Events(eventNum)))
End Sub

Sub CopyEvent(original As Long, newone As Long)
    CopyMemory ByVal VarPtr(Map.Events(newone)), ByVal VarPtr(Map.Events(original)), LenB(Map.Events(original))
End Sub

Sub EventEditorInit(eventNum As Long)
Dim i As Long
    EditorEvent = eventNum
    ' copy the event data to the temp event
    'CopyMemory ByVal VarPtr(tmpEvent), ByVal VarPtr(Map.Events(eventNum)), LenB(Map.Events(eventNum))
    tmpEvent = Map.Events(eventNum)
    frmEditor_Events.InitEventEditorForm
    ' populate form
    With frmEditor_Events
        ' set the tabs
        .tabPages.Tabs.Clear
        For i = 1 To tmpEvent.pageCount
            .tabPages.Tabs.Add , , str(i)
        Next
        ' items
        .cmbHasItem.Clear
        .cmbHasItem.AddItem "None"
        For i = 1 To MAX_ITEMS
            .cmbHasItem.AddItem i & ": " & Trim$(Item(i).name)
        Next
        
        ' variables
        .cmbPlayerVar.Clear
        .cmbPlayerVar.AddItem "None"
        For i = 1 To MAX_VARIABLES
            .cmbPlayerVar.AddItem i & ". " & Variables(i)
        Next
        
        ' variables
        .cmbPlayerSwitch.Clear
        .cmbPlayerSwitch.AddItem "None"
        For i = 1 To MAX_SWITCHES
            .cmbPlayerSwitch.AddItem i & ". " & Switches(i)
        Next
        
        
        ' name
        .txtName.text = tmpEvent.name
        ' enable delete button
        If tmpEvent.pageCount > 1 Then
            .cmdDeletePage.Enabled = True
        Else
            .cmdDeletePage.Enabled = False
        End If
        .cmdPastePage.Enabled = False
        ' Load page 1 to start off with
        curPageNum = 1
        EventEditorLoadPage curPageNum
    End With
    ' show the editor
    frmEditor_Events.Show
End Sub

Sub EventEditorLoadPage(pageNum As Long)
    ' populate form
    With tmpEvent.Pages(pageNum)
        GraphicSelX = .GraphicX
        GraphicSelY = .GraphicY
        GraphicSelX2 = .GraphicX2
        GraphicSelY2 = .GraphicY2
        frmEditor_Events.cmbGraphic.ListIndex = .GraphicType
        
        frmEditor_Events.cmbHasItem.ListIndex = .HasItemIndex
        frmEditor_Events.cmbMoveFreq.ListIndex = .MoveFreq
        frmEditor_Events.cmbMoveSpeed.ListIndex = .MoveSpeed
        frmEditor_Events.cmbMoveType.ListIndex = .MoveType
        
        frmEditor_Events.cmbPlayerVar.ListIndex = .VariableIndex
        frmEditor_Events.cmbPlayerSwitch.ListIndex = .SwitchIndex
        frmEditor_Events.cmbSelfSwitch.ListIndex = .SelfSwitchIndex
        frmEditor_Events.cmbSelfSwitchCompare.ListIndex = .SelfSwitchCompare
        frmEditor_Events.cmbPlayerSwitchCompare.ListIndex = .SwitchCompare
        frmEditor_Events.cmbPlayerVarCompare.ListIndex = .VariableCompare
        
        
        frmEditor_Events.chkGlobal.value = tmpEvent.Global
        
        frmEditor_Events.cmbTrigger.ListIndex = .Trigger
        frmEditor_Events.chkDirFix.value = .DirFix
        frmEditor_Events.chkHasItem.value = .chkHasItem
        frmEditor_Events.chkPlayerVar.value = .chkVariable
        frmEditor_Events.chkPlayerSwitch.value = .chkSwitch
        frmEditor_Events.chkSelfSwitch.value = .chkSelfSwitch
        frmEditor_Events.chkWalkAnim.value = .WalkAnim
        frmEditor_Events.chkWalkThrough.value = .WalkThrough
        frmEditor_Events.chkShowName.value = .ShowName
        frmEditor_Events.txtPlayerVariable = .VariableCondition
        frmEditor_Events.scrlGraphic.value = .Graphic
        
        If .chkHasItem = 0 Then
            frmEditor_Events.cmbHasItem.Enabled = False
        Else
            frmEditor_Events.cmbHasItem.Enabled = True
        End If
        
        
        If .chkSelfSwitch = 0 Then
            frmEditor_Events.cmbSelfSwitch.Enabled = False
            frmEditor_Events.cmbSelfSwitchCompare.Enabled = False
        Else
            frmEditor_Events.cmbSelfSwitch.Enabled = True
            frmEditor_Events.cmbSelfSwitchCompare.Enabled = True
        End If
        
        If .chkSwitch = 0 Then
            frmEditor_Events.cmbPlayerSwitch.Enabled = False
            frmEditor_Events.cmbPlayerSwitchCompare.Enabled = False
        Else
            frmEditor_Events.cmbPlayerSwitch.Enabled = True
            frmEditor_Events.cmbPlayerSwitchCompare.Enabled = True
        End If
        
        
        If .chkVariable = 0 Then
            frmEditor_Events.cmbPlayerVar.Enabled = False
            frmEditor_Events.txtPlayerVariable.Enabled = False
            frmEditor_Events.cmbPlayerVarCompare.Enabled = False
        Else
            frmEditor_Events.cmbPlayerVar.Enabled = True
            frmEditor_Events.txtPlayerVariable.Enabled = True
            frmEditor_Events.cmbPlayerVarCompare.Enabled = True
        End If
        
        If frmEditor_Events.cmbMoveType.ListIndex = 2 Then
            frmEditor_Events.cmdMoveRoute.Enabled = True
        Else
            frmEditor_Events.cmdMoveRoute.Enabled = False
        End If
        
        frmEditor_Events.cmbPositioning.ListIndex = .Position
        
        ' show the commands
        EventListCommands
    End With
End Sub

Sub EventEditorOK()
    ' copy the event data from the temp event
    'CopyMemory ByVal VarPtr(Map.Events(EditorEvent)), ByVal VarPtr(tmpEvent), LenB(tmpEvent)
    Map.Events(EditorEvent) = tmpEvent
    ' unload the form
    Unload frmEditor_Events
End Sub

Public Sub EventListCommands()
Dim i As Long, curlist As Long, oldI As Long, x As Long, indent As String, listleftoff() As Long, conditionalstage() As Long
    frmEditor_Events.lstCommands.Clear
    If tmpEvent.Pages(curPageNum).CommandListCount > 0 Then
    ReDim listleftoff(1 To tmpEvent.Pages(curPageNum).CommandListCount)
    ReDim conditionalstage(1 To tmpEvent.Pages(curPageNum).CommandListCount)
        'Start Up at 1
        curlist = 1
        x = -1
newlist:
        For i = 1 To tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
            If listleftoff(curlist) > 0 Then
                If (tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(listleftoff(curlist)).Index = EventType.evCondition Or tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(listleftoff(curlist)).Index = EventType.evShowChoices) And conditionalstage(curlist) <> 0 Then
                    i = listleftoff(curlist)
                ElseIf listleftoff(curlist) >= i Then
                    i = listleftoff(curlist) + 1
                End If
            End If
            If i <= tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
                If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Index = EventType.evCondition Then
                    x = x + 1
                    Select Case conditionalstage(curlist)
                        Case 0
                            ReDim Preserve EventList(x)
                            EventList(x).CommandList = curlist
                            EventList(x).CommandNum = i
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Condition
                                Case 0
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] >= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] <= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] > " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] < " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] != " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data3
                                    End Select
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] == " & "True"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1) & "] == " & "False"
                                    End If
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Has Item [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1).name) & "]"
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Class Is [" & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1).name) & "]"
                                Case 4
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Knows Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1).name) & "]"
                                Case 5
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is >= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is <= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is > " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is < " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is NOT " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                    End Select
                                Case 6
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2 = 0 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [A] == " & "True"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [B] == " & "True"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [C] == " & "True"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [D] == " & "True"
                                        End Select
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data2 = 1 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [A] == " & "False"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [B] == " & "False"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [C] == " & "False"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [D] == " & "False"
                                        End Select
                                    End If
                            End Select
                            
                            indent = indent & "       "
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 1
                            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.CommandList
                            GoTo newlist
                        Case 1
                            ReDim Preserve EventList(x)
                            EventList(x).CommandList = curlist
                            EventList(x).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "Else"
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 2
                            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).ConditionalBranch.ElseCommandList
                            GoTo newlist
                        Case 2
                            ReDim Preserve EventList(x)
                            EventList(x).CommandList = curlist
                            EventList(x).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "End Branch"
                            indent = Mid(indent, 1, Len(indent) - 7)
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 0
                    End Select
                ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Index = EventType.evShowChoices Then
                    x = x + 1
                    Select Case conditionalstage(curlist)
                        Case 0
                            ReDim Preserve EventList(x)
                            EventList(x).CommandList = curlist
                            EventList(x).CommandNum = i
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Choices - Prompt: " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "..."
                            
                            indent = indent & "       "
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 1
                            GoTo newlist
                        Case 1
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text2) <> "" Then
                                ReDim Preserve EventList(x)
                                EventList(x).CommandList = curlist
                                EventList(x).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text2) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 2
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                GoTo newlist
                            Else
                                x = x - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 2
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 2
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text3) <> "" Then
                                ReDim Preserve EventList(x)
                                EventList(x).CommandList = curlist
                                EventList(x).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text3) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 3
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2
                                GoTo newlist
                            Else
                                x = x - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 3
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 3
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text4) <> "" Then
                                ReDim Preserve EventList(x)
                                EventList(x).CommandList = curlist
                                EventList(x).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text4) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 4
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                GoTo newlist
                            Else
                                x = x - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 4
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 4
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text5) <> "" Then
                                ReDim Preserve EventList(x)
                                EventList(x).CommandList = curlist
                                EventList(x).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text5) & "]"
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 5
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4
                                GoTo newlist
                            Else
                                x = x - 1
                                listleftoff(curlist) = i
                                conditionalstage(curlist) = 5
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 5
                            ReDim Preserve EventList(x)
                            EventList(x).CommandList = curlist
                            EventList(x).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "Branch End"
                            indent = Mid(indent, 1, Len(indent) - 7)
                            listleftoff(curlist) = i
                            conditionalstage(curlist) = 0
                    End Select
                Else
                    x = x + 1
                    ReDim Preserve EventList(x)
                    EventList(x).CommandList = curlist
                    EventList(x).CommandNum = i
                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Index
                        Case EventType.evAddText
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Add Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " - Chat Type: Player"
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Add Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " - Chat Type: Map"
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Add Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " - Chat Type: Global"
                            End Select
                        Case EventType.evShowText
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Text - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "..."
                        Case EventType.evPlayerVar
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] + " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] - " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] Random Between " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " and " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4
                            End Select
                        Case EventType.evPlayerSwitch
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] == True"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "] == False"
                            End If
                        Case EventType.evSelfSwitch
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                Case 0
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [A] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [A] to OFF"
                                    End If
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [B] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [B] to OFF"
                                    End If
                                Case 2
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [C] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [C] to OFF"
                                    End If
                                Case 3
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [D] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Self Switch [D] to OFF"
                                    End If
                            End Select
                        Case EventType.evExitProcess
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Exit Event Processing"
                        
                        Case EventType.evChangeItems
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Item Amount of [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "] to " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Give Player " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "(s)"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Take " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "(s) from Player."
                            End If
                            
                        Case EventType.evRestoreHP
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Restore Player HP"
                        Case EventType.evRestoreMP
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Restore Player MP"
                        Case EventType.evLevelUp
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Level Up Player"
                        Case EventType.evChangeLevel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Level to " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                        Case EventType.evChangeSkills
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Teach Player Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Remove Player Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "]"
                            End If
                        Case EventType.evChangeClass
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Class to " & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name)
                        Case EventType.evChangeSprite
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Sprite to " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                        Case EventType.evChangeSex
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Sex to Male."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Sex to Female."
                            End If
                        Case EventType.evChangePK
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player PK to No."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player PK to Yes."
                            End If
                        Case EventType.evWarpPlayer
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") while retaining direction."
                            Else
                                Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4 - 1
                                    Case DIR_UP
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing upward."
                                    Case DIR_DOWN
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing downward."
                                    Case DIR_LEFT
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing left."
                                    Case DIR_RIGHT
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & ") facing right."
                                End Select
                            End If
                        Case EventType.evSetMoveRoute
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 <= Map.EventCount Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Move Route for Event #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "]"
                            Else
                               frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Move Route for COULD NOT FIND EVENT!"
                            End If
                        Case EventType.evPlayAnimation
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "]" & " on Player"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "]" & " on Event #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & " [" & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3).name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "]" & " on Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4 & ")"
                            End If
                        Case EventType.evCustomScript
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Execute Custom Script Case: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                        Case EventType.evPlayBGM
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play BGM [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evFadeoutBGM
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Fadeout BGM"
                        Case EventType.evPlaySound
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Play Sound [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evStopSound
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Stop Sound"
                        Case EventType.evOpenBank
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Open Bank"
                        Case EventType.evOpenShop
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Open Shop [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & Trim$(Shop(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1).name) & "]"
                        Case EventType.evSetAccess
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player Access [" & frmEditor_Events.cmbSetAccess.List(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "]"
                        Case EventType.evGiveExp
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Give Player " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1 & " Experience."
                        Case EventType.evShowChatBubble
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                Case TARGET_TYPE_PLAYER
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Chat Bubble - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - On Player"
                                Case TARGET_TYPE_NPC
                                    If Map.Npc(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) <= 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Chat Bubble - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - On NPC [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & ". ]"
                                    Else
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Chat Bubble - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - On NPC [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & ". " & Trim$(Npc(Map.Npc(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2)).name) & "]"
                                    End If
                                Case TARGET_TYPE_EVENT
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Chat Bubble - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1, 1, 20) & "... - On Event [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & ". " & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2).name) & "]"
                            End Select
                        Case EventType.evLabel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Label: [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evGotoLabel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Jump to Label: [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Text1) & "]"
                        Case EventType.evSpawnNpc
                            If Map.Npc(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) <= 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Spawn NPC: [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & "]"
                            Else
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Spawn NPC: [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & ". " & Trim$(Npc(Map.Npc(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1)).name) & "]"
                            End If
                        Case EventType.evFadeIn
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Fade In"
                        Case EventType.evFadeOut
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Fade Out"
                        Case EventType.evFlashWhite
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Flash White"
                        Case EventType.evSetFog
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Fog [Fog: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " Speed: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & " Opacity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3) & "]"
                        Case EventType.evSetWeather
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1
                                Case WEATHER_TYPE_NONE
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [None]"
                                Case WEATHER_TYPE_RAIN
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Rain - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & "]"
                                Case WEATHER_TYPE_HAIL
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Hail - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & "]"
                                Case WEATHER_TYPE_SNOW
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Snow - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & "]"
                                Case WEATHER_TYPE_SANDSTORM
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Sand Storm - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & "]"
                                Case WEATHER_TYPE_STORM
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Storm - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & "]"
                            End Select
                        Case EventType.evSetTint
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Map Tint RGBA [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data2) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data3) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data4) & "]"
                        Case EventType.evWait
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Wait " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i).Data1) & " Ms"
                        Case Else
                            'Ghost
                            x = x - 1
                            If x = -1 Then
                                ReDim EventList(0)
                            Else
                                ReDim Preserve EventList(x)
                            End If
                    End Select
                End If
            End If
        Next
        If curlist > 1 Then
            x = x + 1
            ReDim Preserve EventList(x)
            EventList(x).CommandList = curlist
            EventList(x).CommandNum = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount + 1
            frmEditor_Events.lstCommands.AddItem indent & "@> "
            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList
            GoTo newlist
        End If
    End If
    
    frmEditor_Events.lstCommands.AddItem indent & "@> "
End Sub

Sub ListCommandAdd(s As String)
Static x As Long
    frmEditor_Events.lstCommands.AddItem s
    ' scrollbar
    If x < frmEditor_Events.TextWidth(s & "  ") Then
       x = frmEditor_Events.TextWidth(s & "  ")
      If frmEditor_Events.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX ' if twips change to pixels
      SendMessageByNum frmEditor_Events.lstCommands.hwnd, LB_SETHORIZONTALEXTENT, x, 0
    End If
End Sub

Sub AddCommand(Index As Long)
    Dim curlist As Long, i As Long, x As Long, curslot As Long, p As Long, oldCommandList As CommandListRec
    If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    End If
    
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        curlist = 1
    Else
        curlist = EventList(frmEditor_Events.lstCommands.ListIndex).CommandList
    End If
        
    If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist)
    End If
    
    oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
    tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount + 1
    p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
    If p <= 0 Then
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
    Else
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(1 To p)
        tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
        For i = 1 To p - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(i) = oldCommandList.Commands(i)
        Next
    End If
    
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
    Else
        i = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        If i < tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
            For x = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1 To i Step -1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(x + 1) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(x)
            Next
            curslot = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        Else
            curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        End If
    End If
    
    
    Select Case Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtAddText_Text.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlAddText_Colour.value
            If frmEditor_Events.optAddText_Map.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optAddText_Global.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optAddText_Player.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
        Case EventType.evCondition
            'This is the part where the whole entire source goes to hell :D
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 2
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.CommandList = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.ElseCommandList = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.CommandList).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.ElseCommandList).ParentList = curlist
            For i = 0 To 6
                If frmEditor_Events.optCondition_Index(i).value = True Then x = i
            Next
            
            Select Case x
                Case 0 'Player Var
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 0
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = Val(frmEditor_Events.txtCondition_PlayerVarCondition.text)
                Case 1 'Player Switch
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
                Case 2 'Has Item
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 2
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
                Case 3 'Class Is
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 3
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
                Case 4 'Learnt Skill
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 4
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
                Case 5 'Level Is
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 5
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = Val(frmEditor_Events.txtCondition_LevelAmount.text)
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
                Case 6 'Self Switch
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 6
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
            End Select
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtShowText.text
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChoicePrompt.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2 = frmEditor_Events.txtChoices(1).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3 = frmEditor_Events.txtChoices(2).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4 = frmEditor_Events.txtChoices(3).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5 = frmEditor_Events.txtChoices(4).text
            
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 4
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = tmpEvent.Pages(curPageNum).CommandListCount - 3
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = tmpEvent.Pages(curPageNum).CommandListCount - 2
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 3).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 2).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 1).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount).ParentList = curlist
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            For i = 0 To 3
                If frmEditor_Events.optVariableAction(i).value = True Then
                    Exit For
                End If
            Next
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = i
            If i = 3 Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(i).text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = Val(frmEditor_Events.txtVariableData(i + 1).text)
            Else
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(i).text)
            End If
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evExitProcess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtChangeItemsAmount.text)
        Case EventType.evRestoreHP
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evRestoreMP
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evLevelUp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evChangeLevel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeLevel.value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeSprite.value
        Case EventType.evChangeSex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            If frmEditor_Events.optChangeSexMale.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            End If
        Case EventType.evChangePK
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            If frmEditor_Events.optChangePKYes.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWPMap.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWPX.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlWPY.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.chkIgnoreMove.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.chkRepeatRoute.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlPlayAnimTileX.value
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlPlayAnimTileY.value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCustomScript.value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = musicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evFadeoutBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = soundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evStopSound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evOpenBank
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlGiveExp.value
        Case EventType.evShowChatBubble
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChatbubbleText.text
            If frmEditor_Events.optChatBubbleTarget(0).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_PLAYER
            ElseIf frmEditor_Events.optChatBubbleTarget(1).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_NPC
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            ElseIf frmEditor_Events.optChatBubbleTarget(2).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_EVENT
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            End If
        Case EventType.evLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtLabelName.text
        Case EventType.evGotoLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtGotoLabel.text
        Case EventType.evSpawnNpc
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSpawnNPC.ListIndex + 1
        Case EventType.evFadeIn
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evFadeOut
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evFlashWhite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evSetFog
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.ScrlFogData(0).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.ScrlFogData(1).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.ScrlFogData(2).value
        Case EventType.evSetWeather
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.CmbWeather.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWeatherIntensity.value
        Case EventType.evSetTint
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlMapTintData(0).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlMapTintData(1).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlMapTintData(2).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlMapTintData(3).value
        Case EventType.evWait
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWaitAmount.value
    End Select
    EventListCommands
End Sub

Public Sub EditEventCommand()
    Dim i As Long, x As Long, Z As Long, curlist As Long, curslot As Long
    i = frmEditor_Events.lstCommands.ListIndex
    If i = -1 Then Exit Sub
    
    If i > UBound(EventList) Then Exit Sub

    curlist = EventList(i).CommandList
    curslot = EventList(i).CommandNum
    
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    
    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index
        Case EventType.evAddText
            isEdit = True
            frmEditor_Events.txtAddText_Text.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.scrlAddText_Colour.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
                Case 0
                    frmEditor_Events.optAddText_Player.value = True
                Case 1
                    frmEditor_Events.optAddText_Map.value = True
                Case 2
                    frmEditor_Events.optAddText_Global.value = True
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(2).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCondition
            isEdit = True
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(7).Visible = True
            frmEditor_Events.fraCommands.Visible = False
            frmEditor_Events.ClearConditionFrame
            frmEditor_Events.optCondition_Index(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition).value = True
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition
                Case 0
                    frmEditor_Events.cmbCondition_PlayerVarIndex.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarCompare.Enabled = True
                    frmEditor_Events.txtCondition_PlayerVarCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                    frmEditor_Events.txtCondition_PlayerVarCondition.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3
                Case 1
                    frmEditor_Events.cmbCondition_PlayerSwitch.Enabled = True
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                Case 2
                    frmEditor_Events.cmbCondition_HasItem.Enabled = True
                    frmEditor_Events.cmbCondition_HasItem.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 3
                    frmEditor_Events.cmbCondition_ClassIs.Enabled = True
                    frmEditor_Events.cmbCondition_ClassIs.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 4
                    frmEditor_Events.cmbCondition_LearntSkill.Enabled = True
                    frmEditor_Events.cmbCondition_LearntSkill.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 5
                    frmEditor_Events.cmbCondition_LevelCompare.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                    frmEditor_Events.cmbCondition_LevelCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1
                Case 6
                    frmEditor_Events.cmbCondition_SelfSwitch.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
            End Select
        Case EventType.evShowText
            isEdit = True
            frmEditor_Events.txtShowText.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(0).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evShowChoices
            isEdit = True
            frmEditor_Events.txtChoicePrompt.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.txtChoices(1).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2
            frmEditor_Events.txtChoices(2).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3
            frmEditor_Events.txtChoices(3).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4
            frmEditor_Events.txtChoices(4).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(1).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayerVar
            isEdit = True
            frmEditor_Events.cmbVariable.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
                Case 0
                    frmEditor_Events.optVariableAction(0).value = True
                    frmEditor_Events.txtVariableData(0).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                Case 1
                    frmEditor_Events.optVariableAction(1).value = True
                    frmEditor_Events.txtVariableData(1).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                Case 2
                    frmEditor_Events.optVariableAction(2).value = True
                    frmEditor_Events.txtVariableData(2).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                Case 3
                    frmEditor_Events.optVariableAction(3).value = True
                    frmEditor_Events.txtVariableData(3).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                    frmEditor_Events.txtVariableData(4).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(4).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayerSwitch
            isEdit = True
            frmEditor_Events.cmbSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.cmbPlayerSwitchSet.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(5).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSelfSwitch
            isEdit = True
            frmEditor_Events.cmbSetSelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.cmbSetSelfSwitchTo.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(6).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeItems
            isEdit = True
            frmEditor_Events.cmbChangeItemIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0 Then
                frmEditor_Events.optChangeItemSet.value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1 Then
                frmEditor_Events.optChangeItemAdd.value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2 Then
                frmEditor_Events.optChangeItemRemove.value = True
            End If
            frmEditor_Events.txtChangeItemsAmount.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(10).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeLevel
            isEdit = True
            frmEditor_Events.scrlChangeLevel.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(11).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSkills
            isEdit = True
            frmEditor_Events.cmbChangeSkills.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0 Then
                frmEditor_Events.optChangeSkillsAdd.value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1 Then
                frmEditor_Events.optChangeSkillsRemove.value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(12).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeClass
            isEdit = True
            frmEditor_Events.cmbChangeClass.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(13).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSprite
            isEdit = True
            frmEditor_Events.scrlChangeSprite.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(14).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSex
            isEdit = True
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0 Then
                frmEditor_Events.optChangeSexMale.value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1 Then
                frmEditor_Events.optChangeSexFemale.value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(15).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangePK
            isEdit = True
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1 Then
                frmEditor_Events.optChangePKYes.value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0 Then
                frmEditor_Events.optChangePKNo.value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(16).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWarpPlayer
            isEdit = True
            frmEditor_Events.scrlWPMap.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlWPX.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.scrlWPY.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.cmbWarpPlayerDir.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(18).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetMoveRoute
            isEdit = True
            frmEditor_Events.fraMoveRoute.Visible = True
            frmEditor_Events.lstMoveRoute.Clear
            frmEditor_Events.cmbEvent.Clear
            ReDim ListOfEvents(0 To Map.EventCount)
            ListOfEvents(0) = EditorEvent
            frmEditor_Events.cmbEvent.AddItem "This Event"
            frmEditor_Events.cmbEvent.ListIndex = 0
            frmEditor_Events.cmbEvent.Enabled = True
            
            For i = 1 To Map.EventCount
                If i <> EditorEvent Then
                    frmEditor_Events.cmbEvent.AddItem Trim$(Map.Events(i).name)
                    x = x + 1
                    ListOfEvents(x) = i
                    If i = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 Then frmEditor_Events.cmbEvent.ListIndex = x
                End If
            Next
            
                
            IsMoveRouteCommand = True
                
            frmEditor_Events.chkIgnoreMove.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.chkRepeatRoute.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                
            TempMoveRouteCount = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount
            TempMoveRoute = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute
            
            For i = 1 To TempMoveRouteCount
                Select Case TempMoveRoute(i).Index
                    Case 1
                        frmEditor_Events.lstMoveRoute.AddItem "Move Up"
                    Case 2
                        frmEditor_Events.lstMoveRoute.AddItem "Move Down"
                    Case 3
                        frmEditor_Events.lstMoveRoute.AddItem "Move Left"
                    Case 4
                        frmEditor_Events.lstMoveRoute.AddItem "Move Right"
                    Case 5
                        frmEditor_Events.lstMoveRoute.AddItem "Move Randomly"
                    Case 6
                        frmEditor_Events.lstMoveRoute.AddItem "Move Towards Player"
                    Case 7
                        frmEditor_Events.lstMoveRoute.AddItem "Move Away From Player"
                    Case 8
                        frmEditor_Events.lstMoveRoute.AddItem "Step Forward"
                    Case 9
                        frmEditor_Events.lstMoveRoute.AddItem "Step Back"
                    Case 10
                        frmEditor_Events.lstMoveRoute.AddItem "Wait 100ms"
                    Case 11
                        frmEditor_Events.lstMoveRoute.AddItem "Wait 500ms"
                    Case 12
                        frmEditor_Events.lstMoveRoute.AddItem "Wait 1000ms"
                    Case 13
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Up"
                    Case 14
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Down"
                    Case 15
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Left"
                    Case 16
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Right"
                    Case 17
                        frmEditor_Events.lstMoveRoute.AddItem "Turn 90 Degrees To the Right"
                    Case 18
                        frmEditor_Events.lstMoveRoute.AddItem "Turn 90 Degrees To the Left"
                    Case 19
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Around 180 Degrees"
                    Case 20
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Randomly"
                    Case 21
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Towards Player"
                    Case 22
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Away from Player"
                    Case 23
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 8x Slower"
                    Case 24
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 4x Slower"
                    Case 25
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 2x Slower"
                    Case 26
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed to Normal"
                    Case 27
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 2x Faster"
                    Case 28
                        frmEditor_Events.lstMoveRoute.AddItem "Set Speed 4x Faster"
                    Case 29
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Lowest"
                    Case 30
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Lower"
                    Case 31
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Normal"
                    Case 32
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Higher"
                    Case 33
                        frmEditor_Events.lstMoveRoute.AddItem "Set Frequency Highest"
                    Case 34
                        frmEditor_Events.lstMoveRoute.AddItem "Turn On Walking Animation"
                    Case 35
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Off Walking Animation"
                    Case 36
                        frmEditor_Events.lstMoveRoute.AddItem "Turn On Fixed Direction"
                    Case 37
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Off Fixed Direction"
                    Case 38
                        frmEditor_Events.lstMoveRoute.AddItem "Turn On Walk Through"
                    Case 39
                        frmEditor_Events.lstMoveRoute.AddItem "Turn Off Walk Through"
                    Case 40
                        frmEditor_Events.lstMoveRoute.AddItem "Set Position Below Player"
                    Case 41
                        frmEditor_Events.lstMoveRoute.AddItem "Set Position at Player Level"
                    Case 42
                        frmEditor_Events.lstMoveRoute.AddItem "Set Position Above Player"
                    Case 43
                        frmEditor_Events.lstMoveRoute.AddItem "Set Graphic"
                End Select
            Next
                
            frmEditor_Events.fraMoveRoute.Width = 841
            frmEditor_Events.fraMoveRoute.Height = 609
            frmEditor_Events.fraMoveRoute.Visible = True
            
            frmEditor_Events.fraDialogue.Visible = False
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayAnimation
            isEdit = True
            frmEditor_Events.lblPlayAnimX.Visible = False
            frmEditor_Events.lblPlayAnimY.Visible = False
            frmEditor_Events.scrlPlayAnimTileX.Visible = False
            frmEditor_Events.scrlPlayAnimTileY.Visible = False
            frmEditor_Events.cmbPlayAnimEvent.Visible = False
            frmEditor_Events.cmbPlayAnim.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.cmbPlayAnimEvent.Clear
            For i = 1 To Map.EventCount
                frmEditor_Events.cmbPlayAnimEvent.AddItem i & ". " & Trim$(Map.Events(i).name)
            Next
            frmEditor_Events.cmbPlayAnimEvent.ListIndex = 0
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0 Then
                frmEditor_Events.optPlayAnimPlayer.value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1 Then
                frmEditor_Events.optPlayAnimEvent.value = True
                frmEditor_Events.cmbPlayAnimEvent.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 - 1
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2 Then
                frmEditor_Events.optPlayAnimTile.value = True
                frmEditor_Events.scrlPlayAnimTileX.max = Map.MaxX
                frmEditor_Events.scrlPlayAnimTileY.max = Map.MaxY
                frmEditor_Events.scrlPlayAnimTileX.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                frmEditor_Events.scrlPlayAnimTileY.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(20).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCustomScript
            isEdit = True
            frmEditor_Events.scrlCustomScript.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(29).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayBGM
            isEdit = True
            For i = 1 To UBound(musicCache())
                If musicCache(i) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 Then
                    frmEditor_Events.cmbPlayBGM.ListIndex = i - 1
                End If
            Next
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(25).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlaySound
            isEdit = True
            For i = 1 To UBound(soundCache())
                If soundCache(i) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 Then
                    frmEditor_Events.cmbPlaySound.ListIndex = i - 1
                End If
            Next
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(26).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evOpenShop
            isEdit = True
            frmEditor_Events.cmbOpenShop.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(21).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetAccess
            isEdit = True
            frmEditor_Events.cmbSetAccess.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(28).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evGiveExp
            isEdit = True
            frmEditor_Events.scrlGiveExp.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.lblGiveExp.Caption = "Give Exp: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(17).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evShowChatBubble
            isEdit = True
            frmEditor_Events.txtChatbubbleText.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
                Case TARGET_TYPE_PLAYER
                    frmEditor_Events.optChatBubbleTarget(0).value = True
                Case TARGET_TYPE_NPC
                    frmEditor_Events.optChatBubbleTarget(1).value = True
                Case TARGET_TYPE_EVENT
                    frmEditor_Events.optChatBubbleTarget(1).value = True
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(3).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evLabel
            isEdit = True
            frmEditor_Events.txtLabelName.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(8).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evGotoLabel
            isEdit = True
            frmEditor_Events.txtGotoLabel.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(9).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSpawnNpc
            isEdit = True
            frmEditor_Events.cmbSpawnNPC.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(19).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetFog
            isEdit = True
            frmEditor_Events.ScrlFogData(0).value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.ScrlFogData(1).value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.ScrlFogData(2).value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(22).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetWeather
            isEdit = True
            frmEditor_Events.CmbWeather.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlWeatherIntensity.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(23).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetTint
            isEdit = True
            frmEditor_Events.scrlMapTintData(0).value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlMapTintData(1).value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.scrlMapTintData(2).value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.scrlMapTintData(3).value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(24).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWait
            isEdit = True
            frmEditor_Events.scrlWaitAmount.value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(27).Visible = True
            frmEditor_Events.fraCommands.Visible = False
    End Select
End Sub

Public Sub DeleteEventCommand()
    Dim i As Long, x As Long, Z As Long, curlist As Long, curslot As Long, p As Long, oldCommandList As CommandListRec
    i = frmEditor_Events.lstCommands.ListIndex
    If i = -1 Then Exit Sub
    
    If i > UBound(EventList) Then Exit Sub
    
    curlist = EventList(i).CommandList
    curslot = EventList(i).CommandNum
    
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    
    If curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
        Else
            oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(p)
            x = 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
            For i = 1 To p + 1
                If i <> curslot Then
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(x) = oldCommandList.Commands(i)
                    x = x + 1
                End If
            Next
        End If
    Else
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
        x = 1
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
        Else
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(p)
            tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
            For i = 1 To p + 1
                If i <> curslot Then
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(x) = oldCommandList.Commands(i)
                    x = x + 1
                End If
            Next
        End If
    End If
    
    EventListCommands
    
End Sub

Public Sub ClearEventCommands()
    ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    tmpEvent.Pages(curPageNum).CommandListCount = 1
    EventListCommands
End Sub

Public Sub EditCommand()
    Dim i As Long, x As Long, Z As Long, curlist As Long, curslot As Long
    i = frmEditor_Events.lstCommands.ListIndex
    If i = -1 Then Exit Sub
    
    If i > UBound(EventList) Then Exit Sub

    curlist = EventList(i).CommandList
    curslot = EventList(i).CommandNum
    
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    
    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtAddText_Text.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlAddText_Colour.value
            If frmEditor_Events.optAddText_Player.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optAddText_Map.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optAddText_Global.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
        Case EventType.evCondition
            If frmEditor_Events.optCondition_Index(0).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 0
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = Val(frmEditor_Events.txtCondition_PlayerVarCondition.text)
            ElseIf frmEditor_Events.optCondition_Index(1).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(2).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(3).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 3
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(4).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 4
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(5).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 5
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = Val(frmEditor_Events.txtCondition_LevelAmount.text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(6).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 6
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
            End If
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtShowText.text
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChoicePrompt.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2 = frmEditor_Events.txtChoices(1).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3 = frmEditor_Events.txtChoices(2).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4 = frmEditor_Events.txtChoices(3).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5 = frmEditor_Events.txtChoices(4).text
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            For i = 0 To 3
                If frmEditor_Events.optVariableAction(i).value = True Then
                    Exit For
                End If
            Next
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = i
            If i = 3 Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(i).text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = Val(frmEditor_Events.txtVariableData(i + 1).text)
            Else
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(i).text)
            End If
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtChangeItemsAmount.text)
        Case EventType.evChangeLevel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeLevel.value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeSprite.value
        Case EventType.evChangeSex
            If frmEditor_Events.optChangeSexMale.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            End If
        Case EventType.evChangePK
            If frmEditor_Events.optChangePKYes.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWPMap.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWPX.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlWPY.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.chkIgnoreMove.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.chkRepeatRoute.value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlPlayAnimTileX.value
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlPlayAnimTileY.value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCustomScript.value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = musicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = soundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlGiveExp.value
        Case EventType.evShowChatBubble
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChatbubbleText.text
            If frmEditor_Events.optChatBubbleTarget(0).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_PLAYER
            ElseIf frmEditor_Events.optChatBubbleTarget(1).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_NPC
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            ElseIf frmEditor_Events.optChatBubbleTarget(2).value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_EVENT
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            End If
        Case EventType.evLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtLabelName.text
        Case EventType.evGotoLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtGotoLabel.text
        Case EventType.evSpawnNpc
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSpawnNPC.ListIndex + 1
        Case EventType.evSetFog
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.ScrlFogData(0).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.ScrlFogData(1).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.ScrlFogData(2).value
        Case EventType.evSetWeather
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.CmbWeather.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWeatherIntensity.value
        Case EventType.evSetTint
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlMapTintData(0).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlMapTintData(1).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlMapTintData(2).value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlMapTintData(3).value
        Case EventType.evWait
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWaitAmount.value
    End Select
    
    EventListCommands
End Sub

Sub RequestSwitchesAndVariables()
Dim i As Long, buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteLong CRequestSwitchesAndVariables
SendData buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendSwitchesAndVariables()
Dim i As Long, buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchesAndVariables
    For i = 1 To MAX_SWITCHES
        buffer.WriteString Switches(i)
    Next
    For i = 1 To MAX_VARIABLES
        buffer.WriteString Variables(i)
    Next
    SendData buffer.ToArray
Set buffer = Nothing
End Sub














