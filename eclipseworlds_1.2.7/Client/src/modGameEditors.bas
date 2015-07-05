Attribute VB_Name = "modGameEditors"
Option Explicit

Public cpEvent As EventRec

Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public charList() As String
Public g_playersOnline() As String
Public ignoreIndexes() As Long
Public refreshingAdminList As Boolean
Public requestedPlayer As PlayerEditableRec
Public mapEditorCancelNag As Boolean

' Item Editor
Public lastSpawnedItems() As Byte
Public currentlyListedIndexes() As Long
Public adminMin As Boolean
Public EventList() As EventListRec
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long


Public Const GWL_WNDPROC   As Long = (-4)
Private Const WM_NOTIFY     As Long = &H4E
Private Const WM_DESTROY    As Long = &H2
Private Const WM_SETFOCUS   As Long = &H7
Private Const WM_KILLFOCUS  As Long = &H8

' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    
    ' Reset the layer to 1
    frmEditor_Map.OptLayers = 1
    
    ' Set the CurrentLayer
    CurrentLayer = 1
    
    ' Update the lists
    Call MapEditorInitShop
    
    Call MapEditorChooseTile(vbLeftButton, 0, 0)
    
    ' Set the label on the map editor to the current revision of the map
    frmEditor_Map.lblRevision.Caption = "Revision: " & Map.Revision
    
    Call ToggleGUI(False)
    Call ToggleButtons(False)

    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function ProcAdr(addr As Long) As Long
      ProcAdr = addr
End Function
Public Function getWndProcAddr() As Long
    getWndProcAddr = ProcAdr(AddressOf WindowProc)
End Function

Private Function WindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
        Select Case Msg
        Case WM_SETFOCUS
            Exit Function
        Case WM_DESTROY
            WindowProc = CallWindowProc(GetWindowLong(hWnd, -21), hWnd, Msg, wParam, lParam)
            Call SetWindowLong(hWnd, GWL_WNDPROC, GetWindowLong(hWnd, -21))
    End Select
    WindowProc = CallWindowProc(GetWindowLong(hWnd, -21), hWnd, Msg, wParam, lParam)
End Function
 
Public Sub SubClassHwnd(ByVal hWnd As Long)
        SetWindowLong hWnd, -21, SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal MovedMouse As Boolean = True)
    Dim I As Long
    Dim TmpDir As Byte
    Dim RandomSelected As Byte, Tile As Long
    Dim X2 As Long
    Dim Y2 As Long
    
   ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check for subscript out of range
    If Not IsInBounds Then Exit Sub
    
    ' Exit if we're using the fill selection tool
    If ControlDown Or ShiftDown Then Exit Sub

    ' Exit if we're using the eye dropper tool
    If frmMain.chkEyeDropper.Value = 1 Then Exit Sub
    
    If Button = vbLeftButton Then
        EditorSave = False
        If frmEditor_Map.OptLayers.Value Then
            ' No autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurrentLayer, , frmEditor_Map.scrlAutotile.Value
            Else ' Multi tile
                If frmEditor_Map.scrlAutotile.Value = 0 Then
                    MapEditorSetTile CurX, CurY, CurrentLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurrentLayer, , frmEditor_Map.scrlAutotile.Value
                End If
            End If
        ElseIf frmEditor_Map.OptEvents.Value Then
            If frmEditor_Events.Visible = False Then
                AddEvent CurX, CurY
            End If
        ElseIf frmEditor_Map.OptAttributes.Value Then
            Call MapEditorSetAttributes(Button, CurX, CurY, MovedMouse)
        ElseIf frmEditor_Map.OptBlock.Value Then
            ' Subscript out of range
            If MovedMouse Then Exit Sub
            
            ' Find coordinates if clicked
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)
            
            ' See if it hits an arrow
            For I = 1 To 4
                If X >= DirArrowX(I) And X <= DirArrowX(I) + 8 Then
                    If Y >= DirArrowY(I) And Y <= DirArrowY(I) + 8 Then
                        ' Flip the Value
                        SetDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(I), Not IsDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(I))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        EditorSave = False
        If frmEditor_Map.OptLayers.Value Then
            ' No autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then ' Single tile
                MapEditorEraseTile CurX, CurY, CurrentLayer, , frmEditor_Map.scrlAutotile.Value
            Else ' Multi-tile
                If frmEditor_Map.scrlAutotile.Value = 0 Then
                    MapEditorEraseTile CurX, CurY, CurrentLayer, True
                Else
                    MapEditorEraseTile CurX, CurY, CurrentLayer, , frmEditor_Map.scrlAutotile.Value
                End If
            End If
        ElseIf frmEditor_Map.OptAttributes.Value Then
            With Map.Tile(CurX, CurY)
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With
        End If
    End If

    CacheResources
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurrentLayer As Long, Optional ByVal MultiTile As Boolean = False, Optional ByVal Autotile As Byte = 0)
    Dim X2 As Long, Y2 As Long, RandomSelected As Integer, Tile As Integer
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Random tiles
    If frmEditor_Map.chkRandom = 1 Then
        RandomSelected = Int(Rnd * 4)
        Tile = RandomTile(RandomSelected)
        
        ' Subscript out of range
        If RandomTileSheet(RandomSelected) < 1 Or RandomTileSheet(RandomSelected) > NumTileSets Then Exit Sub
        
        X2 = Tile Mod (Tex_Tileset(RandomTileSheet(RandomSelected)).Width / PIC_X)
        Y2 = Tile / ((Tex_Tileset(RandomTileSheet(RandomSelected)).Width + PIC_Y) / PIC_Y)
        
        With Map.Tile(X, Y)
            .Layer(CurrentLayer).Tileset = RandomTileSheet(RandomSelected)
            .Layer(CurrentLayer).X = X2
            .Layer(CurrentLayer).Y = Y2
            If .Autotile(CurrentLayer) <> 0 Then
                .Autotile(CurrentLayer) = 0
                InitAutotiles
            End If
        End With
        
        CacheRenderState X2, Y2, CurrentLayer
        Exit Sub
    End If
            
    If Autotile > 0 Then
        With Map.Tile(X, Y)
            ' Set layer
            .Layer(CurrentLayer).X = EditorTileX
            .Layer(CurrentLayer).Y = EditorTileY
            .Layer(CurrentLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurrentLayer) = Autotile
            CacheRenderState X, Y, CurrentLayer
        End With
        
        ' Do a re-init so we can see our changes
        InitAutotiles
        Exit Sub
    End If
    
    If Not MultiTile Then ' Single
        With Map.Tile(X, Y)
            ' Set layer
            .Layer(CurrentLayer).X = EditorTileX
            .Layer(CurrentLayer).Y = EditorTileY
            .Layer(CurrentLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurrentLayer) = 0
            CacheRenderState X, Y, CurrentLayer
        End With
    Else ' Multi-tile
        Y2 = 0 ' Starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' Re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            .Layer(CurrentLayer).X = EditorTileX + X2
                            .Layer(CurrentLayer).Y = EditorTileY + Y2
                            .Layer(CurrentLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                            If .Autotile(CurrentLayer) <> 0 Then
                                .Autotile(CurrentLayer) = 0
                                InitAutotiles
                            End If
                            CacheRenderState X, Y, CurrentLayer
                        End With
                    End If
                End If
                
                X2 = X2 + 1
            Next
            
            Y2 = Y2 + 1
        Next
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorEraseTile(ByVal X As Long, ByVal Y As Long, ByVal CurrentLayer As Long, Optional ByVal MultiTile As Boolean = False, Optional ByVal Autotile As Byte = 0)
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Autotile > 0 Then
        With Map.Tile(X, Y)
            ' Set layer
            .Layer(CurrentLayer).X = 0
            .Layer(CurrentLayer).Y = 0
            .Layer(CurrentLayer).Tileset = 0
            If .Autotile(CurrentLayer) <> 0 Then
                .Autotile(CurrentLayer) = 0
                InitAutotiles
            End If
            CacheRenderState X, Y, CurrentLayer
        End With
        
        ' Do a re-init so we can see our changes
        InitAutotiles
        Exit Sub
    End If
    
    If Not MultiTile Then ' Single
        With Map.Tile(X, Y)
            ' Set layer
            .Layer(CurrentLayer).X = 0
            .Layer(CurrentLayer).Y = 0
            .Layer(CurrentLayer).Tileset = 0
            .Autotile(CurrentLayer) = 0
            CacheRenderState X, Y, CurrentLayer
        End With
    Else ' Multi-tile
        Y2 = 0 ' Starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' Reset x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            .Layer(CurrentLayer).X = 0
                            .Layer(CurrentLayer).Y = 0
                            .Layer(CurrentLayer).Tileset = 0
                            If .Autotile(CurrentLayer) <> 0 Then
                                .Autotile(CurrentLayer) = 0
                                InitAutotiles
                            End If
                            CacheRenderState X, Y, CurrentLayer
                        End With
                    End If
                End If
                
                X2 = X2 + 1
            Next
            
            Y2 = Y2 + 1
        Next
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorEraseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorSetAttributes(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal MovedMouse As Boolean = True)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With Map.Tile(X, Y)
        ' Blocked Tile
        If frmEditor_Map.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
        
        ' Warp Tile
        If frmEditor_Map.optWarp.Value Then
            .Type = TILE_TYPE_WARP
            .Data1 = EditorWarpMap
            .Data2 = EditorWarpX
            .Data3 = EditorWarpY
        End If
        
        ' Item Spawn
        If frmEditor_Map.optItem.Value Then
            .Type = TILE_TYPE_ITEM
            .Data1 = ItemEditorNum
            .Data2 = ItemEditorValue
            .Data3 = 0
        End If
        
        ' NPC Avoid
        If frmEditor_Map.optNPCAvoid.Value Then
            .Type = TILE_TYPE_NPCAVOID
            .Data1 = 0
            .Data2 = 0
            .Data3 = 0
        End If
        
        ' Resource
        If frmEditor_Map.optResource.Value Then
            .Type = TILE_TYPE_RESOURCE
            .Data1 = ResourceEditorNum
            .Data2 = 0
            .Data3 = 0
        End If
        
        ' NPC Spawn
        If frmEditor_Map.optNPCSpawn.Value Then
            .Type = TILE_TYPE_NPCSPAWN
            .Data1 = SpawnNPCNum
            .Data2 = SpawnNPCDir
            .Data3 = 0
        End If
        
        ' Shop
        If frmEditor_Map.optShop.Value Then
            .Type = TILE_TYPE_SHOP
            .Data1 = EditorShop
            .Data2 = 0
            .Data3 = 0
        End If
        
        ' Bank
        If frmEditor_Map.optBank.Value Then
            .Type = TILE_TYPE_BANK
            .Data1 = 0
            .Data2 = 0
            .Data3 = 0
        End If
        
        ' Heal
        If frmEditor_Map.optHeal.Value Then
            .Type = TILE_TYPE_HEAL
            .Data1 = MapEditorVitalType
            .Data2 = MapEditorVitalAmount
            .Data3 = 0
        End If
        
        ' Trap
        If frmEditor_Map.optTrap.Value Then
            .Type = TILE_TYPE_TRAP
            .Data1 = MapEditorVitalType
            .Data2 = MapEditorVitalAmount
            .Data3 = 0
        End If
        
        ' Slide
        If frmEditor_Map.optSlide.Value Then
            .Type = TILE_TYPE_SLIDE
            .Data1 = MapEditorSlideDir
            .Data2 = 0
            .Data3 = 0
        End If
        
        ' Checkpoint
        If frmEditor_Map.optCheckpoint.Value Then
            X = X - (CurX * 32)
            Y = Y - (CurY * 32)
            .Type = TILE_TYPE_CHECKPOINT
            .Data1 = GetPlayerMap(MyIndex)
            .Data2 = CurX
            .Data3 = CurY
        End If
        
        ' Sound
        If frmEditor_Map.optSound.Value Then
            .Type = TILE_TYPE_SOUND
            .Data1 = 0
            .Data2 = 0
            .Data3 = 0
            .Data4 = MapEditorSound
        End If
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorSetAttributes", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorSave()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call SendSaveMap
    Unload frmEditor_MapProperties
    EditorSave = True
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "MapEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorCancel()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If InMapEditor And IsLogging = False Then
        If AlertMsg("Are you sure you want to discard changes made to the map?", False, False) = YES Then
            SendNeedMap
        Else
            Call MapEditorSave
            EditorSave = False
        End If
    End If
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorClearLayer()
    Dim I As Long, X As Long, Y As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    'If AlertMsg("Are you sure you wish to clear this layer", False, False) = YES Then
        If CurrentLayer = 0 Then Exit Sub
    
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y).Layer(CurrentLayer)
                    .X = 0
                    .Y = 0
                    .Tileset = 0
                    CacheRenderState X, Y, CurrentLayer
                End With
            Next
        Next
    'End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorFillLayer()
    Dim X As Long
    Dim Y As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    'If AlertMsg("Are you sure you wish to fill this layer", False, False) = YES Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y).Layer(CurrentLayer)
                    .X = EditorTileX
                    .Y = EditorTileY
                    .Tileset = frmEditor_Map.scrlTileSet.Value
                End With
                
                Map.Tile(X, Y).Autotile(CurrentLayer) = frmEditor_Map.scrlAutotile.Value
                CacheRenderState X, Y, CurrentLayer
            Next
        Next
    'End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorFillSelection()
    Dim X As Long
    Dim Y As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
        
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If frmEditor_Map.fraLayers.Visible Then
                If Map.Tile(CurX, CurY).Layer(CurrentLayer).X = Map.Tile(X, Y).Layer(CurrentLayer).X Then
                    If Map.Tile(CurX, CurY).Layer(CurrentLayer).Y = Map.Tile(X, Y).Layer(CurrentLayer).Y Then
                        If Map.Tile(CurX, CurY).Layer(CurrentLayer).Tileset = Map.Tile(X, Y).Layer(CurrentLayer).Tileset Then
                            With Map.Tile(X, Y).Layer(CurrentLayer)
                                .X = EditorTileX
                                .Y = EditorTileY
                                .Tileset = frmEditor_Map.scrlTileSet.Value
                            End With
                            
                            Map.Tile(CurX, CurY).Autotile(CurrentLayer) = frmEditor_Map.scrlAutotile.Value
                            CacheRenderState X, Y, CurrentLayer
                        End If
                    End If
                End If
            ElseIf frmEditor_Map.fraAttribs.Visible Then
                If Map.Tile(CurX, CurY).Layer(CurrentLayer).X = Map.Tile(X, Y).Layer(CurrentLayer).X Then
                    If Map.Tile(CurX, CurY).Layer(CurrentLayer).Y = Map.Tile(X, Y).Layer(CurrentLayer).Y Then
                        Call MapEditorSetAttributes(vbLeftButton, X, Y)
                    End If
                End If
            End If
        Next
    Next
    
    ' Now cache the positions
    If frmEditor_Map.fraLayers.Visible Then
        InitAutotiles
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorFillSelection", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorClearSelection()
    Dim X As Long
    Dim Y As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
        
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If frmEditor_Map.fraLayers.Visible Then
                If Map.Tile(CurX, CurY).Layer(CurrentLayer).X = Map.Tile(X, Y).Layer(CurrentLayer).X Then
                    If Map.Tile(CurX, CurY).Layer(CurrentLayer).Y = Map.Tile(X, Y).Layer(CurrentLayer).Y Then
                        If Map.Tile(CurX, CurY).Layer(CurrentLayer).Tileset = Map.Tile(X, Y).Layer(CurrentLayer).Tileset Then
                            With Map.Tile(X, Y).Layer(CurrentLayer)
                                .X = 0
                                .Y = 0
                                .Tileset = 0
                            End With
                        
                            Map.Tile(CurX, CurY).Autotile(CurrentLayer) = 0
                            CacheRenderState X, Y, CurrentLayer
                        End If
                    End If
                End If
            ElseIf frmEditor_Map.fraAttribs.Visible Then
                If Map.Tile(CurX, CurY).Layer(CurrentLayer).X = Map.Tile(X, Y).Layer(CurrentLayer).X Then
                    If Map.Tile(CurX, CurY).Layer(CurrentLayer).Y = Map.Tile(X, Y).Layer(CurrentLayer).Y Then
                        With Map.Tile(X, Y)
                            .Type = 0
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                        End With
                    End If
                End If
            End If
        Next
    Next
    
    ' Now cache the positions
    If frmEditor_Map.fraLayers.Visible Then
        InitAutotiles
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorClearSelection", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorClearAttributes()
    Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    'If AlertMsg("Are you sure you wish to clear all the attributes on this map", False, False) = YES Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    .Type = 0
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End With
            Next
        Next
    'End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorClearAttributes", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorFillAttributes(ByVal Button As Integer)
    Dim X As Long, Y As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    'If AlertMsg("Are you sure you wish to fill this attribute on the entire map", False, False) = YES Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Call MapEditorSetAttributes(Button, X, Y)
            Next
        Next
    'End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorFillAttributes", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorLeaveMap()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If InMapEditor Then
        If EditorSave = False Then
            If AlertMsg("Save changes to current map, before leaving?", False, False) = YES Then
                Call MapEditorSave
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorLeaveMap", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
    Dim I As Long
    Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    Item_Changed(EditorIndex) = True

    ' Populate the cache if we need to
    If Not HasPopulated Then PopulateLists
    
    ' Add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None"

    For I = 1 To UBound(SoundCache)
        frmEditor_Item.cmbSound.AddItem SoundCache(I)
    Next

    With Item(EditorIndex)
        ' Check for invalid values
        If .Pic < 1 Then .Pic = 1
        If .Pic > frmEditor_Item.scrlPic.max Then .Pic = frmEditor_Item.scrlPic.max
        
        ' Basic data
        frmEditor_Item.txtName.text = Trim$(.Name)
        frmEditor_Item.scrlPic.Value = .Pic
        
        If .Type > frmEditor_Item.cmbType.ListCount - 1 Then
            frmEditor_Item.cmbType.ListIndex = frmEditor_Item.cmbType.ListCount - 1
        Else
            frmEditor_Item.cmbType.ListIndex = .Type
        End If
        
        frmEditor_Item.scrlAnim.Value = .Animation
        
        If .Data3 > frmEditor_Item.cmbTool.ListCount - 1 Then
            frmEditor_Item.cmbTool.ListIndex = frmEditor_Item.cmbTool.ListCount - 1
        Else
            frmEditor_Item.cmbTool.ListIndex = .Data3
        End If
        
        frmEditor_Item.scrlPaperdoll = .Paperdoll
        frmEditor_Item.scrlDurability = .Data1
        frmEditor_Item.txtDesc.text = Trim$(.Desc)
        frmEditor_Item.cmbEquipSlot.ListIndex = .EquipSlot
        
        If .Type = ITEM_TYPE_SPELL Then
            frmEditor_Item.scrlSpell.Value = .Data1
        End If
        
        frmEditor_Item.chkHoT.Value = .HoT
        frmEditor_Item.cmbProficiencyReq.ListIndex = .ProficiencyReq
        frmEditor_Item.chkTwoHanded.Value = .TwoHanded
        frmEditor_Item.chkStackable.Value = .stackable
        frmEditor_Item.chkIndestructable = .Indestructable
        frmEditor_Item.cmbSkillReq.ListIndex = .SkillReq
        
        Call UpdateSpellScrollBars
        
        ' Reusable
        If .IsReusable Then
            For I = 1 To frmEditor_Item.chkReusable.Count - 1
                frmEditor_Item.chkReusable.Item(I) = 1
            Next
        Else
            For I = 1 To frmEditor_Item.chkReusable.Count - 1
                frmEditor_Item.chkReusable.Item(I) = 0
            Next
        End If
        
        ' Sprites and Titles
        If .Data1 > 0 Then
            If .Data1 > frmEditor_Item.scrlSprite.max Then
                frmEditor_Item.scrlSprite.Value = frmEditor_Item.scrlSprite.max
            Else
                frmEditor_Item.scrlSprite.Value = .Data1
            End If
            
            If .Data1 > frmEditor_Item.scrlTitle.max Then
                frmEditor_Item.scrlTitle.Value = frmEditor_Item.scrlTitle.max
            Else
                frmEditor_Item.scrlTitle.Value = .Data1
            End If
        Else
            frmEditor_Item.scrlSprite.Value = 1
            frmEditor_Item.scrlTitle.Value = 1
        End If
        
        ' Gender requirement
        frmEditor_Item.cmbGenderReq.ListIndex = .GenderReq

        ' Find the sound we have set
        If frmEditor_Item.cmbSound.ListCount > 0 Then
            For I = 1 To frmEditor_Item.cmbSound.ListCount
                If Len(Trim$(.Sound)) > 0 Then
                    If frmEditor_Item.cmbSound.List(I) = Trim$(.Sound) Then
                        frmEditor_Item.cmbSound.ListIndex = I
                        SoundSet = True
                    End If
                End If
            Next
        End If
        
        If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0

        ' Resources
        If Item(EditorIndex).ChanceModifier = 0 Then Item(EditorIndex).ChanceModifier = 1
        frmEditor_Item.scrlChanceModifier.Value = Item(EditorIndex).ChanceModifier

        With frmEditor_Item
            If frmEditor_Item.cmbTool.ListIndex > 0 And frmEditor_Item.cmbTool.ListIndex <> 4 Then
                .scrlChanceModifier.Visible = True
                .lblChance.Visible = True
                .lblDamage.Visible = False
                .scrlDamage.Visible = False
            Else
                .scrlChanceModifier.Visible = False
                .lblChance.Visible = False
                .lblDamage.Visible = True
                .scrlDamage.Visible = True
            End If
        End With
        
        ' Loop for stats
        For I = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatBonus(I).Value = .Add_Stat(I)
        Next
        
        frmEditor_Item.scrlDamage.Value = .Data2
        
        If .WeaponSpeed < 100 Then .WeaponSpeed = 100
        frmEditor_Item.scrlSpeed.Value = .WeaponSpeed

        ' Auto Life
        If .AddHP < 1 Then
            frmEditor_Item.scrlHP.Value = 1
        Else
            frmEditor_Item.scrlHP.Value = .AddHP
        End If
        
        frmEditor_Item.scrlMP.Value = .AddMP
        
        If .Data1 > 0 Then
            frmEditor_Item.chkWarpAway = 1
        Else
            frmEditor_Item.chkWarpAway = 0
        End If
        
        ' Teleport
        frmEditor_Item.scrlMap.Value = .Data1
        frmEditor_Item.scrlX.Value = .Data2
        frmEditor_Item.scrlY.Value = .Data3

        ' Tools
        frmEditor_Item.cmbTool.ListIndex = .Tool
        
        ' Consumable data
        frmEditor_Item.scrlAddHP.Value = .AddHP
        frmEditor_Item.scrlAddMP.Value = .AddMP
        frmEditor_Item.scrlAddEXP.Value = .AddEXP
        frmEditor_Item.scrlCastSpell.Value = .CastSpell
        frmEditor_Item.chkInstaCast.Value = .InstaCast
        
        If .Data1 < 1 Then
            frmEditor_Item.scrlDuration.Value = 1
        ElseIf .Data1 > frmEditor_Item.scrlDuration.max Then
            frmEditor_Item.scrlDuration.Value = frmEditor_Item.scrlDuration.max
        Else
            frmEditor_Item.scrlDuration.Value = .Data1
        End If
        
        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq
        
        ' Loop for stats
        For I = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(I).Value = .Stat_Req(I)
        Next
        
        ItemClassReqListInit
        
        ' Information
        frmEditor_Item.txtPrice.text = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        
        ' Rarity
        If Trim$(.Name) = vbNullString Then
            frmEditor_Item.scrlRarity.Value = 1
        Else
            frmEditor_Item.scrlRarity.Value = .Rarity
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_RECIPE) Then
            frmEditor_Item.fraRecipe.Visible = True
            frmEditor_Item.scrlItem1.Value = .Data1
            frmEditor_Item.scrlItem2.Value = .Data2
            frmEditor_Item.scrlResult.Value = .Data3
            frmEditor_Item.scrlSkill.Value = .Skill
            frmEditor_Item.ScrlSkillExp.Value = .SkillExp
            frmEditor_Item.ScrlSkillLevelReq = .SkillLevelReq
            frmEditor_Item.ScrlToolRequired.Value = .ToolRequired
        Else
            frmEditor_Item.fraRecipe.Visible = False
        End If
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ItemEditorSave()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ITEMS
        If Item_Changed(I) Then
            Call SendSaveItem(I)
        End If
    Next
    
    'Unload frmEditor_Item
    'Editor = 0
    ClearChanged_Item
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ItemEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Quest()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2 ' 2 = boolean length
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Quest", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ItemEditorCancel()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    ClearChanged_Item
    ClearItems
    SendRequestItems
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Item()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' //////////////////////
' // Animation Editor //
' //////////////////////
Public Sub AnimationEditorInit()
    Dim I As Long
    Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    Animation_Changed(EditorIndex) = True
    
    ' Populate the cache if we need to
    If Not HasPopulated Then PopulateLists

    ' Add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None"
    
    For I = 1 To UBound(SoundCache)
        frmEditor_Animation.cmbSound.AddItem SoundCache(I)
    Next

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.text = Trim$(.Name)
        
        ' Find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount > 0 Then
            For I = 1 To frmEditor_Animation.cmbSound.ListCount
                If Len(Trim$(.Sound)) > 0 Then
                    If frmEditor_Animation.cmbSound.List(I) = Trim$(.Sound) Then
                        frmEditor_Animation.cmbSound.ListIndex = I
                        SoundSet = True
                    End If
                End If
            Next
        End If
        
        If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        
        For I = 0 To 1
            If .Sprite(I) < 1 Or .Sprite(I) > NumAnimations Then
                frmEditor_Animation.scrlSprite(I).Value = 0
            Else
                frmEditor_Animation.scrlSprite(I).Value = .Sprite(I)
            End If
            
            If .Frames(I) = 0 Then .Frames(I) = 1
            frmEditor_Animation.scrlFrameCount(I).Value = .Frames(I)
            
            If .LoopCount(I) = 0 Then .LoopCount(I) = 1
            frmEditor_Animation.scrlLoopCount(I).Value = .LoopCount(I)
            
            If .looptime(I) = 0 Then .looptime(I) = 1
            frmEditor_Animation.scrlLoopTime(I).Value = .looptime(I)
            
            ' Set the loop time to 40 if it is 1
            If frmEditor_Animation.scrlLoopTime(I).Value = 1 Then
                frmEditor_Animation.scrlLoopTime(I).Value = 40
            End If
        Next
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub AnimationEditorSave()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ANIMATIONS
        If Animation_Changed(I) Then
            Call SendSaveAnimation(I)
        End If
    Next
    
    'Unload frmEditor_Animation
    'Editor = 0
    ClearChanged_Animation
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "AnimationEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub AnimationEditorCancel()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Animation()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' ////////////////
' // NPC Editor //
' ////////////////
Public Sub NPCEditorInit()
    Dim I As Long
    Dim MusicSet As Boolean, SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_NPC.Visible = False Then Exit Sub
    
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    NPC_Changed(EditorIndex) = True

    ' Populate the cache if we need to
    If Not HasPopulated Then
        PopulateLists
    End If
    
    ' Add the array to the combo
    frmEditor_NPC.cmbMusic.Clear
    frmEditor_NPC.cmbMusic.AddItem "None"
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None"
    
    For I = 1 To UBound(MusicCache)
        frmEditor_NPC.cmbMusic.AddItem MusicCache(I)
    Next
    
    For I = 1 To UBound(SoundCache)
        frmEditor_NPC.cmbSound.AddItem SoundCache(I)
    Next
    
    With frmEditor_NPC
        ' Check for an invalid Value
        If NPC(EditorIndex).Sprite < 0 Or NPC(EditorIndex).Sprite > .scrlSprite.max Then NPC(EditorIndex).Sprite = 0
        
        .txtName.text = Trim$(NPC(EditorIndex).Name)
        .txtAttackSay.text = Trim$(NPC(EditorIndex).AttackSay)
        .txtTitle.text = Trim$(NPC(EditorIndex).title)
        
        If NumCharacters < NPC(EditorIndex).Sprite Then NPC(EditorIndex).Sprite = NumCharacters
        .scrlSprite.Value = NPC(EditorIndex).Sprite
        
        .txtSpawnSecs.text = CStr(NPC(EditorIndex).SpawnSecs)
        .cmbBehavior.ListIndex = NPC(EditorIndex).Behavior
        .cmbFaction.ListIndex = NPC(EditorIndex).Faction
        .scrlRange.Value = NPC(EditorIndex).Range
        .txtHP.text = NPC(EditorIndex).HP
        .txtMP.text = NPC(EditorIndex).MP
        .txtExp.text = NPC(EditorIndex).exp
        .scrlLevel.Value = NPC(EditorIndex).Level
        .scrlDamage.Value = NPC(EditorIndex).Damage
        
        .scrlAnimation.Value = NPC(EditorIndex).Animation
        
        ' Drops
        .scrlDrop.Value = 1
        .txtChance.text = NPC(EditorIndex).DropChance(1)
        .scrlNum.Value = NPC(EditorIndex).DropItem(1)
        .scrlValue.Value = NPC(EditorIndex).DropValue(1)
        .chkDrop.Value = NPC(EditorIndex).DropRandom(1)
        
        ' Spells
        .scrlSpell.Value = 1
        .scrlSpellNum.Value = NPC(EditorIndex).Spell(1)
        
        ' Faction threat
        If NPC(EditorIndex).FactionThreat = True Then
            .chkFactionThreat.Value = 1
        Else
            .chkFactionThreat.Value = 0
        End If
        
        ' Switches
        .cmbPlayerSwitch.Clear
        .cmbPlayerSwitch.AddItem "None"
        For I = 1 To MAX_SWITCHES
            .cmbPlayerSwitch.AddItem I & ". " & Switches(I)
        Next
        
        ' Variables
        .cmbPlayerVar.Clear
        .cmbPlayerVar.AddItem "None"
        For I = 1 To MAX_VARIABLES
            .cmbPlayerVar.AddItem I & ". " & Variables(I)
        Next
        
        .cmbPlayerVar.ListIndex = NPC(EditorIndex).VariableNum
        .cmbPlayerSwitch.ListIndex = NPC(EditorIndex).SwitchNum
        .scrlPlayerVar.Value = NPC(EditorIndex).VariableVal
        .scrlPlayerSwitch.Value = NPC(EditorIndex).SwitchVal
        .chkPlayerVar.Value = NPC(EditorIndex).AddToVariable
        
        ' Find the music we have set
        If .cmbMusic.ListCount > 0 Then
            For I = 1 To .cmbMusic.ListCount
                If Len(Trim$(NPC(EditorIndex).Music)) > 0 Then
                    If .cmbMusic.List(I) = Trim$(NPC(EditorIndex).Music) Then
                        .cmbMusic.ListIndex = I
                        MusicSet = True
                    End If
                End If
            Next
        End If
        
        If Not MusicSet Or .cmbMusic.ListIndex = -1 Then .cmbMusic.ListIndex = 0
        
        ' Find the sound we have set
        If .cmbSound.ListCount > 0 Then
            For I = 1 To .cmbSound.ListCount
                If Len(Trim$(NPC(EditorIndex).Sound)) > 0 Then
                    If .cmbSound.List(I) = Trim$(NPC(EditorIndex).Sound) Then
                        .cmbSound.ListIndex = I
                        SoundSet = True
                    End If
                End If
            Next
        End If
        
        If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        
        For I = 1 To Stats.Stat_Count - 1
            .scrlStat(I).Value = NPC(EditorIndex).Stat(I)
        Next
        
        .chkAnimated = NPC(EditorIndex).Animated
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "NPCEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub NPCEditorSave()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_NPCS
        If NPC_Changed(I) Then
            Call SendSaveNPC(I)
        End If
    Next
    
    'Unload frmEditor_NPC
    'Editor = 0
    ClearChanged_NPC
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "NPCEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub NPCEditorCancel()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    ClearChanged_NPC
    ClearNPCs
    SendRequestNPCs
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "NPCEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_NPC()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    Exit Sub
      
' Error handler
ErrorHandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' /////////////////////
' // Resource Editor //
' /////////////////////
Public Sub ResourceEditorInit()
    Dim I As Long
    Dim SoundSet As Boolean

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    If frmEditor_Resource.Visible = False Then Exit Sub
    
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    Resource_Changed(EditorIndex) = True
    
    ' Populate the cache if we need to
    If Not HasPopulated Then PopulateLists

    ' Add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None"
    
    For I = 1 To UBound(SoundCache)
        frmEditor_Resource.cmbSound.AddItem SoundCache(I)
    Next
    
    With frmEditor_Resource
        .txtName.text = Trim$(Resource(EditorIndex).Name)
        .txtSuccess.text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtEmpty.text = Trim$(Resource(EditorIndex).EmptyMessage)
        .txtFail.text = Trim$(Resource(EditorIndex).FailMessage)
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlLevelReq.Value = Resource(EditorIndex).LevelReq
        
        If Resource(EditorIndex).ToolRequired = 0 Then Resource(EditorIndex).ToolRequired = 1
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        
        .scrlAnimation.Value = Resource(EditorIndex).Animation
        
        If Resource(EditorIndex).Skill = 0 Then Resource(EditorIndex).Skill = 1
        .scrlSkill.Value = Resource(EditorIndex).Skill
        
        .scrlRespawn.Value = Resource(EditorIndex).RespawnTime
        .scrlExp.Value = Resource(EditorIndex).exp
        
        ' Find the sound we have set
        If .cmbSound.ListCount > 0 Then
            For I = 1 To .cmbSound.ListCount
                If Len(Trim$(Resource(EditorIndex).Sound)) > 0 Then
                    If .cmbSound.List(I) = Trim$(Resource(EditorIndex).Sound) Then
                        .cmbSound.ListIndex = I
                        SoundSet = True
                    End If
                End If
            Next
        End If
        
        If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
    
        If Resource(EditorIndex).Reward_Min = 0 Then Resource(EditorIndex).Reward_Min = 1
        .scrlRewardMin.Value = Resource(EditorIndex).Reward_Min
    
        If Resource(EditorIndex).Reward_Max = 0 Then Resource(EditorIndex).Reward_Max = 1
        .scrlRewardMax.Value = Resource(EditorIndex).Reward_Max
        
        If Resource(EditorIndex).LowChance = 0 Then Resource(EditorIndex).LowChance = 1
        .scrlLowChance.Value = Resource(EditorIndex).LowChance
        
        If Resource(EditorIndex).HighChance = 0 Then Resource(EditorIndex).HighChance = 2
        .scrlHighChance.Value = Resource(EditorIndex).HighChance
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ResourceEditorSave()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_RESOURCES
        If Resource_Changed(I) Then
            Call SendSaveResource(I)
        End If
    Next
    
    'Unload frmEditor_Resource
    'Editor = 0
    ClearChanged_Resource
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ResourceEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ResourceEditorCancel()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    ClearChanged_Resource
    ClearResources
    SendRequestResources
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Resource()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
    Exit Sub
     
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    Shop_Changed(EditorIndex) = True
    
    With frmEditor_Shop
        .txtName.text = Trim$(Shop(EditorIndex).Name)
        
        If Shop(EditorIndex).BuyRate > 0 And Shop(EditorIndex).BuyRate <= .scrlBuy.max Then
            .scrlBuy.Value = Shop(EditorIndex).BuyRate
        Else
            .scrlBuy.Value = 100
        End If
        
        If Shop(EditorIndex).SellRate > 0 And Shop(EditorIndex).SellRate <= .scrlSell.max Then
            .scrlSell.Value = Shop(EditorIndex).SellRate
        Else
            .scrlSell.Value = 100
        End If
        
        .chkCanFix.Value = Shop(EditorIndex).CanFix
    End With
    
    UpdateShopTrade
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    frmEditor_Shop.lstTradeItem.Clear
    
    For I = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(I)
            ' If none, show as none
            If .Item = 0 Or .CostItem = 0 And .CostItem2 = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            ElseIf .CostItem2 = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem I & ": " & .ItemValue & "X " & Trim$(Item(.Item).Name) & " for " & .CostValue & "X " & Trim$(Item(.CostItem).Name)
            ElseIf .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem I & ": " & .ItemValue & "X " & Trim$(Item(.Item).Name) & " for " & .CostValue & "X " & Trim$(Item(.CostItem2).Name)
            ElseIf .CostItem > 0 And .CostItem2 > 0 Then
                frmEditor_Shop.lstTradeItem.AddItem I & ": " & .ItemValue & "X " & Trim$(Item(.Item).Name) & " for " & .CostValue & "X " & Trim$(Item(.CostItem).Name) & " & " & .CostValue2 & "X " & Trim$(Item(.CostItem2).Name)
            End If
        End With
    Next
    
    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
    Exit Sub
     
' Error handler
ErrorHandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ShopEditorSave()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SHOPS
        If Shop_Changed(I) Then
            Call SendSaveShop(I)
        End If
    Next
    
    'Unload frmEditor_Shop
    'Editor = 0
    ClearChanged_Shop
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ShopEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ShopEditorCancel()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    ClearChanged_Shop
    ClearShops
    SendRequestShops
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Shop()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
    
' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
    Dim I As Long
    Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    
    If frmEditor_Spell.Visible = False Then Exit Sub
    
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    Spell_Changed(EditorIndex) = True
    
    ' Populate the cache if we need to
    If Not HasPopulated Then PopulateLists

    ' Add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None"
    
    For I = 1 To UBound(SoundCache)
        frmEditor_Spell.cmbSound.AddItem SoundCache(I)
    Next
    
    With frmEditor_Spell
        If Spell(EditorIndex).IsAoe = True Then
            .scrlAOE.Enabled = True
            .chkAoE = 1
        Else
            .scrlAOE.Enabled = False
            .chkAoE = 0
        End If
        
        If Spell(EditorIndex).WeaponDamage = False Then
            .chkWeaponDamage = 0
        Else
            .chkWeaponDamage = 1
        End If
        
        SpellClassListInit
        
        ' Set values
        .txtName.text = Trim$(Spell(EditorIndex).Name)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.Value = Spell(EditorIndex).MPCost
        .scrlLevel.Value = Spell(EditorIndex).LevelReq
        .scrlAccess.Value = Spell(EditorIndex).AccessReq
        .scrlCast.Value = Spell(EditorIndex).CastTime
        
        If Spell(EditorIndex).CDTime = 0 Then Spell(EditorIndex).CDTime = 1
        .txtCool.text = Spell(EditorIndex).CDTime
        
        .scrlIcon.Value = Spell(EditorIndex).Icon
        .scrlMap.Value = Spell(EditorIndex).Map
        .scrlX.Value = Spell(EditorIndex).X
        .scrlY.Value = Spell(EditorIndex).Y
        .scrlDir.Value = Spell(EditorIndex).Dir
        .scrlVital.Value = Spell(EditorIndex).Vital
        .scrlDuration.Value = Spell(EditorIndex).Duration
        .scrlInterval.Value = Spell(EditorIndex).Interval
        .scrlRange.Value = Spell(EditorIndex).Range
        .scrlAOE.Value = Spell(EditorIndex).AoE
        .scrlAnimCast.Value = Spell(EditorIndex).CastAnim
        .scrlAnim.Value = Spell(EditorIndex).SpellAnim
        .scrlStun.Value = Spell(EditorIndex).StunDuration
        .scrlSprite.Value = Spell(EditorIndex).Sprite
        .txtDesc.text = Trim$(Spell(EditorIndex).Desc)
        .scrlCastRequired.Value = Spell(EditorIndex).CastRequired
        .scrlRankUp.Value = Spell(EditorIndex).NewSpell
        
        If .cmbSound.ListCount > 0 Then
            For I = 1 To .cmbSound.ListCount
                If Len(Trim$(Spell(EditorIndex).Sound)) > 0 Then
                    If .cmbSound.List(I) = Trim$(Spell(EditorIndex).Sound) Then
                        .cmbSound.ListIndex = I
                        SoundSet = True
                    End If
                End If
            Next
        End If
        
        If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SpellEditorSave()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SPELLS
        If Spell_Changed(I) Then
            Call SendSaveSpell(I)
        End If
    Next
    
    'Unload frmEditor_Spell
    'Editor = 0
    ClearChanged_Spell
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SpellEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SpellEditorCancel()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
    
    For I = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(I) > 0 Then
            Call SendRequestSpellCooldown(I)
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Spell()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = Boolean length
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearAttributeFrames()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Map
        .fraMapItem.Visible = False
        .fraResource.Visible = False
        .fraNpcSpawn.Visible = False
        .fraTrap.Visible = False
        .fraShop.Visible = False
        .fraHeal.Visible = False
        .fraMapWarp.Visible = False
        .fraSlide.Visible = False
        .fraSoundEffect.Visible = False
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearAttributeFrames", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapPropertiesInit()
    Dim I As Long, MusicSet As Boolean, SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Populate the cache if we need to
    If Not HasPopulated Then
        PopulateLists
    End If

    With frmEditor_MapProperties
        .chkAutoSpawn.Value = Map.NPCSpawnType(1)
        
        ' Clear the list
        .lstMusic.Clear
        .lstMusic.AddItem "None"
        .cmbSound.Clear
        .cmbSound.AddItem "None"
        
        ' Build music list
        For I = 1 To UBound(MusicCache)
            .lstMusic.AddItem MusicCache(I)
        Next
        
        ' Build sound list
        For I = 1 To UBound(SoundCache)
            frmEditor_MapProperties.cmbSound.AddItem SoundCache(I)
        Next
        
        ' Clear the combo box
        .cmbMoral.Clear
        
        For I = 1 To MAX_MORALS
            .cmbMoral.AddItem I & ": " & Trim$(Moral(I).Name)
        Next

        ' Find the music we have set
        If .lstMusic.ListCount > 1 Then
            For I = 1 To .lstMusic.ListCount
                If Len(Trim$(Map.Music)) > 0 Then
                    If .lstMusic.List(I) = Trim$(Map.Music) Then
                        .lstMusic.ListIndex = I
                        MusicSet = True
                    End If
                End If
            Next
        End If
        
        If Not MusicSet Or .lstMusic.ListIndex = -1 Then .lstMusic.ListIndex = 0
        
        If .cmbSound.ListCount > 1 Then
            For I = 0 To .cmbSound.ListCount
                If Len(Trim$(Map.BGS)) > 0 Then
                    If .cmbSound.List(I) = Trim$(Map.BGS) Then
                        .cmbSound.ListIndex = I
                        SoundSet = True
                    End If
                End If
            Next
        End If
        
        If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        
        ' Update the combo box
        .cmbNpcs.AddItem "None"
        .cmbNpcs.ListIndex = 0
        
        ' Load all the npcs that can be selected into the combo box
        For I = 1 To MAX_NPCS
            .cmbNpcs.AddItem I & ": " & Trim$(NPC(I).Name)
        Next
        
        .cmbWeather.ListIndex = Map.Weather
        .scrlWeatherIntensity.Value = Map.WeatherIntensity
        
        .ScrlFog.Value = Map.Fog
        .ScrlFogSpeed.Value = Map.FogSpeed
        .scrlFogOpacity.Value = Map.FogOpacity
        
        .scrlPanorama.Value = Map.Panorama
        
        .ScrlR.Value = Map.Red
        .ScrlG.Value = Map.Green
        .ScrlB.Value = Map.Blue
        .scrlA.Value = Map.Alpha

        ' Load the npcs into the lstNPCs
        Call LoadMapPropertiesNPCs
        
        ' Reset of it
        .txtUp.text = CStr(Map.Up)
        .txtDown.text = CStr(Map.Down)
        .txtLeft.text = CStr(Map.Left)
        .txtRight.text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral - 1
        .txtBootMap.text = CStr(Map.BootMap)
        .txtBootX.text = CStr(Map.BootX)
        .txtBootY.text = CStr(Map.BootY)
        .txtName.text = Trim$(Map.Name)
        .lblMap.Caption = "Current Map: " & GetPlayerMap(MyIndex)
        
        .txtMaxX.text = Map.MaxX
        .txtMaxY.text = Map.MaxY
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapPropertiesInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub LoadMapPropertiesNPCs()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Load the npcs into the list
    With Map
        For I = 1 To MAX_MAP_NPCS
            If .NPC(I) < 1 Or .NPC(I) > MAX_NPCS Then
                frmEditor_MapProperties.lstNpcs.AddItem I & ": None"
            Else
                frmEditor_MapProperties.lstNpcs.AddItem I & ": " & Trim$(NPC(.NPC(I)).Name)
            End If
        Next
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "LoadMapPropertiesNPCs", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorInitShop()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmEditor_Map.cmbShop.Clear
    frmEditor_Map.cmbShop.AddItem "None"
    
    ' Set shops for the shop attribute
    For I = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem I & ": " & Shop(I).Name, I
    Next
    
    ' Reset the shop list Index
    frmEditor_Map.cmbShop.ListIndex = 0
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "MapEditorInitShop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub BanEditorInit()
    ' Check if the form is visible if not then exit
    If frmEditor_Ban.Visible = False Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorIndex = frmEditor_Ban.lstIndex.ListIndex + 1
    Ban_Changed(EditorIndex) = True

    With frmEditor_Ban
        .txtName.text = Trim$(Ban(EditorIndex).PlayerName)
        .txtLogin.text = Trim$(Ban(EditorIndex).PlayerLogin)
        .txtIP.text = Trim$(Ban(EditorIndex).IP)
        .txtSerial.text = Trim$(Ban(EditorIndex).HDSerial)
        .txtReason.text = Trim$(Ban(EditorIndex).Reason)
        .txtDate.text = Trim$(Ban(EditorIndex).Date)
        .txtTime.text = Trim$(Ban(EditorIndex).time)
        .txtBy.text = Trim$(Ban(EditorIndex).By)
    End With
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "BanEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub BanEditorSave()
    Dim I As Long
    
    ' Subscript out of range
    If EditorIndex < 1 Or EditorIndex > MAX_BANS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_BANS
        If Ban_Changed(I) Then
            Call SendSaveBan(I)
        End If
    Next
    
    'Unload frmEditor_Ban
    'Editor = 0
    ClearChanged_Ban
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "BanEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub BanEditorCancel()
    ' Subscript out of range
    If EditorIndex < 1 Or EditorIndex > MAX_BANS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Editor = 0
    ClearChanged_Ban
    ClearBans
    SendRequestBans
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "BanEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Ban()
    ' Subscript out of range
    If EditorIndex < 1 Or EditorIndex > MAX_BANS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ZeroMemory Ban_Changed(1), MAX_BANS * 2 ' 2 = Boolean length
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Ban", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub TitleEditorInit()
    ' Check if the form is visible if not then exit
    If frmEditor_Title.Visible = False Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorIndex = frmEditor_Title.lstIndex.ListIndex + 1
    Title_Changed(EditorIndex) = True
    
    With frmEditor_Title
        .txtName.text = Trim$(title(EditorIndex).Name)
        .scrlColor.Value = title(EditorIndex).Color
        .lblLevelReq.Caption = "Level Requirement: " & Trim$(title(EditorIndex).LevelReq)
        .scrlLevelReq.Value = title(EditorIndex).LevelReq
        .txtDesc = Trim$(title(EditorIndex).Desc)
        .scrlPKReq = title(EditorIndex).PKReq
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "TitleEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub TitleEditorSave()
    Dim I As Long
    
    ' Subscript out of range
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_TITLES
        If Title_Changed(I) Then
            Call SendSaveTitle(I)
        End If
    Next
    
    'Unload frmEditor_Title
    'Editor = 0
    ClearChanged_Title
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "TitleEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub TitleEditorCancel()
    ' Subscript out of range
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Editor = 0
    ClearChanged_Title
    ClearTitles
    SendRequestTitles
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "TitleEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Title()
    ' Subscript out of range
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ZeroMemory Title_Changed(1), MAX_TITLES * 2 ' 2 = Boolean length
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Title", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MoralEditorSave()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_MORALS
        If Moral_Changed(I) Then
            Call SendSaveMoral(I)
        End If
    Next
    
    'Unload frmEditor_Moral
    'Editor = 0
    ClearChanged_Moral
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MoralEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MoralEditorCancel()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Editor = 0
    ClearChanged_Moral
    ClearMorals
    SendRequestMorals
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MoralEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Moral()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ZeroMemory Moral_Changed(1), MAX_MORALS * 2 ' 2 = Boolean length
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Moral", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MoralEditorInit()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorIndex = frmEditor_Moral.lstIndex.ListIndex + 1
    Moral_Changed(EditorIndex) = True
    
    With frmEditor_Moral
        .txtName.text = Trim$(Moral(EditorIndex).Name)
        If Moral(EditorIndex).Color = 0 Then Moral(EditorIndex).Color = 1
        .scrlColor.Value = Moral(EditorIndex).Color
        .chkCanCast.Value = Moral(EditorIndex).CanCast
        .chkCanPK.Value = Moral(EditorIndex).CanPK
        .chkCanUseItem.Value = Moral(EditorIndex).CanUseItem
        .chkDropItems.Value = Moral(EditorIndex).DropItems
        .chkLoseExp.Value = Moral(EditorIndex).LoseExp
        .chkCanPickupItem.Value = Moral(EditorIndex).CanPickupItem
        .chkCanDropItem.Value = Moral(EditorIndex).CanDropItem
        .chkPlayerBlocked.Value = Moral(EditorIndex).PlayerBlocked
    End With
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "MoralEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClassEditorSave()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_CLASSES
        If Class_Changed(I) Then
            Call SendSaveClass(I)
        End If
    Next
    
    'Unload frmEditor_Class
    'Editor = 0
    ClearChanged_Class
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClassEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClassEditorCancel()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Editor = 0
    ClearChanged_Class
    ClearClasses
    SendRequestClasses
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClassEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Class()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ZeroMemory Class_Changed(1), MAX_CLASSES * 2 ' 2 = Boolean length
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Class", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClassEditorInit()
    Dim I As Long

    ' Check if the form is visible if not then exit
    If frmEditor_Class.Visible = False Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorIndex = frmEditor_Class.lstIndex.ListIndex + 1
    Class_Changed(EditorIndex) = True
    
    With frmEditor_Class
        .txtName.text = Trim$(Class(EditorIndex).Name)
        .chkSwapGender = 0
        
        .chkLocked.Value = Class(EditorIndex).Locked
        
        ' Start Item
        .scrlStartItem.Value = 1
        .scrlStartSpell.Value = 1
        .scrlItemNum.Value = Class(EditorIndex).StartItem(1)
        .scrlItemValue.Value = Class(EditorIndex).StartItemValue(1)
        
        ' Start Spell
        .scrlStartSpell.Value = 1
        .scrlStartSpell.Value = 1
        .scrlSpellNum.Value = Class(EditorIndex).StartSpell(1)
        
        ' Color
         .cmbColor.ListIndex = Class(EditorIndex).Color

        ' Start position
        If Class(EditorIndex).Map = 0 Then Class(EditorIndex).Map = 1
        .scrlMap.Value = Class(EditorIndex).Map
        .scrlX.Value = Class(EditorIndex).X
        .scrlY.Value = Class(EditorIndex).Y
        .scrlDir.Value = Class(EditorIndex).Dir
        
        .chkAnimated = Class(EditorIndex).Animated
        
        ' Loop for stats
        For I = 1 To Stats.Stat_Count - 1
            If Class(EditorIndex).Stat(I) < 1 Then Class(EditorIndex).Stat(I) = 1
            .scrlStat(I).Value = Class(EditorIndex).Stat(I)
        Next
        
        ' Set visibility on
        .scrlFFace.Visible = True
        .scrlFSprite.Visible = True
        .lblFFace.Visible = True
        .lblFSprite.Visible = True
        .scrlMFace.Visible = True
        .scrlMSprite.Visible = True
        .lblMFace.Visible = True
        .lblMSprite.Visible = True
        
        ' Set combat tree
        If Class(EditorIndex).CombatTree = 0 Then Class(EditorIndex).CombatTree = 1
        .scrlCombatTree.Value = Class(EditorIndex).CombatTree
        
        ' Sprites
        If Class(EditorIndex).MaleSprite > NumCharacters Then Class(EditorIndex).MaleSprite = NumCharacters
        .scrlMSprite.Value = Class(EditorIndex).MaleSprite
        
        If Class(EditorIndex).FemaleSprite > NumCharacters Then Class(EditorIndex).FemaleSprite = NumCharacters
        .scrlFSprite.Value = Class(EditorIndex).FemaleSprite
        
        ' Faces
        If Class(EditorIndex).MaleFace > NumFaces Then Class(EditorIndex).MaleFace = NumFaces
        .scrlMFace.Value = Class(EditorIndex).MaleFace
        
        If Class(EditorIndex).FemaleFace > NumFaces Then Class(EditorIndex).FemaleFace = NumFaces
        .scrlFFace.Value = Class(EditorIndex).FemaleFace
        
        ' Set visibility off for female gender
        .scrlFFace.Visible = False
        .scrlFSprite.Visible = False
        .lblFFace.Visible = False
        .lblFSprite.Visible = False
    End With
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "ClassEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub UpdateSpellScrollBars()
    With frmEditor_Item
        ' Don't set it if it isn't visible
        If .fraSpell.Visible = False Then Exit Sub
        
        If .scrlSpell.Value = 0 Then
            .lblSpellName.Caption = "Name: None"
        Else
            .lblSpellName.Caption = "Name: " & Trim$(Spell(.scrlSpell.Value).Name)
        End If
        
        .lblSpell.Caption = "Spell: " & .scrlSpell.Value
        Item(EditorIndex).Data1 = .scrlSpell.Value
    End With
End Sub

' Item Spawner
Public Function populateSpecificType(ByRef tempItems() As ItemRec, ItemType As Byte) As Boolean
    Dim I As Long, counter As Long, found As Boolean
    For I = 1 To MAX_ITEMS
        If Item(I).Type = ItemType And Item(I).Pic > 0 And Len(Item(I).Name) > 0 Then
            found = True
            ReDim Preserve tempItems(counter)
            tempItems(counter) = Item(I)
            ReDim Preserve currentlyListedIndexes(counter)
            currentlyListedIndexes(counter) = I
            frmItemSpawner.itemsImageList.ListImages.Add , , LoadPictureGDIPlus(App.Path & GFX_PATH & "items\" & Item(I).Pic & GFX_EXT, False, 32, 32, 16777215)

            counter = counter + 1
        End If
    Next
    
    If found Then
        populateSpecificType = True
    End If
End Function

Public Function countFreeSlots() As Byte
    Dim I As Long, counter As Byte
    
        For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) = 0 Then
            counter = counter + 1
        End If
        countFreeSlots = counter
    Next
End Function
Public Sub SpellClassListInit()
    Dim I As Long
    
    With frmEditor_Spell
        ' Build Class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        
        For I = 1 To MAX_CLASSES
            .cmbClass.AddItem Trim$(Class(I).Name)
        Next
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
    End With
End Sub

Public Sub ItemClassReqListInit()
    Dim I As Long
    
    ' Build cmbClassReq
    frmEditor_Item.cmbClassReq.Clear
    frmEditor_Item.cmbClassReq.AddItem "None"

    For I = 1 To MAX_CLASSES
        frmEditor_Item.cmbClassReq.AddItem Trim$(Class(I).Name)
    Next
    frmEditor_Item.cmbClassReq.ListIndex = Item(EditorIndex).ClassReq
End Sub

' //////////////////
' // Emoticon Editor //
' /////////////////
Public Sub EmoticonEditorInit()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    With frmEditor_Emoticon
        ' Check if the form is visible if not then exit
        If .Visible = False Then Exit Sub
        
        EditorIndex = .lstIndex.ListIndex + 1
        Emoticon_Changed(EditorIndex) = True
    
        .txtCommand.text = Trim$(Emoticon(EditorIndex).Command)
        
        If Emoticon(EditorIndex).Pic > NumEmoticons Then Emoticon(EditorIndex).Pic = NumEmoticons
        .scrlEmoticon.Value = Emoticon(EditorIndex).Pic
        If NumEmoticons < .scrlEmoticon.Value Then .scrlEmoticon.Value = NumEmoticons
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EmoticonEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EmoticonEditorSave()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_EMOTICONS
        If Emoticon_Changed(I) Then
            Call SendSaveEmoticon(I)
        End If
    Next
    
    'Unload frmEditor_Emoticon
    'Editor = 0
    ClearChanged_Emoticon
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EmoticonEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EmoticonEditorCancel()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Editor = 0
    ClearChanged_Emoticon
    ClearEmoticons
    SendRequestEmoticons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EmoticonEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChanged_Emoticon()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ZeroMemory Emoticon_Changed(1), MAX_EMOTICONS * 2 ' 2 = Boolean length
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Emoticon", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub CopyEvent_Map(X As Long, Y As Long)
    Dim Count As Long, I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Count = Map.EventCount
    If Count = 0 Then Exit Sub
    
    For I = 1 To Count
        If Map.events(I).X = X And Map.events(I).Y = Y Then
            ' Copy it
            cpEvent = Map.events(I)
            Exit Sub
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CopyEvent_Map", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub PasteEvent_Map(X As Long, Y As Long)
    Dim Count As Long, I As Long, EventNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Count = Map.EventCount
    
    If Count > 0 Then
        For I = 1 To Count
            If Map.events(I).X = X And Map.events(I).Y = Y Then
                ' Already an event - paste over it
                EventNum = I
            End If
        Next
    End If
    
    ' Couldn't find one - create one
    If EventNum = 0 Then
        ' Increment count
        AddEvent X, Y, True
        EventNum = Count + 1
    End If
    
    ' Copy it
    Map.events(EventNum) = cpEvent
    
    ' Set position
    Map.events(EventNum).X = X
    Map.events(EventNum).Y = Y
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChanged_Emoticon", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DeleteEvent(X As Long, Y As Long)
    Dim Count As Long, I As Long, lowIndex As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not InMapEditor Then Exit Sub
    If FormVisible("frmEditor_Events") Then Exit Sub
    
    Count = Map.EventCount
    
    For I = 1 To Count
        If Map.events(I).X = X And Map.events(I).Y = Y Then
            ' Delete it
            ClearEvent I
            lowIndex = I
            Exit For
        End If
    Next
    
    ' Didn't found anything
    If lowIndex = 0 Then Exit Sub
    
    ' Move everything down an index
    For I = lowIndex To Count - 1
        CopyEvent I + 1, I
    Next
    
    ' Delete the last index
    ClearEvent Count
    
    ' Set the new count
    Map.EventCount = Count - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DeleteEvent", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub AddEvent(X As Long, Y As Long, Optional ByVal CancelLoad As Boolean = False)
    Dim Count As Long, PageCount As Long, I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Count = Map.EventCount + 1
    
    ' Make sure there's not already an event
    If Count - 1 > 0 Then
        For I = 1 To Count - 1
            If Map.events(I).X = X And Map.events(I).Y = Y Then
                ' Already an event - edit it
                If Not CancelLoad Then EventEditorInit I
                Exit Sub
            End If
        Next
    End If
    
    ' Increment count
    Map.EventCount = Count
    ReDim Preserve Map.events(0 To Count)
    
    ' Set the new event
    Map.events(Count).X = X
    Map.events(Count).Y = Y
    
    ' Give it a new page
    PageCount = Map.events(Count).PageCount + 1
    Map.events(Count).PageCount = PageCount
    ReDim Preserve Map.events(Count).Pages(PageCount)
    
    ' Load the editor
    If Not CancelLoad Then EventEditorInit Count
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "AddEvent", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearEvent(EventNum As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ZeroMemory(ByVal VarPtr(Map.events(EventNum)), LenB(Map.events(EventNum)))
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearEvent", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub CopyEvent(original As Long, newone As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CopyMemory ByVal VarPtr(Map.events(newone)), ByVal VarPtr(Map.events(original)), LenB(Map.events(original))
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CopyEvent", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub EventEditorInit(EventNum As Long, Optional ByVal CommonEvent As Boolean = False)
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EventNum < 1 Then
        frmEditor_Events.Visible = True
        
        If FormVisible("frmAdmin") Then
            frmEditor_Events.Move frmAdmin.Left - frmEditor_Events.Width, frmAdmin.Top
        Else
            frmEditor_Events.Move frmMain.Left + frmMain.Width - frmEditor_Events.Width, frmMain.Top
        End If
        Exit Sub
    End If
    
    If CommonEvent Then
        frmEditor_Events.fraEvents.Visible = True
    Else
        frmEditor_Events.fraEvents.Visible = False
        frmEditor_Events.InitEventEditorForm
        Editor = EDITOR_EVENTS
    End If
    
    ' Populate the cache if we need to
    If Not HasPopulated Then
        PopulateLists
    End If
    
    EditorEvent = EventNum
    
    ' Copy the event data to the temp event
    tmpEvent = Map.events(EventNum)
    
    ' Populate form
    With frmEditor_Events
        ' Set the tabs
        .tabPages.Tabs.Clear
        
        For I = 1 To tmpEvent.PageCount
            .tabPages.Tabs.Add , , str$(I)
        Next
        
        ' Variables
        .cmbPlayerVar.Clear
        .cmbPlayerVar.AddItem "None"
        
        For I = 1 To MAX_VARIABLES
            .cmbPlayerVar.AddItem I & ". " & Variables(I)
        Next
        
        ' Switches
        .cmbPlayerSwitch.Clear
        .cmbPlayerSwitch.AddItem "None"
        
        For I = 1 To MAX_SWITCHES
            .cmbPlayerSwitch.AddItem I & ". " & Switches(I)
        Next
        
        ' Items
        .cmbHasItem.Clear
        .cmbHasItem.AddItem "None"
        
        For I = 1 To MAX_ITEMS
            .cmbHasItem.AddItem I & ": " & Trim$(Item(I).Name)
        Next
        
        ' Name
        .txtName.text = Trim$(tmpEvent.Name)
        
        ' Enable delete button
        If tmpEvent.PageCount > 1 Then
            .cmdDeletePage.Enabled = True
        Else
            .cmdDeletePage.Enabled = False
        End If
        
        .cmdPastePage.Enabled = False
        
        ' Load page 1 to start off with
        curPageNum = 1
        EventEditorLoadPage curPageNum
        
        .Visible = True
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EventEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub EventEditorLoadPage(pageNum As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Populate form
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
        
        frmEditor_Events.chkGlobal.Value = tmpEvent.Global
        
        frmEditor_Events.cmbTrigger.ListIndex = .Trigger
        frmEditor_Events.chkDirFix.Value = .DirFix
        frmEditor_Events.chkHasItem.Value = .chkHasItem
        frmEditor_Events.chkPlayerVar.Value = .chkVariable
        frmEditor_Events.chkPlayerSwitch.Value = .chkSwitch
        frmEditor_Events.chkSelfSwitch.Value = .chkSelfSwitch
        frmEditor_Events.chkWalkAnim.Value = .WalkAnim
        frmEditor_Events.chkWalkThrough.Value = .WalkThrough
        frmEditor_Events.chkShowName.Value = .ShowName
        frmEditor_Events.txtPlayerVariable = .VariableCondition
        frmEditor_Events.scrlGraphic.Value = .Graphic
        
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
        
        ' Show the commands
        EventListCommands
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EventEditorLoadPage", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub EventEditorSave()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Copy the event data from the temp event
    Map.events(EditorEvent) = tmpEvent
    
    ' Unload the form
    Unload frmEditor_Events
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EventEditorSave", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EventListCommands()
Dim I As Long, CurList As Long, oldI As Long, X As Long, Indent As String, listleftoff() As Long, conditionalstage() As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmEditor_Events.lstCommands.Clear
    
    If tmpEvent.Pages(curPageNum).CommandListCount > 0 Then
        ReDim listleftoff(1 To tmpEvent.Pages(curPageNum).CommandListCount)
        ReDim conditionalstage(1 To tmpEvent.Pages(curPageNum).CommandListCount)
        
        ' Startup at 1
        CurList = 1
        X = -1
        
newlist:
        For I = 1 To tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount
            If listleftoff(CurList) > 0 Then
                If (tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(listleftoff(CurList)).Index = EventType.evCondition Or tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(listleftoff(CurList)).Index = EventType.evShowChoices) And conditionalstage(CurList) <> 0 Then
                    I = listleftoff(CurList)
                ElseIf listleftoff(CurList) >= I Then
                    I = listleftoff(CurList) + 1
                End If
            End If
            
            If I <= tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount Then
                If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Index = EventType.evCondition Then
                    X = X + 1
                    Select Case conditionalstage(CurList)
                        Case 0
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = CurList
                            EventList(X).CommandNum = I
                            
                            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Condition
                                Case 0
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data3
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1) & "] >= " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data3
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1) & "] <= " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data3
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1) & "] > " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data3
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1) & "] < " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data3
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1) & "] != " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data3
                                    End Select
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1) & "] == " & "True"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1) & "] == " & "False"
                                    End If
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Has Item [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1).Name) & "]"
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player's Class Is [" & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1).Name) & "]"
                                Case 4
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player Knows Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1).Name) & "]"
                                Case 5
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player's Level is == " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player's Level is >= " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player's Level is <= " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player's Level is > " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player's Level is < " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Player's Level is NOT " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1
                                    End Select
                                Case 6
                                    If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data2 = 0 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Self Switch [A] == " & "True"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Self Switch [B] == " & "True"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Self Switch [C] == " & "True"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Self Switch [D] == " & "True"
                                        End Select
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data2 = 1 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Self Switch [A] == " & "False"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Self Switch [B] == " & "False"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Self Switch [C] == " & "False"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Conditional Branch: Self Switch [D] == " & "False"
                                        End Select
                                    End If
                            End Select
                            
                            Indent = Indent & "       "
                            listleftoff(CurList) = I
                            conditionalstage(CurList) = 1
                            CurList = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.CommandList
                            GoTo newlist
                        Case 1
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = CurList
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid$(Indent, 1, Len(Indent) - 4) & " : " & "Else"
                            listleftoff(CurList) = I
                            conditionalstage(CurList) = 2
                            CurList = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).ConditionalBranch.ElseCommandList
                            GoTo newlist
                        Case 2
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = CurList
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid$(Indent, 1, Len(Indent) - 4) & " : " & "End Branch"
                            Indent = Mid$(Indent, 1, Len(Indent) - 7)
                            listleftoff(CurList) = I
                            conditionalstage(CurList) = 0
                    End Select
                ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Index = EventType.evShowChoices Then
                    X = X + 1
                    
                    Select Case conditionalstage(CurList)
                        Case 0
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = CurList
                            EventList(X).CommandNum = I
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Show Choices - Prompt: " & Mid$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1, 1, 20) & "..."
                            
                            Indent = Indent & "       "
                            listleftoff(CurList) = I
                            conditionalstage(CurList) = 1
                            GoTo newlist
                        Case 1
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text2) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = CurList
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid$(Indent, 1, Len(Indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text2) & "]"
                                listleftoff(CurList) = I
                                conditionalstage(CurList) = 2
                                CurList = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(CurList) = I
                                conditionalstage(CurList) = 2
                                CurList = CurList
                                GoTo newlist
                            End If
                        Case 2
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text3) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = CurList
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid$(Indent, 1, Len(Indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text3) & "]"
                                listleftoff(CurList) = I
                                conditionalstage(CurList) = 3
                                CurList = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(CurList) = I
                                conditionalstage(CurList) = 3
                                CurList = CurList
                                GoTo newlist
                            End If
                        Case 3
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text4) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = CurList
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid$(Indent, 1, Len(Indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text4) & "]"
                                listleftoff(CurList) = I
                                conditionalstage(CurList) = 4
                                CurList = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(CurList) = I
                                conditionalstage(CurList) = 4
                                CurList = CurList
                                GoTo newlist
                            End If
                        Case 4
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text5) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = CurList
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid$(Indent, 1, Len(Indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text5) & "]"
                                listleftoff(CurList) = I
                                conditionalstage(CurList) = 5
                                CurList = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data4
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(CurList) = I
                                conditionalstage(CurList) = 5
                                CurList = CurList
                                GoTo newlist
                            End If
                        Case 5
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = CurList
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid$(Indent, 1, Len(Indent) - 4) & " : " & "Branch End"
                            Indent = Mid$(Indent, 1, Len(Indent) - 7)
                            listleftoff(CurList) = I
                            conditionalstage(CurList) = 0
                    End Select
                Else
                    X = X + 1
                    ReDim Preserve EventList(X)
                    EventList(X).CommandList = CurList
                    EventList(X).CommandNum = I
                    
                    Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Index
                        Case EventType.evAddText
                            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Add Text - " & Mid$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1, 1, 20) & "... - Color: " & GetColorName(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & " - Chat Type: Player"
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Add Text - " & Mid$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1, 1, 20) & "... - Color: " & GetColorName(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & " - Chat Type: Map"
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Add Text - " & Mid$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1, 1, 20) & "... - Color: " & GetColorName(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & " - Chat Type: Global"
                            End Select
                        Case EventType.evShowText
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Show Text - " & Mid$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1, 1, 20) & "..."
                        Case EventType.evPlayerVar
                            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & "] + " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & "] - " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & "] Random Between " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & " and " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data4
                            End Select
                        Case EventType.evPlayerSwitch
                            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & "] == True"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & "] == False"
                            End If
                        Case EventType.evSelfSwitch
                            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1
                                Case 0
                                    If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Self Switch [A] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Self Switch [A] to OFF"
                                    End If
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Self Switch [B] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Self Switch [B] to OFF"
                                    End If
                                Case 2
                                    If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Self Switch [C] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Self Switch [C] to OFF"
                                    End If
                                Case 3
                                    If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Self Switch [D] to ON"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Self Switch [D] to OFF"
                                    End If
                            End Select
                        Case EventType.evExitProcess
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Exit Event Processing"
                        Case EventType.evChangeItems
                            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Item Amount of [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "] to " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Give Player " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "(s)"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Take " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "(s) from Player."
                            End If
                        Case EventType.evRestoreHP
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Restore Player HP"
                        Case EventType.evRestoreMP
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Restore Player MP"
                        Case EventType.evLevelUp
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Level Up Player"
                        Case EventType.evChangeLevel
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Level to " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1
                        Case EventType.evChangeSkills
                            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Teach Player Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Remove Player Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "]"
                            End If
                        Case EventType.evChangeClass
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Class to " & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name)
                        Case EventType.evChangeSprite
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Sprite to " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1
                        Case EventType.evChangeGender
                            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Sex to Male."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Sex to Female."
                            End If
                        Case EventType.evChangePK
                            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player PK to No."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player PK to Yes."
                            End If
                        Case EventType.evWarpPlayer
                            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data4 = 0 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & ") while retaining direction."
                            Else
                                Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data4 - 1
                                    Case DIR_UP
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & ") facing upward."
                                    Case DIR_DOWN
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & ") facing downward."
                                    Case DIR_LEFT
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & ") facing left."
                                    Case DIR_RIGHT
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Warp Player To Map: " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & ") facing right."
                                End Select
                            End If
                        Case EventType.evSetMoveRoute
                            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 <= Map.EventCount Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Move Route for Event #" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " [" & Trim$(Map.events(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "]"
                            Else
                               frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Move Route for COULD NOT FIND EVENT!"
                            End If
                        Case EventType.evPlayAnimation
                            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "]" & " on Player"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "]" & " on Event #" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & " [" & Trim$(Map.events(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3).Name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Play Animation " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "]" & " on Tile(" & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3 & "," & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data4 & ")"
                            End If
                        Case EventType.evCustomScript
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Execute Custom Script Case: " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1
                        Case EventType.evPlayBGM
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Play BGM [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1) & "]"
                        Case EventType.evFadeoutBGM
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Fadeout BGM"
                        Case EventType.evPlaySound
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Play Sound [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1) & "]"
                        Case EventType.evStopSound
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Stop Sound"
                        Case EventType.evOpenBank
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Open Bank"
                        Case EventType.evOpenShop
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Open Shop [" & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & ". " & Trim$(Shop(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1).Name) & "]"
                        Case EventType.evSetAccess
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Player Access [" & frmEditor_Events.cmbSetAccess.List(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & "]"
                        Case EventType.evGiveExp
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Give Player " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1 & " Experience."
                        Case EventType.evShowChatBubble
                            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1
                                Case TARGET_TYPE_PLAYER
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Show Chat Bubble - " & Mid$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1, 1, 20) & "... - On Player"
                                Case TARGET_TYPE_NPC
                                    If Map.NPC(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) <= 0 Then
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Show Chat Bubble - " & Mid$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1, 1, 20) & "... - On NPC [" & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) & ". ]"
                                    Else
                                        frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Show Chat Bubble - " & Mid$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1, 1, 20) & "... - On NPC [" & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) & ". " & Trim$(NPC(Map.NPC(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2)).Name) & "]"
                                    End If
                                Case TARGET_TYPE_EVENT
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Show Chat Bubble - " & Mid$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1, 1, 20) & "... - On Event [" & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) & ". " & Trim$(Map.events(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2).Name) & "]"
                            End Select
                        Case EventType.evLabel
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Label: [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1) & "]"
                        Case EventType.evGotoLabel
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Jump to Label: [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Text1) & "]"
                        Case EventType.evSpawnNPC
                            If Map.NPC(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) <= 0 Then
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Spawn NPC: [" & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & ". " & "]"
                            Else
                                frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Spawn NPC: [" & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & ". " & Trim$(NPC(Map.NPC(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1)).Name) & "]"
                            End If
                        Case EventType.evFadeIn
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Fade In"
                        Case EventType.evFadeOut
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Fade Out"
                        Case EventType.evFlashWhite
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Flash White"
                        Case EventType.evSetFog
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Fog [Fog: " & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & " Speed: " & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) & " Opacity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3) & "]"
                        Case EventType.evSetWeather
                            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1
                                Case WEATHER_TYPE_NONE
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Weather [None]"
                                Case WEATHER_TYPE_RAIN
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Weather [Rain - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) & "]"
                                Case WEATHER_TYPE_SNOW
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Weather [Snow - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) & "]"
                                Case WEATHER_TYPE_SANDSTORM
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Weather [Sand Storm - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) & "]"
                                Case WEATHER_TYPE_STORM
                                    frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Weather [Storm - Intensity: " & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) & "]"
                            End Select
                        Case EventType.evSetTint
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Set Map Tint RGBA [" & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data2) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data3) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data4) & "]"
                        Case EventType.evWait
                            frmEditor_Events.lstCommands.AddItem Indent & "@>" & "Wait " & CStr(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I).Data1) & " Ms"
                        Case Else
                            'Ghost
                            X = X - 1
                            If X = -1 Then
                                ReDim EventList(0)
                            Else
                                ReDim Preserve EventList(X)
                            End If
                    End Select
                End If
            End If
        Next
        
        If CurList > 1 Then
            X = X + 1
            ReDim Preserve EventList(X)
            EventList(X).CommandList = CurList
            EventList(X).CommandNum = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount + 1
            frmEditor_Events.lstCommands.AddItem Indent & "@> "
            CurList = tmpEvent.Pages(curPageNum).CommandList(CurList).ParentList
            GoTo newlist
        End If
    End If
    
    frmEditor_Events.lstCommands.AddItem Indent & "@> "
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EventListCommands", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ListCommandAdd(S As String)
    Static X As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    frmEditor_Events.lstCommands.AddItem S
    
    ' Scrollbar
    If X < frmEditor_Events.TextWidth(S & "  ") Then
       X = frmEditor_Events.TextWidth(S & "  ")
      If frmEditor_Events.ScaleMode = vbTwips Then X = X / Screen.TwipsPerPixelX ' If twips change to pixels
      SendMessageByNum frmEditor_Events.lstCommands.hWnd, LB_SETHORIZONTALEXTENT, X, 0
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ListCommandAdd", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub AddCommand(Index As Long)
    Dim CurList As Long, I As Long, X As Long, CurSlot As Long, p As Long, oldCommandList As CommandListRec
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    End If
    
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        CurList = 1
    Else
        CurList = EventList(frmEditor_Events.lstCommands.ListIndex).CommandList
    End If
        
    If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(CurList)
    End If
    
    oldCommandList = tmpEvent.Pages(curPageNum).CommandList(CurList)
    tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount + 1
    p = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount
    
    If p <= 0 Then
        ReDim tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(0)
    Else
        ReDim tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(1 To p)
        tmpEvent.Pages(curPageNum).CommandList(CurList).ParentList = oldCommandList.ParentList
        tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount = p
        For I = 1 To p - 1
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(I) = oldCommandList.Commands(I)
        Next
    End If
    
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        CurSlot = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount
    Else
        I = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        If I < tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount Then
            For X = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount - 1 To I Step -1
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(X + 1) = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(X)
            Next
            CurSlot = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        Else
            CurSlot = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount
        End If
    End If
    
    Select Case Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtAddText_Text.text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlAddText_Colour.Value
            
            If frmEditor_Events.optAddText_Map.Value Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0
            ElseIf frmEditor_Events.optAddText_Global.Value Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1
            ElseIf frmEditor_Events.optAddText_Player.Value Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 2
            End If
        Case EventType.evCondition
            ' This is the part where the whole entire source goes to hell :D
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 2
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.CommandList = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.ElseCommandList = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.CommandList).ParentList = CurList
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.ElseCommandList).ParentList = CurList
            
            For I = 0 To 6
                If frmEditor_Events.optCondition_Index(I).Value = True Then X = I
            Next
            
            Select Case X
                Case 0 ' Player Var
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 0
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data3 = val(frmEditor_Events.txtCondition_PlayerVarCondition.text)
                Case 1 ' Player Switch
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 1
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
                Case 2 ' Has Item
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 2
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
                Case 3 'Class Is
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 3
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
                Case 4 ' Learned Skill
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 4
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
                Case 5 ' Level Is
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 5
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = val(frmEditor_Events.txtCondition_LevelAmount.text)
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
                Case 6 ' Self Switch
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 6
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
            End Select
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtShowText.text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data5 = frmEditor_Events.scrlFace.Value
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtChoicePrompt.text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text2 = frmEditor_Events.txtChoices(1).text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text3 = frmEditor_Events.txtChoices(2).text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text4 = frmEditor_Events.txtChoices(3).text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text5 = frmEditor_Events.txtChoices(4).text
            
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 4
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = tmpEvent.Pages(curPageNum).CommandListCount - 3
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = tmpEvent.Pages(curPageNum).CommandListCount - 2
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4 = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 3).ParentList = CurList
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 2).ParentList = CurList
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 1).ParentList = CurList
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount).ParentList = CurList
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data5 = frmEditor_Events.scrlFace2.Value
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            For I = 0 To 3
                If frmEditor_Events.optVariableAction(I).Value = True Then
                    Exit For
                End If
            Next
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = I
            If I = 3 Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = val(frmEditor_Events.txtVariableData(I).text)
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4 = val(frmEditor_Events.txtVariableData(I + 1).text)
            Else
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = val(frmEditor_Events.txtVariableData(I).text)
            End If
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evExitProcess
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = val(frmEditor_Events.txtChangeItemsAmount.text)
        Case EventType.evRestoreHP
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evRestoreMP
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evLevelUp
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evChangeLevel
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlChangeLevel.Value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlChangeSprite.Value
        Case EventType.evChangeGender
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            If frmEditor_Events.optChangeSexMale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 1
            End If
        Case EventType.evChangePK
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            If frmEditor_Events.optChangePKYes.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlWPMap.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.scrlWPX.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.scrlWPY.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.chkIgnoreMove.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.chkRepeatRoute.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.scrlPlayAnimTileX.Value
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4 = frmEditor_Events.scrlPlayAnimTileY.Value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlCustomScript.Value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = MusicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evFadeoutBGM
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = SoundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evStopSound
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evOpenBank
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlGiveExp.Value
        Case EventType.evShowChatBubble
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtChatbubbleText.text
            If frmEditor_Events.optChatBubbleTarget(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = TARGET_TYPE_PLAYER
            ElseIf frmEditor_Events.optChatBubbleTarget(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = TARGET_TYPE_NPC
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            ElseIf frmEditor_Events.optChatBubbleTarget(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = TARGET_TYPE_EVENT
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            End If
        Case EventType.evLabel
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtLabelName.text
        Case EventType.evGotoLabel
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtGotoLabel.text
        Case EventType.evSpawnNPC
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbSpawnNPC.ListIndex + 1
        Case EventType.evFadeIn
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evFadeOut
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evFlashWhite
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
        Case EventType.evSetFog
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.ScrlFogData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.ScrlFogData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.ScrlFogData(2).Value
        Case EventType.evSetWeather
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbWeather.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.scrlWeatherIntensity.Value
        Case EventType.evSetTint
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlMapTintData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.scrlMapTintData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.scrlMapTintData(2).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4 = frmEditor_Events.scrlMapTintData(3).Value
        Case EventType.evWait
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlWaitAmount.Value
    End Select
    
    EventListCommands
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "AddCommand", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditEventCommand()
    Dim I As Long, X As Long, Z As Long, CurList As Long, CurSlot As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    I = frmEditor_Events.lstCommands.ListIndex
    
    If I = -1 Then Exit Sub
    If I > UBound(EventList) Then Exit Sub

    CurList = EventList(I).CommandList
    CurSlot = EventList(I).CommandNum
    
    If CurList = 0 Then Exit Sub
    If CurSlot = 0 Then Exit Sub
    
    If CurList > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If CurSlot > tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount Then Exit Sub
    
    Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index
        Case EventType.evAddText
            isEdit = True
            frmEditor_Events.txtAddText_Text.text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1
            frmEditor_Events.scrlAddText_Colour.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            
            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2
                Case 0
                    frmEditor_Events.optAddText_Player.Value = True
                Case 1
                    frmEditor_Events.optAddText_Map.Value = True
                Case 2
                    frmEditor_Events.optAddText_Global.Value = True
            End Select
            
            frmEditor_Events.scrlFace.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
            
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(2).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCondition
            isEdit = True
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(7).Visible = True
            frmEditor_Events.fraCommands.Visible = False
            frmEditor_Events.ClearConditionFrame
            frmEditor_Events.optCondition_Index(tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition).Value = True
            
            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition
                Case 0
                    frmEditor_Events.cmbCondition_PlayerVarIndex.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarCompare.Enabled = True
                    frmEditor_Events.txtCondition_PlayerVarCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2
                    frmEditor_Events.txtCondition_PlayerVarCondition.text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data3
                Case 1
                    frmEditor_Events.cmbCondition_PlayerSwitch.Enabled = True
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2
                Case 2
                    frmEditor_Events.cmbCondition_HasItem.Enabled = True
                    frmEditor_Events.cmbCondition_HasItem.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 - 1
                Case 3
                    frmEditor_Events.cmbCondition_ClassIs.Enabled = True
                    frmEditor_Events.cmbCondition_ClassIs.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 - 1
                Case 4
                    frmEditor_Events.cmbCondition_LearntSkill.Enabled = True
                    frmEditor_Events.cmbCondition_LearntSkill.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 - 1
                Case 5
                    frmEditor_Events.cmbCondition_LevelCompare.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2
                    frmEditor_Events.cmbCondition_LevelCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1
                Case 6
                    frmEditor_Events.cmbCondition_SelfSwitch.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2
            End Select
        Case EventType.evShowText
            isEdit = True
            frmEditor_Events.txtShowText.text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(0).Visible = True
            frmEditor_Events.fraCommands.Visible = False
            frmEditor_Events.scrlFace.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data5
        Case EventType.evShowChoices
            isEdit = True
            frmEditor_Events.txtChoicePrompt.text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1
            frmEditor_Events.txtChoices(1).text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text2
            frmEditor_Events.txtChoices(2).text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text3
            frmEditor_Events.txtChoices(3).text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text4
            frmEditor_Events.txtChoices(4).text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text5
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(1).Visible = True
            frmEditor_Events.fraCommands.Visible = False
            frmEditor_Events.scrlFace2.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data5
        Case EventType.evPlayerVar
            isEdit = True
            frmEditor_Events.cmbVariable.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 - 1
            
            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2
                Case 0
                    frmEditor_Events.optVariableAction(0).Value = True
                    frmEditor_Events.txtVariableData(0).text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
                Case 1
                    frmEditor_Events.optVariableAction(1).Value = True
                    frmEditor_Events.txtVariableData(1).text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
                Case 2
                    frmEditor_Events.optVariableAction(2).Value = True
                    frmEditor_Events.txtVariableData(2).text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
                Case 3
                    frmEditor_Events.optVariableAction(3).Value = True
                    frmEditor_Events.txtVariableData(3).text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
                    frmEditor_Events.txtVariableData(4).text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4
            End Select
            
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(4).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayerSwitch
            isEdit = True
            frmEditor_Events.cmbSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 - 1
            frmEditor_Events.cmbPlayerSwitchSet.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(5).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSelfSwitch
            isEdit = True
            frmEditor_Events.cmbSetSelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.cmbSetSelfSwitchTo.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(6).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeItems
            isEdit = True
            frmEditor_Events.cmbChangeItemIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 - 1
            
            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0 Then
                frmEditor_Events.optChangeItemSet.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1 Then
                frmEditor_Events.optChangeItemAdd.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 2 Then
                frmEditor_Events.optChangeItemRemove.Value = True
            End If
            
            frmEditor_Events.txtChangeItemsAmount.text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(10).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeLevel
            isEdit = True
            frmEditor_Events.scrlChangeLevel.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(11).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSkills
            isEdit = True
            frmEditor_Events.cmbChangeSkills.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 - 1
            
            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0 Then
                frmEditor_Events.optChangeSkillsAdd.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1 Then
                frmEditor_Events.optChangeSkillsRemove.Value = True
            End If
            
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(12).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeClass
            isEdit = True
            frmEditor_Events.cmbChangeClass.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(13).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSprite
            isEdit = True
            frmEditor_Events.scrlChangeSprite.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(14).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeGender
            isEdit = True
            
            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 0 Then
                frmEditor_Events.optChangeSexMale.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 1 Then
                frmEditor_Events.optChangeSexFemale.Value = True
            End If
            
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(15).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangePK
            isEdit = True
            
            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 1 Then
                frmEditor_Events.optChangePKYes.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 0 Then
                frmEditor_Events.optChangePKNo.Value = True
            End If
            
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(16).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWarpPlayer
            isEdit = True
            frmEditor_Events.scrlWPMap.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.scrlWPX.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2
            frmEditor_Events.scrlWPY.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
            frmEditor_Events.cmbWarpPlayerDir.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4
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
            
            For I = 1 To Map.EventCount
                If I <> EditorEvent Then
                    frmEditor_Events.cmbEvent.AddItem Trim$(Map.events(I).Name)
                    X = X + 1
                    ListOfEvents(X) = I
                    If I = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 Then frmEditor_Events.cmbEvent.ListIndex = X
                End If
            Next
                
            IsMoveRouteCommand = True
                
            frmEditor_Events.chkIgnoreMove.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2
            frmEditor_Events.chkRepeatRoute.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
                
            TempMoveRouteCount = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).MoveRouteCount
            TempMoveRoute = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).MoveRoute
            
            For I = 1 To TempMoveRouteCount
                Select Case TempMoveRoute(I).Index
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
            frmEditor_Events.cmbPlayAnim.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 - 1
            frmEditor_Events.cmbPlayAnimEvent.Clear
            
            For I = 1 To Map.EventCount
                frmEditor_Events.cmbPlayAnimEvent.AddItem I & ". " & Trim$(Map.events(I).Name)
            Next
            
            frmEditor_Events.cmbPlayAnimEvent.ListIndex = 0
            If tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0 Then
                frmEditor_Events.optPlayAnimPlayer.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1 Then
                frmEditor_Events.optPlayAnimEvent.Value = True
                frmEditor_Events.cmbPlayAnimEvent.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 - 1
            ElseIf tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 2 Then
                frmEditor_Events.optPlayAnimTile.Value = True
                frmEditor_Events.scrlPlayAnimTileX.max = Map.MaxX
                frmEditor_Events.scrlPlayAnimTileY.max = Map.MaxY
                frmEditor_Events.scrlPlayAnimTileX.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
                frmEditor_Events.scrlPlayAnimTileY.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(20).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCustomScript
            isEdit = True
            frmEditor_Events.scrlCustomScript.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(29).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayBGM
            isEdit = True
            
            For I = 1 To UBound(MusicCache())
                If MusicCache(I) = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 Then
                    frmEditor_Events.cmbPlayBGM.ListIndex = I - 1
                End If
            Next
            
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(25).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlaySound
            isEdit = True
            
            For I = 1 To UBound(SoundCache())
                If SoundCache(I) = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 Then
                    frmEditor_Events.cmbPlaySound.ListIndex = I - 1
                End If
            Next
            
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(26).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evOpenShop
            isEdit = True
            frmEditor_Events.cmbOpenShop.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(21).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetAccess
            isEdit = True
            frmEditor_Events.cmbSetAccess.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(28).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evGiveExp
            isEdit = True
            frmEditor_Events.scrlGiveExp.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.lblGiveExp.Caption = "Give Exp: " & tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(17).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evShowChatBubble
            isEdit = True
            frmEditor_Events.txtChatbubbleText.text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1
            
            Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
                Case TARGET_TYPE_PLAYER
                    frmEditor_Events.optChatBubbleTarget(0).Value = True
                Case TARGET_TYPE_NPC
                    frmEditor_Events.optChatBubbleTarget(1).Value = True
                Case TARGET_TYPE_EVENT
                    frmEditor_Events.optChatBubbleTarget(1).Value = True
            End Select
            
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(3).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evLabel
            isEdit = True
            frmEditor_Events.txtLabelName.text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(8).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evGotoLabel
            isEdit = True
            frmEditor_Events.txtGotoLabel.text = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(9).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSpawnNPC
            isEdit = True
            frmEditor_Events.cmbSpawnNPC.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(19).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetFog
            isEdit = True
            frmEditor_Events.ScrlFogData(0).Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.ScrlFogData(1).Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2
            frmEditor_Events.ScrlFogData(2).Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(22).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetWeather
            isEdit = True
            frmEditor_Events.cmbWeather.ListIndex = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.scrlWeatherIntensity.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(23).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetTint
            isEdit = True
            frmEditor_Events.scrlMapTintData(0).Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.scrlMapTintData(1).Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2
            frmEditor_Events.scrlMapTintData(2).Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3
            frmEditor_Events.scrlMapTintData(3).Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(24).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWait
            isEdit = True
            frmEditor_Events.scrlWaitAmount.Value = tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(27).Visible = True
            frmEditor_Events.fraCommands.Visible = False
    End Select
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditEventCommand", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DeleteEventCommand()
    Dim I As Long, X As Long, Z As Long, CurList As Long, CurSlot As Long, p As Long, oldCommandList As CommandListRec
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    I = frmEditor_Events.lstCommands.ListIndex
    
    If I = -1 Then Exit Sub
    If I > UBound(EventList) Then Exit Sub
    
    CurList = EventList(I).CommandList
    CurSlot = EventList(I).CommandNum
    
    If CurList = 0 Then Exit Sub
    If CurSlot = 0 Then Exit Sub
    
    If CurList > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If CurSlot > tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount Then Exit Sub
    
    If CurSlot = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount Then
        tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(0)
        Else
            oldCommandList = tmpEvent.Pages(curPageNum).CommandList(CurList)
            ReDim tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(p)
            X = 1
            tmpEvent.Pages(curPageNum).CommandList(CurList).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount = p
            For I = 1 To p + 1
                If I <> CurSlot Then
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(X) = oldCommandList.Commands(I)
                    X = X + 1
                End If
            Next
        End If
    Else
        tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount
        oldCommandList = tmpEvent.Pages(curPageNum).CommandList(CurList)
        X = 1
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(0)
        Else
            ReDim tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(p)
            tmpEvent.Pages(curPageNum).CommandList(CurList).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount = p
            For I = 1 To p + 1
                If I <> CurSlot Then
                    tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(X) = oldCommandList.Commands(I)
                    X = X + 1
                End If
            Next
        End If
    End If
    
    EventListCommands
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DeleteEventCommand", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearEventCommands()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    tmpEvent.Pages(curPageNum).CommandListCount = 1
    EventListCommands
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearEventCommands", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditCommand()
    Dim I As Long, X As Long, Z As Long, CurList As Long, CurSlot As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    I = frmEditor_Events.lstCommands.ListIndex
    If I = -1 Then Exit Sub
    
    If I > UBound(EventList) Then Exit Sub
    
    CurList = EventList(I).CommandList
    CurSlot = EventList(I).CommandNum
    
    If CurList = 0 Then Exit Sub
    If CurSlot = 0 Then Exit Sub
    
    If CurList > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If CurSlot > tmpEvent.Pages(curPageNum).CommandList(CurList).CommandCount Then Exit Sub
    
    Select Case tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtAddText_Text.text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlAddText_Colour.Value
            
            If frmEditor_Events.optAddText_Player.Value Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0
            ElseIf frmEditor_Events.optAddText_Map.Value Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1
            ElseIf frmEditor_Events.optAddText_Global.Value Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 2
            End If
        Case EventType.evCondition
            If frmEditor_Events.optCondition_Index(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 0
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data3 = val(frmEditor_Events.txtCondition_PlayerVarCondition.text)
            ElseIf frmEditor_Events.optCondition_Index(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 1
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 2
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(3).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 3
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(4).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 4
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(5).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 5
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = val(frmEditor_Events.txtCondition_LevelAmount.text)
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(6).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Condition = 6
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
            End If
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtShowText.text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data5 = frmEditor_Events.scrlFace.Value
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtChoicePrompt.text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text2 = frmEditor_Events.txtChoices(1).text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text3 = frmEditor_Events.txtChoices(2).text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text4 = frmEditor_Events.txtChoices(3).text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text5 = frmEditor_Events.txtChoices(4).text
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data5 = frmEditor_Events.scrlFace2.Value
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            For I = 0 To 3
                If frmEditor_Events.optVariableAction(I).Value = True Then
                    Exit For
                End If
            Next
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = I
            If I = 3 Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = val(frmEditor_Events.txtVariableData(I).text)
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4 = val(frmEditor_Events.txtVariableData(I + 1).text)
            Else
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = val(frmEditor_Events.txtVariableData(I).text)
            End If
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = val(frmEditor_Events.txtChangeItemsAmount.text)
        Case EventType.evChangeLevel
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlChangeLevel.Value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlChangeSprite.Value
        Case EventType.evChangeGender
            If frmEditor_Events.optChangeSexMale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 1
            End If
        Case EventType.evChangePK
            If frmEditor_Events.optChangePKYes.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlWPMap.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.scrlWPX.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.scrlWPY.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.chkIgnoreMove.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.chkRepeatRoute.Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.scrlPlayAnimTileX.Value
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4 = frmEditor_Events.scrlPlayAnimTileY.Value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlCustomScript.Value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = MusicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = SoundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlGiveExp.Value
        Case EventType.evShowChatBubble
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtChatbubbleText.text
            If frmEditor_Events.optChatBubbleTarget(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = TARGET_TYPE_PLAYER
            ElseIf frmEditor_Events.optChatBubbleTarget(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = TARGET_TYPE_NPC
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            ElseIf frmEditor_Events.optChatBubbleTarget(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = TARGET_TYPE_EVENT
                tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            End If
        Case EventType.evLabel
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtLabelName.text
        Case EventType.evGotoLabel
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Text1 = frmEditor_Events.txtGotoLabel.text
        Case EventType.evSpawnNPC
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbSpawnNPC.ListIndex + 1
        Case EventType.evSetFog
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.ScrlFogData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.ScrlFogData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.ScrlFogData(2).Value
        Case EventType.evSetWeather
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.cmbWeather.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.scrlWeatherIntensity.Value
        Case EventType.evSetTint
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlMapTintData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data2 = frmEditor_Events.scrlMapTintData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data3 = frmEditor_Events.scrlMapTintData(2).Value
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data4 = frmEditor_Events.scrlMapTintData(3).Value
        Case EventType.evWait
            tmpEvent.Pages(curPageNum).CommandList(CurList).Commands(CurSlot).Data1 = frmEditor_Events.scrlWaitAmount.Value
    End Select
    
    EventListCommands
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditCommand", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Button = vbLeftButton Then
        ' Change selected shape for autotiles
        If frmEditor_Map.scrlAutotile.Value > 0 Then
            Select Case frmEditor_Map.scrlAutotile.Value
                Case 1 ' Autotile
                    EditorTileWidth = 2
                    EditorTileHeight = 3
                Case 2 ' Fake autotile
                    EditorTileWidth = 1
                    EditorTileHeight = 1
                Case 3 ' Animated
                    EditorTileWidth = 6
                    EditorTileHeight = 3
                Case 4 ' Cliff
                    EditorTileWidth = 2
                    EditorTileHeight = 2
                Case 5 ' Waterfall
                    EditorTileWidth = 2
                    EditorTileHeight = 3
            End Select
        Else
            EditorTileHeight = 1
            EditorTileWidth = 1
        End If
        
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y

        ' Random tile
        If frmEditor_Map.chkRandom.Value = 1 Then
            RandomTile(RandomTileSelected) = EditorTileY * (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / 32) + EditorTileX
            RandomTileSheet(RandomTileSelected) = frmEditor_Map.scrlTileSet.Value
            Exit Sub
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SetMapAutotileScrollbar()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Map.scrlAutotile.Value > 0 Then
        Select Case frmEditor_Map.scrlAutotile.Value
            Case 1 ' Autotile (VX)
                EditorTileWidth = 2
                EditorTileHeight = 3
            Case 2 ' Fake autotile (VX)
                EditorTileWidth = 1
                EditorTileHeight = 1
            Case 3 ' Animated (VX)
                EditorTileWidth = 6
                EditorTileHeight = 3
            Case 4 ' Cliff (VX)
                EditorTileWidth = 2
                EditorTileHeight = 2
            Case 5 ' Waterfall (VX)
                EditorTileWidth = 2
                EditorTileHeight = 3
        End Select
    Else
        EditorTileHeight = 1
        EditorTileWidth = 1
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetMapAutotileScrollbar", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' *********************
' ** Event Utilities **
' *********************
Public Function GetSubEventCount(ByVal Index As Long)
    GetSubEventCount = 0
    
    If Index <= 0 Or Index > MAX_EVENTS Then Exit Function
    
    If events(Index).HasSubEvents Then
        GetSubEventCount = UBound(events(Index).SubEvents)
    End If
End Function

Public Sub MapEditorEyeDropper()
    Dim TileLeft As Single
    Dim TileTop As Single
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If Not IsInBounds Then Exit Sub

    With Map.Tile(CurX, CurY)
        If .Layer(CurrentLayer).Tileset > 0 Then
            frmEditor_Map.scrlTileSet.Value = .Layer(CurrentLayer).Tileset
        Else
            frmEditor_Map.scrlTileSet.Value = 1
        End If
        
        TileTop = .Layer(CurrentLayer).Y * PIC_Y
        TileLeft = .Layer(CurrentLayer).X * PIC_X
        
        Call MapEditorChooseTile(vbLeftButton, TileLeft, TileTop)
   
        frmEditor_Map.scrlAutotile.Value = .Autotile(CurrentLayer)
        frmMain.chkEyeDropper.Value = False
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorEyeDropper", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub InitAdminPanel()
    frmAdmin.Visible = True
    
    frmAdmin.Left = frmMain.Left + frmMain.Width + 145
    frmAdmin.Top = frmMain.Top
    
    If Not adminMin Then

        frmAdmin.picSizer.BorderStyle = 0
        frmAdmin.picSizer.Picture = LoadResPicture("MIN", vbResBitmap)
    Else
        frmAdmin.picSizer.BorderStyle = 0
        frmAdmin.picSizer.Picture = LoadResPicture("MAX", vbResBitmap)
        frmAdmin.reverse = False
        frmAdmin.picSizer_Click
    End If

    refreshingAdminList = True
    SendRequestPlayersOnline
End Sub

Public Sub InsertByPtr(pArray() As Byte, ByVal StartIndex As Long, Optional ByVal NumElements As Long = 1)
    Dim lSize As Long, lBase As Long
    Dim lMemSize As Long
    Dim tempBytes() As Byte

    If NumElements < 1 Then Exit Sub

    lBase = LBound(pArray)
    lSize = (UBound(pArray) - (lBase - 1))

    If StartIndex < lBase Or (StartIndex + NumElements) > lSize Then
        Err.Raise 9
    ElseIf (StartIndex + NumElements) = lSize Then
        ReDim Preserve pArray(lBase To (lSize - lBase - 1) + NumElements)
        Exit Sub
    End If

    lMemSize = LenB(pArray(lBase))

    ReDim tempBytes(1 To (NumElements * lMemSize)) As Byte

    ReDim Preserve pArray(lBase To (lSize - lBase - 1) + NumElements)
    

    Call CopyMemory(ByVal VarPtr(pArray(StartIndex)) + (NumElements * lMemSize), _
                    ByVal VarPtr(pArray(StartIndex)), _
                    (lSize - StartIndex) * lMemSize)

    Call CopyMemory(ByVal VarPtr(pArray(StartIndex)), tempBytes(1), NumElements * lMemSize)
End Sub

Public Sub DeleteByPtr(pArray() As Byte, ByVal StartIndex As Long, Optional ByVal NumElements As Long = 1)
    Dim lSize As Long, lBase As Long
    Dim lMemSize As Long
    Dim tempBytes() As Byte

    If NumElements < 1 Then Exit Sub

    lBase = LBound(pArray)
    lSize = (UBound(pArray) - (lBase - 1))

    If StartIndex < lBase Or (StartIndex + NumElements) > lSize Then
        Err.Raise 9
    ElseIf (StartIndex + NumElements) = lSize Then
        ReDim Preserve pArray(lBase To (lSize - lBase - NumElements - 1))
        Exit Sub
    End If

    lMemSize = LenB(pArray(lBase))

    ReDim tempBytes(1 To (lMemSize * NumElements)) As Byte

    Call CopyMemory(ByVal VarPtr(tempBytes(1)), _
                    ByVal VarPtr(pArray(StartIndex)), _
                    lMemSize * NumElements)

    Call CopyMemory(ByVal VarPtr(pArray(StartIndex)), _
                    ByVal VarPtr(pArray(StartIndex)) + (NumElements * lMemSize), _
                    (lSize - (StartIndex + 1)) * lMemSize)

    Call CopyMemory(ByVal VarPtr(pArray(lSize - lBase - 1)) - (lMemSize * (NumElements - 1)), tempBytes(1), lMemSize * NumElements)

    ReDim Preserve pArray(lBase To (lSize - lBase - NumElements - 1))
End Sub
Public Function ArrayIsInitialized(arr) As Boolean

  Dim memVal As Long

  CopyMemory memVal, ByVal VarPtr(arr) + 8, ByVal 4 'get pointer to array
  CopyMemory memVal, ByVal memVal, ByVal 4  'see if it points to an address...
  ArrayIsInitialized = (memVal <> 0)        '...if it does, array is intialized

End Function

Public Function getItemType(itype As Byte) As String
  Select Case itype
        
            Case 0
                getItemType = "None"
            Case 1
                getItemType = "Equipment"
            Case 2
                getItemType = "Consumable"
            Case 3
                getItemType = "Title"
            Case 4
                getItemType = "Spell"
            Case 5
                getItemType = "Teleport"
            Case 6
                getItemType = "Reset Stats"
            Case 7
                getItemType = "Auto Life"
            Case 8
                getItemType = "Change Sprite"
            Case 9
                getItemType = "Recipe"
        End Select
End Function
