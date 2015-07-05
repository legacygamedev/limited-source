Attribute VB_Name = "modGameEditor"
Option Explicit

Public Sub EditorInit()
    Dim I As Long

    InEditor = True

    frmMapEditor.Show vbModeless

    EditorSet = 0

    MapEditorSelectedType = 1

    For I = 0 To 10
        If frmMapEditor.Option1(I).Value = True Then
            frmMapEditor.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles" & I & ".bmp")

            EditorSet = I
        End If
    Next I

    frmMapEditor.scrlPicture.max = Int((frmMapEditor.picBackSelect.Height - frmMapEditor.picBack.Height) / PIC_Y)
    frmMapEditor.picBack.Width = 448
End Sub

Public Sub MainMenuInit()
    frmLogin.txtName.Text = Trim$(ReadINI("CONFIG", "Account", App.Path & "\config.ini"))
    frmLogin.txtPassword.Text = Trim$(ReadINI("CONFIG", "Password", App.Path & "\config.ini"))

    If frmLogin.Check1.Value = 0 Then
        frmLogin.Check2.Value = 0
    End If

    If ConnectToServer = True And AutoLogin = 1 Then
        frmMainMenu.picAutoLogin.Visible = True
        frmChars.Label1.Visible = True
    Else
        frmMainMenu.picAutoLogin.Visible = False
        frmChars.Label1.Visible = False
    End If
End Sub

Public Sub ParseNews()
    Dim FileData As String
    Dim FileTitle As String
    Dim FileBody As String
    Dim RED As Integer
    Dim BLUE As Integer
    Dim GRN As Integer

    FileData = ReadINI("DATA", "News", App.Path & "\News.ini")
    FileTitle = Replace(FileData, "*", vbNewLine)

    FileData = ReadINI("DATA", "Desc", App.Path & "\News.ini")
    FileBody = Replace(FileData, "*", vbNewLine)

    frmMainMenu.picNews.Caption = FileTitle & vbNewLine & vbNewLine & FileBody

    RED = Val(ReadINI("COLOR", "Red", App.Path & "\News.ini"))
    GRN = Val(ReadINI("COLOR", "Green", App.Path & "\News.ini"))
    BLUE = Val(ReadINI("COLOR", "Blue", App.Path & "\News.ini"))

    If RED < 0 Or RED > 255 Or GRN < 0 Or GRN > 255 Or BLUE < 0 Or BLUE > 255 Then
        frmMainMenu.picNews.ForeColor = RGB(255, 255, 255)
    Else
        frmMainMenu.picNews.ForeColor = RGB(RED, GRN, BLUE)
    End If
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, X As Long, y As Long)
    Dim x2 As Long, y2 As Long, PicX As Long

    If InEditor Then

        If frmMapEditor.MousePointer = 2 Then
            If MapEditorSelectedType = 1 Then
                With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                    If frmMapEditor.optGround.Value = True Then
                        PicX = .Ground
                        EditorSet = .GroundSet
                    End If
                    If frmMapEditor.optMask.Value = True Then
                        PicX = .Mask
                        EditorSet = .MaskSet
                    End If
                    If frmMapEditor.optAnim.Value = True Then
                        PicX = .Anim
                        EditorSet = .AnimSet
                    End If
                    If frmMapEditor.optMask2.Value = True Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    End If
                    If frmMapEditor.optM2Anim.Value = True Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    End If
                    If frmMapEditor.optFringe.Value = True Then
                        PicX = .Fringe
                        EditorSet = .FringeSet
                    End If
                    If frmMapEditor.optFAnim.Value = True Then
                        PicX = .FAnim
                        EditorSet = .FAnimSet
                    End If
                    If frmMapEditor.optFringe2.Value = True Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    End If
                    If frmMapEditor.optF2Anim.Value = True Then
                        PicX = .F2Anim
                        EditorSet = .F2AnimSet
                    End If

                    EditorTileY = Int(PicX / TilesInSheets)
                    EditorTileX = (PicX - Int(PicX / TilesInSheets) * TilesInSheets)
                    frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                    frmMapEditor.shpSelected.Height = PIC_Y
                    frmMapEditor.shpSelected.Width = PIC_X
                End With
                
            ElseIf MapEditorSelectedType = 3 Then
                EditorTileY = Int(Map(GetPlayerMap(MyIndex)).Tile(X, y).light / TilesInSheets)
                EditorTileX = (Map(GetPlayerMap(MyIndex)).Tile(X, y).light - Int(Map(GetPlayerMap(MyIndex)).Tile(X, y).light / TilesInSheets) * TilesInSheets)
                frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                frmMapEditor.shpSelected.Height = PIC_Y
                frmMapEditor.shpSelected.Width = PIC_X
                
            ElseIf MapEditorSelectedType = 2 Then
                With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                    If .Type = TILE_TYPE_BLOCKED Then
                        frmMapEditor.optBlocked.Value = True
                    End If
                    If .Type = TILE_TYPE_WALKTHRU Then
                        frmMapEditor.optWalkThru.Value = True
                    End If
                    If .Type = TILE_TYPE_WARP Then
                        EditorWarpMap = .Data1
                        EditorWarpX = .Data2
                        EditorWarpY = .Data3
                        frmMapEditor.optWarp.Value = True
                    End If
                    If .Type = TILE_TYPE_HEAL Then
                        frmMapEditor.optHeal.Value = True
                    End If
                    If .Type = TILE_TYPE_ROOFBLOCK Then
                        frmMapEditor.optRoofBlock.Value = True
                        RoofId = .String1
                    End If
                    If .Type = TILE_TYPE_ROOF Then
                        frmMapEditor.optRoof.Value = True
                        RoofId = .String1
                    End If
                    If .Type = TILE_TYPE_KILL Then
                        frmMapEditor.optKill.Value = True
                    End If
                    If .Type = TILE_TYPE_ITEM Then
                        ItemEditorNum = .Data1
                        ItemEditorValue = .Data2
                        frmMapEditor.optItem.Value = True
                    End If
                    If .Type = TILE_TYPE_NPCAVOID Then
                        frmMapEditor.optNpcAvoid.Value = True
                    End If
                    If .Type = TILE_TYPE_KEY Then
                        KeyEditorNum = .Data1
                        KeyEditorTake = .Data2
                        frmMapEditor.optKey.Value = True
                    End If
                    If .Type = TILE_TYPE_KEYOPEN Then
                        KeyOpenEditorX = .Data1
                        KeyOpenEditorY = .Data2
                        KeyOpenEditorMsg = .String1
                        frmMapEditor.optKeyOpen.Value = True
                    End If
                    If .Type = TILE_TYPE_SHOP Then
                        EditorShopNum = .Data1
                        frmMapEditor.optShop.Value = True
                    End If
                    If .Type = TILE_TYPE_CBLOCK Then
                        EditorItemNum1 = .Data1
                        EditorItemNum2 = .Data2
                        EditorItemNum3 = .Data3
                        frmMapEditor.optCBlock.Value = True
                    End If
                    If .Type = TILE_TYPE_ARENA Then
                        Arena1 = .Data1
                        Arena2 = .Data2
                        Arena3 = .Data3
                        frmMapEditor.optArena.Value = True
                    End If
                    If .Type = TILE_TYPE_SOUND Then
                        SoundFileName = .String1
                        frmMapEditor.optSound.Value = True
                    End If
                    If .Type = TILE_TYPE_SPRITE_CHANGE Then
                        SpritePic = .Data1
                        SpriteItem = .Data2
                        SpritePrice = .Data3
                        frmMapEditor.optSprite.Value = True
                    End If
                    If .Type = TILE_TYPE_SIGN Then
                        SignLine1 = .String1
                        SignLine2 = .String2
                        SignLine3 = .String3
                        frmMapEditor.optSign.Value = True
                    End If
                    If .Type = TILE_TYPE_DOOR Then
                        frmMapEditor.optDoor.Value = True
                    End If
                    If .Type = TILE_TYPE_NOTICE Then
                        NoticeTitle = .String1
                        NoticeText = .String2
                        NoticeSound = .String3
                        frmMapEditor.optNotice.Value = True
                    End If
                    If .Type = TILE_TYPE_CHEST Then
                        frmMapEditor.optChest.Value = True
                    End If
                    If .Type = TILE_TYPE_CLASS_CHANGE Then
                        ClassChange = .Data1
                        ClassChangeReq = .Data2
                        frmMapEditor.optClassChange.Value = True
                    End If
                    If .Type = TILE_TYPE_SCRIPTED Then
                        ScriptNum = .Data1
                        frmMapEditor.optScripted.Value = True
                    End If
                    If .Type = TILE_TYPE_HOUSE Then
                        HouseItem = .Data1
                        HousePrice = .Data2
                        frmMapEditor.optHouse.Value = True
                    End If
                    If .Type = TILE_TYPE_GUILDBLOCK Then
                        GuildBlock = .Data1
                        frmMapEditor.optGuildBlock.Value = True
                    End If
                    If .Type = TILE_TYPE_BANK Then
                        frmMapEditor.optBank.Value = True
                    End If
                    If .Type = TILE_TYPE_HOOKSHOT Then
                        frmMapEditor.OptGHook.Value = True
                    End If
                    If .Type = TILE_TYPE_ONCLICK Then
                        ClickScript = .Data1
                        frmMapEditor.optClick.Value = True
                    End If
                    If .Type = TILE_TYPE_LOWER_STAT Then
                        MinusHp = .Data1
                        MinusMp = .Data2
                        MinusSp = .Data3
                        MessageMinus = .String1
                        frmMapEditor.optMinusStat.Value = True
                    End If
                End With
            End If
            frmMapEditor.MousePointer = 1
            frmStable.MousePointer = 1
        Else
            If (Button = 1) And (X >= 0) And (X <= MAX_MAPX) And (y >= 0) And (y <= MAX_MAPY) Then
                If frmMapEditor.shpSelected.Height <= PIC_Y And frmMapEditor.shpSelected.Width <= PIC_X Then
                    If MapEditorSelectedType = 1 Then
                        With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                            If frmMapEditor.optGround.Value = True Then
                                .Ground = EditorTileY * TilesInSheets + EditorTileX
                                .GroundSet = EditorSet
                            End If
                            If frmMapEditor.optMask.Value = True Then
                                .Mask = EditorTileY * TilesInSheets + EditorTileX
                                .MaskSet = EditorSet
                            End If
                            If frmMapEditor.optAnim.Value = True Then
                                .Anim = EditorTileY * TilesInSheets + EditorTileX
                                .AnimSet = EditorSet
                            End If
                            If frmMapEditor.optMask2.Value = True Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            End If
                            If frmMapEditor.optM2Anim.Value = True Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe.Value = True Then
                                .Fringe = EditorTileY * TilesInSheets + EditorTileX
                                .FringeSet = EditorSet
                            End If
                            If frmMapEditor.optFAnim.Value = True Then
                                .FAnim = EditorTileY * TilesInSheets + EditorTileX
                                .FAnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe2.Value = True Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            End If
                            If frmMapEditor.optF2Anim.Value = True Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            End If
                        End With
                    ElseIf MapEditorSelectedType = 3 Then
                        Map(GetPlayerMap(MyIndex)).Tile(X, y).light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf MapEditorSelectedType = 2 Then
                        With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                            If frmMapEditor.optBlocked.Value = True Then
                                .Type = TILE_TYPE_BLOCKED
                            End If
                            If frmMapEditor.optRoofBlock.Value = True Then
                                .Type = TILE_TYPE_ROOFBLOCK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = RoofId
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optRoof.Value = True Then
                                .Type = TILE_TYPE_ROOF
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = RoofId
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optWarp.Value = True Then
                                .Type = TILE_TYPE_WARP
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If

                            If frmMapEditor.optHeal.Value = True Then
                                .Type = TILE_TYPE_HEAL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If

                            If frmMapEditor.optKill.Value = True Then
                                .Type = TILE_TYPE_KILL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                                If frmMapEditor.optItem.Value = True Then
                                    .Type = TILE_TYPE_ITEM
                                    .Data1 = ItemEditorNum
                                    .Data2 = ItemEditorValue
                                    .Data3 = 0
                                    .String1 = vbNullString
                                    .String2 = vbNullString
                                    .String3 = vbNullString
                                End If
                            End If
                            If frmMapEditor.optNpcAvoid.Value = True Then
                                .Type = TILE_TYPE_NPCAVOID
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKey.Value = True Then
                                .Type = TILE_TYPE_KEY
                                .Data1 = KeyEditorNum
                                .Data2 = KeyEditorTake
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKeyOpen.Value = True Then
                                .Type = TILE_TYPE_KEYOPEN
                                .Data1 = KeyOpenEditorX
                                .Data2 = KeyOpenEditorY
                                .Data3 = 0
                                .String1 = KeyOpenEditorMsg
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optShop.Value = True Then
                                .Type = TILE_TYPE_SHOP
                                .Data1 = EditorShopNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optCBlock.Value = True Then
                                .Type = TILE_TYPE_CBLOCK
                                .Data1 = EditorItemNum1
                                .Data2 = EditorItemNum2
                                .Data3 = EditorItemNum3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optArena.Value = True Then
                                .Type = TILE_TYPE_ARENA
                                .Data1 = Arena1
                                .Data2 = Arena2
                                .Data3 = Arena3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSound.Value = True Then
                                .Type = TILE_TYPE_SOUND
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SoundFileName
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSprite.Value = True Then
                                .Type = TILE_TYPE_SPRITE_CHANGE
                                .Data1 = SpritePic
                                .Data2 = SpriteItem
                                .Data3 = SpritePrice
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSign.Value = True Then
                                .Type = TILE_TYPE_SIGN
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SignLine1
                                .String2 = SignLine2
                                .String3 = SignLine3
                            End If
                            If frmMapEditor.optDoor.Value = True Then
                                .Type = TILE_TYPE_DOOR
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optNotice.Value = True Then
                                .Type = TILE_TYPE_NOTICE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = NoticeTitle
                                .String2 = NoticeText
                                .String3 = NoticeSound
                            End If
                            If frmMapEditor.optChest.Value = True Then
                                .Type = TILE_TYPE_CHEST
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                                If frmMapEditor.optClassChange.Value = True Then
                                    .Type = TILE_TYPE_CLASS_CHANGE
                                    .Data1 = ClassChange
                                    .Data2 = ClassChangeReq
                                    .Data3 = 0
                                    .String1 = vbNullString
                                    .String2 = vbNullString
                                    .String3 = vbNullString
                                End If
                            End If
                            If frmMapEditor.optScripted.Value = True Then
                                .Type = TILE_TYPE_SCRIPTED
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optHouse.Value = True Then
                                .Type = TILE_TYPE_HOUSE
                                .Data1 = HouseItem
                                .Data2 = HousePrice
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optGuildBlock.Value = True Then
                                .Type = TILE_TYPE_GUILDBLOCK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = GuildBlock
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optBank.Value = True Then
                                .Type = TILE_TYPE_BANK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.OptGHook.Value = True Then
                                .Type = TILE_TYPE_HOOKSHOT
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optWalkThru.Value = True Then
                                .Type = TILE_TYPE_WALKTHRU
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optClick.Value = True Then
                                .Type = TILE_TYPE_ONCLICK
                                .Data1 = ClickScript
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optMinusStat.Value = True Then
                                .Type = TILE_TYPE_LOWER_STAT
                                .Data1 = MinusHp
                                .Data2 = MinusMp
                                .Data3 = MinusSp
                                .String1 = MessageMinus
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                        End With
                    End If
                Else
                    For y2 = 0 To Int(frmMapEditor.shpSelected.Height / PIC_Y) - 1
                        For x2 = 0 To Int(frmMapEditor.shpSelected.Width / PIC_X) - 1
                            If X + x2 <= MAX_MAPX Then
                                If y + y2 <= MAX_MAPY Then
                                    If MapEditorSelectedType = 1 Then
                                        With Map(GetPlayerMap(MyIndex)).Tile(X + x2, y + y2)
                                            If frmMapEditor.optGround.Value = True Then
                                                .Ground = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .GroundSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask.Value = True Then
                                                .Mask = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .MaskSet = EditorSet
                                            End If
                                            If frmMapEditor.optAnim.Value = True Then
                                                .Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask2.Value = True Then
                                                .Mask2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask2Set = EditorSet
                                            End If
                                            If frmMapEditor.optM2Anim.Value = True Then
                                                .M2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M2AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe.Value = True Then
                                                .Fringe = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FringeSet = EditorSet
                                            End If
                                            If frmMapEditor.optFAnim.Value = True Then
                                                .FAnim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FAnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe2.Value = True Then
                                                .Fringe2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe2Set = EditorSet
                                            End If
                                            If frmMapEditor.optF2Anim.Value = True Then
                                                .F2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .F2AnimSet = EditorSet
                                            End If
                                        End With
                                    ElseIf MapEditorSelectedType = 3 Then
                                        Map(GetPlayerMap(MyIndex)).Tile(X + x2, y + y2).light = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                    End If
                                End If
                            End If
                        Next x2
                    Next y2
                End If
            End If

            If (Button = 2) And (X >= 0) And (X <= MAX_MAPX) And (y >= 0) And (y <= MAX_MAPY) Then
                If MapEditorSelectedType = 1 Then
                    With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                        If frmMapEditor.optGround.Value = True Then
                            .Ground = 0
                        End If
                        If frmMapEditor.optMask.Value = True Then
                            .Mask = 0
                        End If
                        If frmMapEditor.optAnim.Value = True Then
                            .Anim = 0
                        End If
                        If frmMapEditor.optMask2.Value = True Then
                            .Mask2 = 0
                        End If
                        If frmMapEditor.optM2Anim.Value = True Then
                            .M2Anim = 0
                        End If
                        If frmMapEditor.optFringe.Value = True Then
                            .Fringe = 0
                        End If
                        If frmMapEditor.optFAnim.Value = True Then
                            .FAnim = 0
                        End If
                        If frmMapEditor.optFringe2.Value = True Then
                            .Fringe2 = 0
                        End If
                        If frmMapEditor.optF2Anim.Value = True Then
                            .F2Anim = 0
                        End If
                    End With
                ElseIf MapEditorSelectedType = 3 Then
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).light = 0
                ElseIf MapEditorSelectedType = 2 Then
                    With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                        .Type = 0
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End With
                End If
            End If
        End If
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        EditorTileX = Int(X / PIC_X)
        EditorTileY = Int(y / PIC_Y)
    End If
    frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
End Sub

Public Sub EditorTileScroll()
    frmMapEditor.picBackSelect.Top = (frmMapEditor.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    ScreenMode = 0
    NightMode = 0
    GridMode = 0

    ' Set the type back to default.
    MapEditorSelectedType = 1

    ' Set the map controls to default.
    frmMapEditor.fraAttribs.Visible = False
    frmMapEditor.fraLayers.Visible = True
    frmMapEditor.frmtile.Visible = True

    InEditor = False
    frmMapEditor.Visible = False

    frmStable.Show
    frmMapEditor.MousePointer = 1
    frmStable.MousePointer = 1

    Call LoadMap(GetPlayerMap(MyIndex))
End Sub

Public Sub EditorClearLayer()
    Dim Choice As Integer
    Dim X As Byte
    Dim y As Byte

    ' Ground Layer
    If frmMapEditor.optGround.Value Then
        Choice = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).Ground = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).GroundSet = 0
                Next X
            Next y
        End If
    End If

    ' Mask Layer
    If frmMapEditor.optMask.Value Then
        Choice = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).Mask = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).MaskSet = 0
                Next X
            Next y
        End If
    End If

    ' Mask Animation Layer
    If frmMapEditor.optAnim.Value Then
        Choice = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).AnimSet = 0
                Next X
            Next y
        End If
    End If

    ' Mask 2 Layer
    If frmMapEditor.optMask2.Value Then
        Choice = MsgBox("Are you sure you wish to clear the mask 2 layer?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).Mask2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).Mask2Set = 0
                Next X
            Next y
        End If
    End If

    ' Mask 2 Animation layer
    If frmMapEditor.optM2Anim.Value Then
        Choice = MsgBox("Are you sure you wish to clear the mask 2 animation layer?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).M2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).M2AnimSet = 0
                Next X
            Next y
        End If
    End If

    ' Fringe Layer
    If frmMapEditor.optFringe.Value Then
        Choice = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).FringeSet = 0
                Next X
            Next y
        End If
    End If

    ' Fringe Animation Layer
    If frmMapEditor.optFAnim.Value Then
        Choice = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).FAnimSet = 0
                Next X
            Next y
        End If
    End If

    ' Fringe 2 Layer
    If frmMapEditor.optFringe2.Value Then
        Choice = MsgBox("Are you sure you wish to clear the fringe 2 layer?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).Fringe2Set = 0
                Next X
            Next y
        End If
    End If

    ' Fringe 2 Animation Layer
    If frmMapEditor.optF2Anim.Value Then
        Choice = MsgBox("Are you sure you wish to clear the fringe 2 animation layer?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).F2AnimSet = 0
                Next X
            Next y
        End If
    End If
End Sub

Public Sub EditorClearAttribs()
    Dim Choice As Integer
    Dim X As Byte
    Dim y As Byte

    Choice = MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME)

    If Choice = vbYes Then
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = 0
            Next X
        Next y
    End If
End Sub

Public Sub EmoticonEditorInit()
    frmEmoticonEditor.scrlEmoticon.max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.Value = Emoticons(EditorIndex - 1).Pic
    frmEmoticonEditor.txtCommand.Text = Trim$(Emoticons(EditorIndex - 1).Command)

    frmEmoticonEditor.picEmoticons.Picture = LoadPicture(App.Path & "\GFX\Emoticons.bmp")

    frmEmoticonEditor.Show vbModal
End Sub

Public Sub ElementEditorInit()
    frmElementEditor.txtName.Text = Trim$(Element(EditorIndex - 1).name)
    frmElementEditor.scrlStrong.Value = Element(EditorIndex - 1).Strong
    frmElementEditor.scrlWeak.Value = Element(EditorIndex - 1).Weak
    frmElementEditor.Show vbModal
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEmoticonEditor.scrlEmoticon.Value
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = vbNullString
    End If

    Call SendSaveEmoticon(EditorIndex - 1)
    Call EmoticonEditorCancel
End Sub

Public Sub ElementEditorOk()
    Element(EditorIndex - 1).name = frmElementEditor.txtName.Text
    Element(EditorIndex - 1).Strong = frmElementEditor.scrlStrong.Value
    Element(EditorIndex - 1).Weak = frmElementEditor.scrlWeak.Value
    Call SendSaveElement(EditorIndex - 1)
    Call ElementEditorCancel
End Sub

Public Sub EmoticonEditorCancel()
    InEmoticonEditor = False
    Unload frmEmoticonEditor
End Sub

Public Sub ElementEditorCancel()
    InElementEditor = False
    Unload frmElementEditor
End Sub

Public Sub ArrowEditorInit()
    frmEditArrows.scrlArrow.max = MAX_ARROWS
    If Arrows(EditorIndex).Pic = 0 Then
        Arrows(EditorIndex).Pic = 1
    End If
    frmEditArrows.scrlArrow.Value = Arrows(EditorIndex).Pic
    frmEditArrows.txtName.Text = Arrows(EditorIndex).name
    If Arrows(EditorIndex).Range = 0 Then
        Arrows(EditorIndex).Range = 1
    End If
    frmEditArrows.scrlRange.Value = Arrows(EditorIndex).Range
    If Arrows(EditorIndex).Amount = 0 Then
        Arrows(EditorIndex).Amount = 1
    End If
    frmEditArrows.scrlAmount.Value = Arrows(EditorIndex).Amount

    frmEditArrows.picArrows.Picture = LoadPicture(App.Path & "\GFX\Arrows.bmp")

    frmEditArrows.Show vbModal
End Sub

Public Sub ArrowEditorOk()
    Arrows(EditorIndex).Pic = frmEditArrows.scrlArrow.Value
    Arrows(EditorIndex).Range = frmEditArrows.scrlRange.Value
    Arrows(EditorIndex).name = frmEditArrows.txtName.Text
    Arrows(EditorIndex).Amount = frmEditArrows.scrlAmount.Value
    Call SendSaveArrow(EditorIndex)
    Call ArrowEditorCancel
End Sub

Public Sub ArrowEditorCancel()
    InArrowEditor = False
    Unload frmEditArrows
End Sub

Public Sub ItemEditorInit()
    Dim I As Long

    EditorItemY = Int(Item(EditorIndex).Pic / 6)
    EditorItemX = (Item(EditorIndex).Pic - Int(Item(EditorIndex).Pic / 6) * 6)

    frmItemEditor.scrlClassReq.max = Max_Classes

    frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\GFX\Items.bmp")

    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).name)
    frmItemEditor.txtDesc.Text = Trim$(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    frmItemEditor.txtPrice.Text = Item(EditorIndex).Price
    frmItemEditor.chkStackable.Value = Item(EditorIndex).Stackable
    frmItemEditor.chkBound.Value = Item(EditorIndex).Bound

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_NECKLACE) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        If frmItemEditor.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            frmItemEditor.fraBow.Visible = True
        End If

        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
        frmItemEditor.scrlClassReq.Value = Item(EditorIndex).ClassReq
        frmItemEditor.scrlAccessReq.Value = Item(EditorIndex).AccessReq
        frmItemEditor.scrlAddHP.Value = Item(EditorIndex).AddHP
        frmItemEditor.scrlAddMP.Value = Item(EditorIndex).AddMP
        frmItemEditor.scrlAddSP.Value = Item(EditorIndex).AddSP
        frmItemEditor.scrlAddStr.Value = Item(EditorIndex).AddSTR
        frmItemEditor.scrlAddDef.Value = Item(EditorIndex).AddDEF
        frmItemEditor.scrlAddMagi.Value = Item(EditorIndex).AddMAGI
        frmItemEditor.scrlAddSpeed.Value = Item(EditorIndex).AddSpeed
        ' frmItemEditor.scrlAddEXP.Value = Item(EditorIndex).AddEXP
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed

        If Item(EditorIndex).Data3 > 0 Then
            If Item(EditorIndex).Stackable = 1 Then
                frmItemEditor.chkBow.Value = Checked
                frmItemEditor.chkGrapple.Value = Checked
            Else
                frmItemEditor.chkBow.Value = Checked
                frmItemEditor.chkGrapple.Value = Unchecked
            End If
        Else
            frmItemEditor.chkBow.Value = Unchecked
        End If


        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.Value = Checked Then
            For I = 1 To 100
                frmItemEditor.cmbBow.addItem I & ": " & Arrows(I).name
            Next I
            frmItemEditor.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
            frmItemEditor.picBow.Top = (Arrows(Item(EditorIndex).Data3).Pic * 32) * -1
            frmItemEditor.cmbBow.Enabled = True
        Else
            frmItemEditor.cmbBow.addItem "None"
            frmItemEditor.cmbBow.ListIndex = 0
            frmItemEditor.cmbBow.Enabled = False
        End If
        frmItemEditor.chkStackable.Visible = False
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.chkStackable.Visible = True
        frmItemEditor.scrlVitalMod.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
        frmItemEditor.chkStackable.Visible = False
    Else
        frmItemEditor.fraSpell.Visible = False
    End If

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_SCRIPTED) Then
        frmItemEditor.fraScript.Visible = True
        frmItemEditor.scrlScript.Value = Item(EditorIndex).Data1
        frmItemEditor.lblScript.Caption = Item(EditorIndex).Data1
        
        frmItemEditor.chkStackable.Visible = True
    Else
        frmItemEditor.fraScript.Visible = False
    End If
    frmItemEditor.VScroll1.Value = EditorItemY
    frmItemEditor.picItems.Top = (EditorItemY) * -32
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
    Item(EditorIndex).Price = Val(frmItemEditor.txtPrice.Text)
    Item(EditorIndex).Bound = frmItemEditor.chkBound.Value

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_NECKLACE) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        If frmItemEditor.chkBow.Value = Checked Then
            If frmItemEditor.chkGrapple.Value = Checked Then
                Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
                Item(EditorIndex).Stackable = 1
            Else
                Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
                Item(EditorIndex).Stackable = 0
            End If
        Else
            Item(EditorIndex).Data3 = 0
            Item(EditorIndex).Stackable = 0
        End If
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value

        Item(EditorIndex).ClassReq = frmItemEditor.scrlClassReq.Value
        Item(EditorIndex).AccessReq = frmItemEditor.scrlAccessReq.Value

        Item(EditorIndex).AddHP = frmItemEditor.scrlAddHP.Value
        Item(EditorIndex).AddMP = frmItemEditor.scrlAddMP.Value
        Item(EditorIndex).AddSP = frmItemEditor.scrlAddSP.Value
        Item(EditorIndex).AddSTR = frmItemEditor.scrlAddStr.Value
        Item(EditorIndex).AddDEF = frmItemEditor.scrlAddDef.Value
        Item(EditorIndex).AddMAGI = frmItemEditor.scrlAddMagi.Value
        Item(EditorIndex).AddSpeed = frmItemEditor.scrlAddSpeed.Value
        Item(EditorIndex).AddEXP = frmItemEditor.scrlAddEXP.Value
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
    End If

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0

        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddSTR = 0
        Item(EditorIndex).AddDEF = 0
        Item(EditorIndex).AddMAGI = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value

    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_NONE) Then
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0

        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddSTR = 0
        Item(EditorIndex).AddDEF = 0
        Item(EditorIndex).AddMAGI = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCRIPTED) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlScript.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0

        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddSTR = 0
        Item(EditorIndex).AddDEF = 0
        Item(EditorIndex).AddMAGI = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
    End If
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_THROW) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlScript.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0

        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddSTR = 0
        Item(EditorIndex).AddDEF = 0
        Item(EditorIndex).AddMAGI = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
        Item(EditorIndex).Stackable = 0
    End If
    Call SendSaveItem(EditorIndex)
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub NpcEditorInit()
    On Error Resume Next

    frmNpcEditor.picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

    frmNpcEditor.txtName.Text = Trim$(Npc(EditorIndex).name)
    frmNpcEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.Value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.Value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.Value = Npc(EditorIndex).speed
    frmNpcEditor.scrlMAGI.Value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.Value = Npc(EditorIndex).Big
    frmNpcEditor.StartHP.Value = Npc(EditorIndex).MaxHp
    frmNpcEditor.ExpGive.Value = Npc(EditorIndex).Exp
    frmNpcEditor.scrlChance.Value = Npc(EditorIndex).ItemNPC(1).chance
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
    If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SCRIPTED Then
        frmNpcEditor.scrlScript.Value = Npc(EditorIndex).SpawnSecs
        frmNpcEditor.scrlElement.Value = Npc(EditorIndex).Element
    End If
    If Val(0 + Npc(EditorIndex).SpriteSize) = 0 Then
        frmNpcEditor.Opt32.Value = 1
        frmNpcEditor.Opt64.Value = 0
    Else
        frmNpcEditor.Opt64.Value = 1
        frmNpcEditor.Opt32.Value = 0
    End If
    If Npc(EditorIndex).SpawnTime = 0 Then
        frmNpcEditor.chkDay.Value = Checked
        frmNpcEditor.chkNight.Value = Checked
    ElseIf Npc(EditorIndex).SpawnTime = 1 Then
        frmNpcEditor.chkDay.Value = Checked
        frmNpcEditor.chkNight.Value = Unchecked
    ElseIf Npc(EditorIndex).SpawnTime = 2 Then
        frmNpcEditor.chkDay.Value = Unchecked
        frmNpcEditor.chkNight.Value = Checked
    End If

    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.Value
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    If Npc(EditorIndex).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
        Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Else
        Npc(EditorIndex).SpawnSecs = frmNpcEditor.scrlScript.Value
    End If
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.Value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.Value
    Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.Value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.Value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.Value
    Npc(EditorIndex).MaxHp = frmNpcEditor.StartHP.Value
    Npc(EditorIndex).Exp = frmNpcEditor.ExpGive.Value

    If frmNpcEditor.Opt64.Value = True Then
        Npc(EditorIndex).SpriteSize = 1
    Else
        Npc(EditorIndex).SpriteSize = 0
    End If

    If frmNpcEditor.chkDay.Value = Checked And frmNpcEditor.chkNight.Value = Checked Then
        Npc(EditorIndex).SpawnTime = 0
    ElseIf frmNpcEditor.chkDay.Value = Checked And frmNpcEditor.chkNight.Value = Unchecked Then
        Npc(EditorIndex).SpawnTime = 1
    ElseIf frmNpcEditor.chkDay.Value = Unchecked And frmNpcEditor.chkNight.Value = Checked Then
        Npc(EditorIndex).SpawnTime = 2
    End If

    Call SendSaveNPC(EditorIndex)
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorBltSprite()
    If frmNpcEditor.BigNpc.Value = Checked Then
        frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * 64
        frmNpcEditor.picSprites.Left = 3360
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, 64, 64, frmNpcEditor.picSprites.hDC, 3 * 64, frmNpcEditor.scrlSprite.Value * 64, SRCCOPY)
    Else
        frmNpcEditor.picSprites.Left = 3600

        If SpriteSize = 1 Then
            frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * 64
            Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, 64, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.Value * 64, SRCCOPY)
        Else
            frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * PIC_Y
            Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.Value * PIC_Y, SRCCOPY)
        End If
    End If
End Sub

' Initializes the shop editor
Public Sub ShopEditorInit()
    Dim I As Integer
    Dim itemN As Integer
    Dim cItemMade As Boolean

    On Error GoTo ShopEditorInit_Error


    frmShopEditor.txtName.Text = Trim$(Shop(EditorIndex).name)
    frmShopEditor.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    frmShopEditor.chkShow.Value = Shop(EditorIndex).ShowInfo
    frmShopEditor.chkSellsItems.Value = Shop(EditorIndex).BuysItems

    cItemMade = False

    frmShopEditor.cmbCurrency.Clear
    frmShopEditor.lstItems.Clear

    ' Add all the currency items to cmbCurrency
    For I = 1 To MAX_ITEMS
        If Item(I).Type = ITEM_TYPE_CURRENCY Then
            ' It's a currency item, so add it to the list
            frmShopEditor.cmbCurrency.addItem (I & " - " & Trim(Item(I).name))
            ' Add it to the item data so that we know the number
            frmShopEditor.cmbCurrency.ItemData(frmShopEditor.cmbCurrency.ListCount - 1) = I
            cItemMade = True 'we have at least 1 currency item
            If Shop(EditorIndex).currencyItem = I Then
                frmShopEditor.cmbCurrency.ListIndex = frmShopEditor.cmbCurrency.ListCount - 1
            End If
        End If
    Next I

    If Not cItemMade Then
        Call MsgBox("Please make at least one type of currency first!")
        Call ShopEditorCancel
        Exit Sub
    End If

    ' Add all the items to the list
    For I = 1 To MAX_SHOP_ITEMS
        itemN = Shop(EditorIndex).ShopItem(I).ItemNum

        ' If the item is not empty
        If itemN > 0 Then
            ' Add the item to the shop list
            Call frmShopEditor.AddShopItem(itemN, Shop(EditorIndex).ShopItem(I).Price, Shop(EditorIndex).currencyItem, Shop(EditorIndex).ShopItem(I).Amount)
        End If
    Next I

    ' Add all items to the 'add item' list
    For I = 1 To MAX_ITEMS
        frmShopEditor.cmbItemList.addItem (I & " - " & Trim(Item(I).name))
    Next I

    frmShopEditor.frmAddEditItem.Visible = False

    ' Init shop editor temp array
    frmShopEditor.LoadShopItemData (EditorIndex)

    frmShopEditor.Show vbModal

    On Error GoTo 0
    Exit Sub

ShopEditorInit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShopEditorInit of Module modGameLogic"
    ' Close the shop editor
    frmShopEditor.Visible = False
    Call ShopEditorCancel
End Sub


Public Sub ShopEditorOk()
    Dim I As Integer
    Dim currencyItem As Integer

    If frmShopEditor.cmbCurrency.ListIndex < 0 Then
        MsgBox "Please pick a currency item!", vbExclamation
        Exit Sub
    End If

    currencyItem = frmShopEditor.cmbCurrency.ItemData(frmShopEditor.cmbCurrency.ListIndex)

    Shop(EditorIndex).name = frmShopEditor.txtName.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.Value
    Shop(EditorIndex).BuysItems = frmShopEditor.chkSellsItems.Value
    Shop(EditorIndex).ShowInfo = frmShopEditor.chkShow.Value
    Shop(EditorIndex).currencyItem = currencyItem

    For I = 1 To MAX_SHOP_ITEMS
        Shop(EditorIndex).ShopItem(I).Amount = frmShopEditor.GetShopItemAmt(I)
        Shop(EditorIndex).ShopItem(I).ItemNum = frmShopEditor.GetShopItemNum(I)
        Shop(EditorIndex).ShopItem(I).Price = frmShopEditor.GetShopItemPrice(I)
    Next I

    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub SpellEditorInit()
    Dim I As Long

    frmSpellEditor.iconn.Picture = LoadPicture(App.Path & "\GFX\Icons.bmp")

    frmSpellEditor.cmbClassReq.addItem "All Classes"
    For I = 0 To Max_Classes
        frmSpellEditor.cmbClassReq.addItem Trim$(Class(I).name)
    Next I

    frmSpellEditor.txtName.Text = Trim$(Spell(EditorIndex).name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.Value = Spell(EditorIndex).LevelReq

    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    frmSpellEditor.scrlVitalMod.Value = Spell(EditorIndex).Data1

    frmSpellEditor.scrlCost.Value = Spell(EditorIndex).MPCost
    frmSpellEditor.scrlSound.Value = Spell(EditorIndex).Sound

    If Spell(EditorIndex).Range = 0 Then
        Spell(EditorIndex).Range = 1
    End If
    frmSpellEditor.scrlRange.Value = Spell(EditorIndex).Range

    frmSpellEditor.scrlSpellAnim.Value = Spell(EditorIndex).SpellAnim
    frmSpellEditor.scrlSpellTime.Value = Spell(EditorIndex).SpellTime
    frmSpellEditor.scrlSpellDone.Value = Spell(EditorIndex).SpellDone

    frmSpellEditor.chkArea.Value = Spell(EditorIndex).AE
    frmSpellEditor.chkBig.Value = Spell(EditorIndex).Big

    frmSpellEditor.scrlElement.Value = Spell(EditorIndex).Element
    frmSpellEditor.scrlElement.max = MAX_ELEMENTS

    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.Value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.Value
    Spell(EditorIndex).Data3 = 0
    Spell(EditorIndex).MPCost = frmSpellEditor.scrlCost.Value
    Spell(EditorIndex).Sound = frmSpellEditor.scrlSound.Value
    Spell(EditorIndex).Range = frmSpellEditor.scrlRange.Value

    Spell(EditorIndex).SpellAnim = frmSpellEditor.scrlSpellAnim.Value
    Spell(EditorIndex).SpellTime = frmSpellEditor.scrlSpellTime.Value
    Spell(EditorIndex).SpellDone = frmSpellEditor.scrlSpellDone.Value

    Spell(EditorIndex).AE = frmSpellEditor.chkArea.Value
    Spell(EditorIndex).Big = frmSpellEditor.chkBig.Value

    Spell(EditorIndex).Element = frmSpellEditor.scrlElement.Value

    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

'==>House editor stuff<==

Public Sub HouseEditorInit()
    Dim I As Long

    InHouseEditor = True

    frmHouseEditor.Show vbModeless

    EditorSet = 0

    HouseEditorSelectedType = 1

    frmHouseEditor.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles9.bmp")

    EditorSet = 9

    frmHouseEditor.scrlPicture.max = Int((frmHouseEditor.picBackSelect.Height - frmHouseEditor.picBack.Height) / PIC_Y)
    frmHouseEditor.picBack.Width = 448
End Sub

Public Sub HouseEditorCancel()
    ScreenMode = 0
    NightMode = 0
    GridMode = 0

    ' Set the type back to default.
    HouseEditorSelectedType = 1

    ' Set the map controls to default.
    frmHouseEditor.fraAttribs.Visible = False
    frmHouseEditor.fraLayers.Visible = True
    frmHouseEditor.frmtile.Visible = True

    InHouseEditor = False
    frmHouseEditor.Visible = False

    frmStable.Show
    frmHouseEditor.MousePointer = 1
    frmStable.MousePointer = 1

    Call LoadMap(GetPlayerMap(MyIndex))
End Sub

Public Sub HouseEditorSend()
    Call SendMap
    Call HouseEditorCancel
End Sub

Public Sub HouseEditorChooseTile(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        EditorTileX = Int(X / PIC_X)
        EditorTileY = Int(y / PIC_Y)
    End If
    frmHouseEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
    frmHouseEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
End Sub

Public Sub HouseEditorTileScroll()
    frmHouseEditor.picBackSelect.Top = (frmHouseEditor.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub HouseEditorMouseDown(Button As Integer, Shift As Integer, X As Long, y As Long)
    Dim x2 As Long, y2 As Long, PicX As Long

    If InHouseEditor Then

        If frmHouseEditor.MousePointer = 2 Then
            If HouseEditorSelectedType = 1 Then
                With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                    If frmHouseEditor.optMask2.Value = True Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    End If
                    If frmHouseEditor.optM2Anim.Value = True Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    End If
                    If frmHouseEditor.optFringe2.Value = True Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    End If
                    If frmHouseEditor.optF2Anim.Value = True Then
                        PicX = .F2Anim
                        EditorSet = .F2AnimSet
                    End If

                    EditorTileY = Int(PicX / TilesInSheets)
                    EditorTileX = (PicX - Int(PicX / TilesInSheets) * TilesInSheets)
                    frmHouseEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                    frmHouseEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                    frmHouseEditor.shpSelected.Height = PIC_Y
                    frmHouseEditor.shpSelected.Width = PIC_X
                End With
                
            ElseIf HouseEditorSelectedType = 3 Then
                EditorTileY = Int(Map(GetPlayerMap(MyIndex)).Tile(X, y).light / TilesInSheets)
                EditorTileX = (Map(GetPlayerMap(MyIndex)).Tile(X, y).light - Int(Map(GetPlayerMap(MyIndex)).Tile(X, y).light / TilesInSheets) * TilesInSheets)
                frmHouseEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                frmHouseEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                frmHouseEditor.shpSelected.Height = PIC_Y
                frmHouseEditor.shpSelected.Width = PIC_X
                
            ElseIf HouseEditorSelectedType = 2 Then
                With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                    If .Type = TILE_TYPE_BLOCKED Then
                        frmHouseEditor.optBlocked.Value = True
                    End If
                End With
            End If
            frmHouseEditor.MousePointer = 1
            frmStable.MousePointer = 1
        Else
            If (Button = 1) And (X >= 0) And (X <= MAX_MAPX) And (y >= 0) And (y <= MAX_MAPY) Then
                If frmHouseEditor.shpSelected.Height <= PIC_Y And frmHouseEditor.shpSelected.Width <= PIC_X Then
                    If HouseEditorSelectedType = 1 Then
                        With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                            If frmHouseEditor.optMask2.Value = True Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            End If
                            If frmHouseEditor.optM2Anim.Value = True Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            End If
                            If frmHouseEditor.optFringe2.Value = True Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            End If
                            If frmHouseEditor.optF2Anim.Value = True Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            End If
                        End With
                    ElseIf HouseEditorSelectedType = 3 Then
                        Map(GetPlayerMap(MyIndex)).Tile(X, y).light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf HouseEditorSelectedType = 2 Then
                        With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                            If frmHouseEditor.optBlocked.Value = True Then
                                .Type = TILE_TYPE_BLOCKED
                            End If
                        End With
                    End If
                Else
                    For y2 = 0 To Int(frmHouseEditor.shpSelected.Height / PIC_Y) - 1
                        For x2 = 0 To Int(frmHouseEditor.shpSelected.Width / PIC_X) - 1
                            If X + x2 <= MAX_MAPX Then
                                If y + y2 <= MAX_MAPY Then
                                    If HouseEditorSelectedType = 1 Then
                                        With Map(GetPlayerMap(MyIndex)).Tile(X + x2, y + y2)
                                            If frmHouseEditor.optMask2.Value = True Then
                                                .Mask2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask2Set = EditorSet
                                            End If
                                            If frmHouseEditor.optM2Anim.Value = True Then
                                                .M2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M2AnimSet = EditorSet
                                            End If
                                            If frmHouseEditor.optFringe2.Value = True Then
                                                .Fringe2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe2Set = EditorSet
                                            End If
                                            If frmHouseEditor.optF2Anim.Value = True Then
                                                .F2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .F2AnimSet = EditorSet
                                            End If
                                        End With
                                    ElseIf HouseEditorSelectedType = 3 Then
                                        Map(GetPlayerMap(MyIndex)).Tile(X + x2, y + y2).light = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                    End If
                                End If
                            End If
                        Next x2
                    Next y2
                End If
            End If

            If (Button = 2) And (X >= 0) And (X <= MAX_MAPX) And (y >= 0) And (y <= MAX_MAPY) Then
                If HouseEditorSelectedType = 1 Then
                    With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                        If frmHouseEditor.optMask2.Value = True Then
                            .Mask2 = 0
                        End If
                        If frmHouseEditor.optM2Anim.Value = True Then
                            .M2Anim = 0
                        End If
                        If frmHouseEditor.optFringe2.Value = True Then
                            .Fringe2 = 0
                        End If
                        If frmHouseEditor.optF2Anim.Value = True Then
                            .F2Anim = 0
                        End If
                    End With
                ElseIf HouseEditorSelectedType = 3 Then
                    Map(GetPlayerMap(MyIndex)).Tile(X, y).light = 0
                ElseIf HouseEditorSelectedType = 2 Then
                    With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                        .Type = 0
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End With
                End If
            End If
        End If
    End If
End Sub
