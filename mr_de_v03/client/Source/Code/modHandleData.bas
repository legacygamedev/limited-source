Attribute VB_Name = "modHandleData"
Option Explicit
 
Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CMsgAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(CMsgClientMsg) = GetAddress(AddressOf HandleClientMsg)
    HandleDataSub(CMsgAllChars) = GetAddress(AddressOf HandleAllChars)
    HandleDataSub(CMsgLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(CMsgNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(CMsgClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(CMsgInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(CMsgPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(CMsgPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(CMsgPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(CMsgPlayerVital) = GetAddress(AddressOf HandlePlayerVital)
    HandleDataSub(CMsgPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(CMsgPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(CMsgPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CMsgNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(CMsgPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CMsgNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(CMsgPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(CMsgAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CMsgNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(CMsgCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(CMsgMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CMsgMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(CMsgMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(CMsgMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(CMsgChatMsg) = GetAddress(AddressOf HandleChatMsg)
    HandleDataSub(CMsgSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CMsgItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(CMsgUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(CMsgUpdateItems) = GetAddress(AddressOf HandleUpdateItems)
    HandleDataSub(CMsgEditItem) = GetAddress(AddressOf HandleEditItem)
    HandleDataSub(CMsgEditEmoticon) = GetAddress(AddressOf HandleEditEmoticon)
    HandleDataSub(CMsgUpdateEmoticon) = GetAddress(AddressOf HandleUpdateEmoticon)
    HandleDataSub(CMsgUpdateEmoticons) = GetAddress(AddressOf HandleUpdateEmoticons)
    HandleDataSub(CMsgEmoticonEditor) = GetAddress(AddressOf HandleEmoticonEditor)
    HandleDataSub(CMsgCheckEmoticon) = GetAddress(AddressOf HandleCheckEmoticon)
    HandleDataSub(CMsgNewTarget) = GetAddress(AddressOf HandleNewTarget)
    HandleDataSub(CMsgSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(CMsgNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(CMsgNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(CMsgUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(CMsgUpdateNpcs) = GetAddress(AddressOf HandleUpdateNpcs)
    HandleDataSub(CMsgEditNpc) = GetAddress(AddressOf HandleEditNpc)
    HandleDataSub(CMsgMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(CMsgEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(CMsgShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(CMsgUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(CMsgUpdateShops) = GetAddress(AddressOf HandleUpdateShops)
    HandleDataSub(CMsgEditShop) = GetAddress(AddressOf HandleEditShop)
    HandleDataSub(CMsgSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(CMsgUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(CMsgUpdateSpells) = GetAddress(AddressOf HandleUpdateSpells)
    HandleDataSub(CMsgEditSpell) = GetAddress(AddressOf HandleEditSpell)
    HandleDataSub(CMsgAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(CMsgUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(CMsgUpdateAnimations) = GetAddress(AddressOf HandleUpdateAnimations)
    HandleDataSub(CMsgEditAnimation) = GetAddress(AddressOf HandleEditAnimation)
    HandleDataSub(CMsgTrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(CMsgSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CMsgActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(CMsgAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(CMsgPlayerGuild) = GetAddress(AddressOf HandlePlayerGuild)
    HandleDataSub(CMsgPlayerExp) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(CMsgCancelSpell) = GetAddress(AddressOf HandleCancelSpell)
    HandleDataSub(CMsgSpellReady) = GetAddress(AddressOf HandleSpellReady)
    HandleDataSub(CMsgSpellCooldown) = GetAddress(AddressOf HandleSpellCooldown)
    HandleDataSub(CMsgLeftGame) = GetAddress(AddressOf HandleLeftGame)
    HandleDataSub(CMsgPlayerDead) = GetAddress(AddressOf HandlePlayerDead)
    HandleDataSub(CMsgPlayerGold) = GetAddress(AddressOf HandlePlayerGold)
    HandleDataSub(CMsgPlayerRevival) = GetAddress(AddressOf HandlePlayerRevival)
    HandleDataSub(CMsgQuestEditor) = GetAddress(AddressOf HandleQuestEditor)
    HandleDataSub(CMsgUpdateQuest) = GetAddress(AddressOf HandleUpdateQuest)
    HandleDataSub(CMsgUpdateQuests) = GetAddress(AddressOf HandleUpdateQuests)
    HandleDataSub(CMsgEditQuest) = GetAddress(AddressOf HandleEditQuest)
    HandleDataSub(CMsgAvailableQuests) = GetAddress(AddressOf HandleAvailableQuests)
    HandleDataSub(CMsgPlayerQuests) = GetAddress(AddressOf HandlePlayerQuests)
    HandleDataSub(CMsgPlayerQuest) = GetAddress(AddressOf HandlePlayerQuest)
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long

    If EncryptPackets Then
        Encryption_XOR_DecryptByte Data(), PacketKeys(PacketInIndex)
        PacketInIndex = PacketInIndex + 1
        If PacketInIndex > PacketEncKeys - 1 Then PacketInIndex = 0
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        MsgBox "Packet Error.", vbOKOnly
        GameDestroy
        Exit Sub
    End If

    If MsgType >= CMSG_COUNT Then
        MsgBox "Packet Error.", vbOKOnly
        GameDestroy
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    frmSendGetData.Visible = False
    frmMainMenu.Visible = False
    
    Msg = Buffer.ReadString
    CurrentState = MenuStates.Shutdown

    frmEvent.Visible = True
    frmEvent.lblInformation.Caption = Msg
    
    ' Have to clear out key otherwise we wouldn't be able to reconnect to the server
    PacketInIndex = 0
    PacketOutIndex = 0
End Sub

Private Sub HandleClientMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    CurrentState = Buffer.ReadByte
    
    frmEvent.Visible = True
    frmEvent.lblInformation.Caption = Msg
End Sub

Private Sub HandleAllChars(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long, n As Long
Dim Level As Long
Dim Name As String, ClassName As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    frmChars.Visible = True
    frmSendGetData.Visible = False
    
    frmChars.lstChars.Clear
    For i = 1 To MAX_CHARS
        Name = Buffer.ReadString
        ClassName = Buffer.ReadString
        Level = Buffer.ReadLong
        
        If Trim$(Name) = vbNullString Then
            frmChars.lstChars.AddItem "Free Adventurer Slot"
        Else
            frmChars.lstChars.AddItem "[" & Level & "] " & Name & " - " & ClassName
        End If
    Next

    frmChars.lstChars.ListIndex = Val(ReadIniValue(App.Path & "\Core Files\Configuration.ini", "Account Information", "LastChar"))
End Sub

Private Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    
    frmSendGetData.Visible = True
    frmChars.Visible = False
    
    SetStatus "Receiving game data..."
End Sub

Private Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    frmNewChar.cmbclass.Clear
    MAX_CLASSES = Buffer.ReadByte
    ReDim Class(0 To MAX_CLASSES)
    For i = 0 To MAX_CLASSES
        Class(i).Name = Buffer.ReadString
        For n = 1 To Vitals.Vital_Count
            Class(i).Vital(n) = Buffer.ReadLong
        Next
        
        For n = 1 To Stats.Stat_Count
            Class(i).Stat(n) = Buffer.ReadLong
        Next
        Class(i).MaleSprite = Buffer.ReadString
        Class(i).FemaleSprite = Buffer.ReadString
        
        frmNewChar.cmbclass.AddItem Trim$(Class(i).Name)
    Next
    
    frmNewChar.cmbclass.ListIndex = 0
    
    ' Used for if the player is creating a new character
    frmNewChar.Visible = True
    frmSendGetData.Visible = False
End Sub

Private Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim n As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MAX_CLASSES = Buffer.ReadByte
    ReDim Class(0 To MAX_CLASSES)
    For i = 0 To MAX_CLASSES
        Class(i).Name = Buffer.ReadString
        For n = 1 To Vitals.Vital_Count
            Class(i).Vital(n) = Buffer.ReadLong
        Next
        
        For n = 1 To Stats.Stat_Count
            Class(i).Stat(n) = Buffer.ReadLong
        Next
    Next
End Sub

Private Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InGame = True
    GameInit
    GameLoop
End Sub

Private Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    For i = 1 To MAX_INV
        Update_InvItemNum MyIndex, i, Buffer.ReadLong
        Update_InvItemValue MyIndex, i, Buffer.ReadLong
        Update_InvItemBound MyIndex, i, Buffer.ReadByte
    Next
    
    UpdateInventory
End Sub

Private Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim InvSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    InvSlot = Buffer.ReadLong
    
    Update_InvItemNum MyIndex, InvSlot, Buffer.ReadLong
    Update_InvItemValue MyIndex, InvSlot, Buffer.ReadLong
    Update_InvItemBound MyIndex, InvSlot, Buffer.ReadByte
    UpdateInventory
End Sub

Private Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Update_EquipmentSlot MyIndex, Buffer.ReadLong, Buffer.ReadLong
    
    UpdateEquipment
End Sub

Private Sub HandlePlayerVital(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Vital As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Vital = Buffer.ReadLong
    Update_MaxVital MyIndex, Vital, Buffer.ReadLong
    Update_Vital MyIndex, Vital, Buffer.ReadLong
    
    frmMainGame.lblVital(Vital - 1).Caption = VitalName(Vital) & ": " & Current_Vital(MyIndex, Vital) & " / " & Current_MaxVital(MyIndex, Vital)
    
    Select Case Vital
        Case Vitals.HP
            frmMainGame.picHP.Width = (Current_Vital(MyIndex, Vitals.HP) / Current_MaxVital(MyIndex, Vitals.HP)) * 178
            
        Case Vitals.MP
            frmMainGame.PicMP.Width = (Current_Vital(MyIndex, Vitals.MP) / Current_MaxVital(MyIndex, Vitals.MP)) * 178
    End Select
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Update_Level MyIndex, Buffer.ReadLong
    Update_Points MyIndex, Buffer.ReadLong
    
    For i = 1 To Stats.Stat_Count
        Update_BaseStat MyIndex, i, Buffer.ReadLong
        Update_ModStat MyIndex, i, Buffer.ReadLong
    Next
    
    Update_StatsWindow
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
        
    Update_Name i, Buffer.ReadString
    Update_Class i, Buffer.ReadByte
    Update_Sprite i, Buffer.ReadLong
    Update_Map i, Buffer.ReadLong
    Update_X i, Buffer.ReadLong
    Update_Y i, Buffer.ReadLong
    Update_Dir i, Buffer.ReadLong
    Update_Access i, Buffer.ReadLong
    Update_PK i, Buffer.ReadLong
    Update_GuildName i, Buffer.ReadString
    Update_GuildAbbreviation i, Buffer.ReadString
    Update_IsDead i, Buffer.ReadByte
    Update_IsDeadTimer i, Buffer.ReadLong + GetTickCount
    
    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    
    ' Check if the player is the client player, and if so reset directions
    If i = MyIndex Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
        Update_StatsWindow
    End If
    
    ' Update our mapplayers for looping
    UpdateMapPlayers
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    Update_X i, Buffer.ReadLong
    Update_Y i, Buffer.ReadLong
    Update_Dir i, Buffer.ReadLong
    
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    Player(i).Moving = 1

    Select Case Current_Dir(i)
        Case DIR_UP
            Player(i).YOffset = PIC_Y
        Case DIR_DOWN
            Player(i).YOffset = -PIC_Y
        Case DIR_LEFT
            Player(i).XOffset = PIC_X
        Case DIR_RIGHT
            Player(i).XOffset = -PIC_X
    End Select
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    i = Buffer.ReadLong
    
    ' check it
    If i < 0 Then Exit Sub
    If i > MapNpcCount Then Exit Sub
    
    MapNpc(i).X = Buffer.ReadLong
    MapNpc(i).Y = Buffer.ReadLong
    MapNpc(i).Dir = Buffer.ReadLong
    MapNpc(i).Moving = Buffer.ReadLong
    MapNpc(i).XOffset = 0
    MapNpc(i).YOffset = 0
    
    Select Case MapNpc(i).Dir
        Case DIR_UP
            MapNpc(i).YOffset = PIC_Y
        Case DIR_DOWN
            MapNpc(i).YOffset = PIC_Y * -1
        Case DIR_LEFT
            MapNpc(i).XOffset = PIC_X
        Case DIR_RIGHT
            MapNpc(i).XOffset = PIC_X * -1
    End Select
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    Update_Dir i, Buffer.ReadLong
    
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    Player(i).Moving = 0
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    MapNpc(i).Dir = Buffer.ReadLong
    
    MapNpc(i).XOffset = 0
    MapNpc(i).YOffset = 0
    MapNpc(i).Moving = 0
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Update_X MyIndex, Buffer.ReadLong
    Update_Y MyIndex, Buffer.ReadLong
    
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).XOffset = 0
    Player(MyIndex).YOffset = 0
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
        
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
        
    i = Buffer.ReadLong
        
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount
    
    Set Buffer = Nothing
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Update_Map i, 0
        End If
    Next

    LoadMap Buffer.ReadLong
    
    SendNeedMap SaveMap.Revision
End Sub

Private Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim MapNum As Long
Dim X As Long
Dim Y As Long
Dim Buffer As clsBuffer
Dim TileSize As Long
Dim TileData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Buffer.DecompressBuffer
    
    MapNum = Buffer.ReadLong

    If MapNum <= 0 Then Exit Sub
    If MapNum > MAX_MAPS Then Exit Sub
    
    With SaveMap
        .Name = Buffer.ReadString
        .Revision = Buffer.ReadLong
        .Moral = Buffer.ReadByte
        .Up = Buffer.ReadInteger
        .Down = Buffer.ReadInteger
        .Left = Buffer.ReadInteger
        .Right = Buffer.ReadInteger
        .Music = Buffer.ReadByte
        .BootMap = Buffer.ReadInteger
        .BootX = Buffer.ReadByte
        .BootY = Buffer.ReadByte
        .TileSet = Buffer.ReadByte
        .MaxX = Buffer.ReadByte
        .MaxY = Buffer.ReadByte

        For X = 1 To MAX_MOBS
            .Mobs(X).NpcCount = Buffer.ReadLong
            ReDim .Mobs(X).Npc(.Mobs(X).NpcCount)

            If .Mobs(X).NpcCount > 0 Then
                For Y = 1 To .Mobs(X).NpcCount
                    .Mobs(X).Npc(Y) = Buffer.ReadLong
                Next
            End If
        Next

        ' set the Tile()
        ReDim .Tile(0 To .MaxX, 0 To .MaxY)
        
        TileSize = LenB(.Tile(0, 0)) * ((UBound(.Tile, 1) + 1) * (UBound(.Tile, 2) + 1))
        ReDim TileData(0 To TileSize - 1)
        TileData = Buffer.ReadBytes(TileSize)
        CopyMemory ByVal VarPtr(.Tile(0, 0)), ByVal VarPtr(TileData(0)), TileSize
    End With
    
    ' Save the map
    SaveLocalMap MapNum
    
    Map = SaveMap
    
    UpdateMapNpcCount
    ClearMapNpcs
    ClearTempTile
    
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InEditor Then
        InEditor = False
        'frmMainGame.picMapEditor.Visible = False
        frmMapEditor.Visible = False
        
        If frmMapWarp.Visible Then
            Unload frmMapWarp
        End If

        If frmMapProperties.Visible Then
            Unload frmMapProperties
        End If
    End If
End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_MAP_ITEMS
        MapItem(i).Num = Buffer.ReadLong
        MapItem(i).Value = Buffer.ReadLong
        MapItem(i).X = Buffer.ReadByte
        MapItem(i).Y = Buffer.ReadByte
    Next
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    If MapNpcCount <= 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MapNpcCount = Buffer.ReadLong
    ReDim Preserve MapNpc(MapNpcCount)
    For i = 1 To MapNpcCount
        MapNpc(i).Num = Buffer.ReadLong
        MapNpc(i).X = Buffer.ReadByte
        MapNpc(i).Y = Buffer.ReadByte
        MapNpc(i).Dir = Buffer.ReadByte
    Next
End Sub

Private Sub HandleMapDone(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
    
    ' Clear out ActionMsg and spell anims
    For i = 1 To MAX_BYTE
        ClearActionMsg i
        ClearAnim i
    Next
        
    GettingMap = False
    
    ' Play music
    If Map.Music > 0 Then
        Call PlayMidi("music" & Trim$(CStr(Map.Music)) & ".mid")
    Else
        Call StopMidi
    End If
End Sub

Private Sub HandleChatMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    AddText Buffer.ReadString, Buffer.ReadByte
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    MapItem(i).Num = Buffer.ReadLong
    MapItem(i).Value = Buffer.ReadLong
    MapItem(i).X = Buffer.ReadLong
    MapItem(i).Y = Buffer.ReadLong
End Sub

Private Sub HandleItemEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ItemEditor
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ItemNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ItemNum = Buffer.ReadLong
    
    If ItemNum <= 0 Then Exit Sub
    If ItemNum > MAX_ITEMS Then Exit Sub
    
    ' Set the item from the byte array
    Set_ItemData ItemNum, Buffer.ReadBytes(ItemSize)
        
    CacheItem ItemNum       ' Caches the item for hover description
End Sub

Private Sub HandleUpdateItems(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ItemCount As Long
Dim ItemNum As Long
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    Buffer.DecompressBuffer
    
    ItemCount = Buffer.ReadLong
    
    For i = 1 To ItemCount
        ItemNum = Buffer.ReadLong
        
        If ItemNum <= 0 Then Exit Sub
        If ItemNum > MAX_ITEMS Then Exit Sub
        
        ' Set the item from the byte array
        Set_ItemData ItemNum, Buffer.ReadBytes(ItemSize)
        
        CacheItem ItemNum
    Next
End Sub

Private Sub HandleEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ItemNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ItemNum = Buffer.ReadLong
    
    If ItemNum <= 0 Then Exit Sub
    If ItemNum > MAX_ITEMS Then Exit Sub
    
    ' Set the item from the byte array
    Set_ItemData ItemNum, Buffer.ReadBytes(ItemSize)
    
    CacheItem ItemNum       ' Caches the item for hover description
    
    ItemEditorInit
End Sub

Private Sub HandleUpdateEmoticon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim EmoticonNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    EmoticonNum = Buffer.ReadLong
    
    If EmoticonNum <= 0 Then Exit Sub
    If EmoticonNum > MAX_EMOTICONS Then Exit Sub
    
    ' Set the Emoticon from the byte array
    Set_EmoticonData EmoticonNum, Buffer.ReadBytes(EmoticonSize)
End Sub

Private Sub HandleUpdateEmoticons(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim EmoticonCount As Long
Dim EmoticonNum As Long
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Buffer.DecompressBuffer
    
    EmoticonCount = Buffer.ReadLong
    
    For i = 1 To EmoticonCount
        EmoticonNum = Buffer.ReadLong
        
        If EmoticonNum <= 0 Then Exit Sub
        If EmoticonNum > MAX_EMOTICONS Then Exit Sub
        
        ' Set the Emoticon from the byte array
        Set_EmoticonData EmoticonNum, Buffer.ReadBytes(EmoticonSize)
    Next
End Sub

Private Sub HandleEditEmoticon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim EmoticonNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    EmoticonNum = Buffer.ReadLong
    
    If EmoticonNum <= 0 Then Exit Sub
    If EmoticonNum > MAX_EMOTICONS Then Exit Sub
    
    ' Set the Emoticon from the byte array
    Set_EmoticonData EmoticonNum, Buffer.ReadBytes(EmoticonSize)
    
    EmoticonEditorInit
End Sub

Private Sub HandleEmoticonEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    InEmoticonEditor = True
    
    frmIndex.Show
    frmIndex.lstIndex.Clear
    
    For i = 1 To MAX_EMOTICONS
        frmIndex.lstIndex.AddItem i & ": " & Trim$(Emoticons(i).Command)
    Next
    
    frmIndex.lstIndex.ListIndex = 0
End Sub

Private Sub HandleCheckEmoticon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
        
    Player(i).EmoticonNum = Emoticons(Buffer.ReadLong).Pic
    Player(i).EmoticonTime = GetTickCount + 2000
End Sub

Private Sub HandleNewTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MyTarget = Buffer.ReadByte
    MyTargetType = Buffer.ReadByte
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    
    If i < 0 Then Exit Sub
    If i > MapNpcCount Then Exit Sub
    
    MapNpc(i).Num = Buffer.ReadLong
    MapNpc(i).X = Buffer.ReadLong
    MapNpc(i).Y = Buffer.ReadLong
    MapNpc(i).Dir = Buffer.ReadLong
    
    ' Client use only
    MapNpc(i).XOffset = 0
    MapNpc(i).YOffset = 0
    MapNpc(i).Moving = 0
End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    
    If i < 0 Then Exit Sub
    If i > MapNpcCount Then Exit Sub
    
    MapNpc(i).Num = 0
    MapNpc(i).X = 0
    MapNpc(i).Y = 0
    MapNpc(i).Dir = 0
    
    ' Client use only
    MapNpc(i).XOffset = 0
    MapNpc(i).YOffset = 0
    MapNpc(i).Moving = 0
End Sub

Private Sub HandleNpcEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    NpcEditor
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    n = Buffer.ReadLong
    
    ' Update the npc
    Npc(n).Name = Buffer.ReadString
    Npc(n).AttackSay = vbNullString
    Npc(n).Sprite = Buffer.ReadInteger
    Npc(n).Behavior = Buffer.ReadByte
    Npc(n).MovementSpeed = Buffer.ReadByte
    Npc(n).MovementFrequency = Buffer.ReadByte
End Sub

Private Sub HandleUpdateNpcs(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim NpcCount As Long
Dim NpcNum As Long
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Buffer.DecompressBuffer
    
    NpcCount = Buffer.ReadLong
    
    For i = 1 To NpcCount
        ' Update the npc
        NpcNum = Buffer.ReadLong
        Npc(NpcNum).Name = Buffer.ReadString
        Npc(NpcNum).AttackSay = vbNullString
        Npc(NpcNum).Sprite = Buffer.ReadInteger
        Npc(NpcNum).Behavior = Buffer.ReadByte
        Npc(NpcNum).MovementSpeed = Buffer.ReadByte
        Npc(NpcNum).MovementFrequency = Buffer.ReadByte
    Next
End Sub

Private Sub HandleEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim NpcNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    NpcNum = Buffer.ReadLong
    
    If NpcNum <= 0 Then Exit Sub
    If NpcNum > MAX_NPCS Then Exit Sub
    
    ' Set the npc from the byte array
    Set_NpcData NpcNum, Buffer.ReadBytes(NpcSize)
    
    ' Initialize the npc editor
    NpcEditorInit
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, Key As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Key = Buffer.ReadByte
    
    If IsValidMapPoint(X, Y) Then
        TempTile(X, Y).Open = Key
    End If
End Sub

Private Sub HandleEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    MapEditorInit
End Sub

Private Sub HandleShopEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ShopEditor
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ShopNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ShopNum = Buffer.ReadLong
    
    If ShopNum <= 0 Then Exit Sub
    If ShopNum > MAX_SHOPS Then Exit Sub
    
    ' Set the Shop from the byte array
    Set_ShopData ShopNum, Buffer.ReadBytes(ShopSize)
End Sub

Private Sub HandleUpdateShops(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ShopCount As Long
Dim ShopNum As Long
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    Buffer.DecompressBuffer
    
    ShopCount = Buffer.ReadLong
    
    For i = 1 To ShopCount
        ShopNum = Buffer.ReadLong
        
        If ShopNum <= 0 Then Exit Sub
        If ShopNum > MAX_SHOPS Then Exit Sub
        
        ' Set the Shop from the byte array
        Set_ShopData ShopNum, Buffer.ReadBytes(ShopSize)
    Next
End Sub

Private Sub HandleEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ShopNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ShopNum = Buffer.ReadLong
    
    If ShopNum <= 0 Then Exit Sub
    If ShopNum > MAX_SHOPS Then Exit Sub
    
   ' Set the Shop from the byte array
    Set_ShopData ShopNum, Buffer.ReadBytes(ShopSize)

    ' Initialize the shop editor
    ShopEditorInit
End Sub

Private Sub HandleSpellEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SpellEditor
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim SpellNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SpellNum = Buffer.ReadLong
    
    If SpellNum <= 0 Then Exit Sub
    If SpellNum > MAX_SPELLS Then Exit Sub
    
    ' Set the Spell from the byte array
    Set_SpellData SpellNum, Buffer.ReadBytes(SpellSize)
End Sub

Private Sub HandleUpdateSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim SpellCount As Long
Dim SpellNum As Long
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    Buffer.DecompressBuffer
    
    SpellCount = Buffer.ReadLong
    
    For i = 1 To SpellCount
        SpellNum = Buffer.ReadLong
        
        If SpellNum <= 0 Then Exit Sub
        If SpellNum > MAX_SPELLS Then Exit Sub
        
        ' Set the Spell from the byte array
        Set_SpellData SpellNum, Buffer.ReadBytes(SpellSize)
    Next
End Sub

Private Sub HandleEditSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim SpellNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SpellNum = Buffer.ReadLong
    
    If SpellNum <= 0 Then Exit Sub
    If SpellNum > MAX_SPELLS Then Exit Sub
    
    ' Set the Spell from the byte array
    Set_SpellData SpellNum, Buffer.ReadBytes(SpellSize)
    
    Call SpellEditorInit
End Sub

'///////////////
'// Animation //
'///////////////
Private Sub HandleAnimationEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    InAnimationEditor = True
    
    frmIndex.Show
    frmIndex.lstIndex.Clear
    
    ' Add the names
    For i = 1 To MAX_ANIMATIONS
        frmIndex.lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
    Next
    
    frmIndex.lstIndex.ListIndex = 0
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim AnimationNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    AnimationNum = Buffer.ReadLong
    
    If AnimationNum <= 0 Then Exit Sub
    If AnimationNum > MAX_ANIMATIONS Then Exit Sub
    
    ' Set the animation from the byte array
    Set_AnimationData AnimationNum, Buffer.ReadBytes(AnimationSize)
End Sub

Private Sub HandleUpdateAnimations(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim AnimationCount As Long
Dim AnimationNum As Long
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Buffer.DecompressBuffer
    
    AnimationCount = Buffer.ReadLong
    
    For i = 1 To AnimationCount
        AnimationNum = Buffer.ReadLong
        
        If AnimationNum <= 0 Then Exit Sub
        If AnimationNum > MAX_ANIMATIONS Then Exit Sub
        
        ' Set the animation from the byte array
        Set_AnimationData AnimationNum, Buffer.ReadBytes(AnimationSize)
    Next
End Sub

Private Sub HandleEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim AnimationNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    AnimationNum = Buffer.ReadLong
    
    If AnimationNum <= 0 Then Exit Sub
    If AnimationNum > MAX_ANIMATIONS Then Exit Sub
    
   ' Set the animation from the byte array
    Set_AnimationData AnimationNum, Buffer.ReadBytes(AnimationSize)
    
    ' Initialize the shop editor
    AnimationEditorInit
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    ShopNpcNum = Buffer.ReadByte
    InShop = Buffer.ReadLong
    
    frmTrade.lblShopName.Caption = Buffer.ReadString
    frmTrade.lblShopDescription.Caption = Buffer.ReadString
    
    If Shop(InShop).Type = SHOP_TYPE_SHOP Then
        frmTrade.Width = 5550
        BltTrade InShop
    ElseIf Shop(InShop).Type = SHOP_TYPE_INN Then
        frmTrade.Width = 3735
        frmTrade.picSetHome.Visible = True
        frmTrade.lblSetHome.Visible = True
    End If
    
    If Not frmTrade.Visible Then
        frmTrade.Show vbModal
    End If
End Sub

Private Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    frmMainGame.lstSpells.Clear
    
    ' Put spells known in player record
    For i = 1 To MAX_PLAYER_SPELLS
        Player(MyIndex).Spell(i).SpellNum = Buffer.ReadLong
        Player(MyIndex).Spell(i).Cooldown = Buffer.ReadLong
        If Player(MyIndex).Spell(i).SpellNum <> 0 Then
            frmMainGame.lstSpells.AddItem "[End]: " & Trim$(Spell(Player(MyIndex).Spell(i).SpellNum).Name)
        Else
            frmMainGame.lstSpells.AddItem "[*]: " '& i & ": " '< --- empty ability space --- >"
        End If
    Next

    frmMainGame.lstSpells.ListIndex = 0
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1
    
    With ActionMsg(ActionMsgIndex)
        .Message = Buffer.ReadString
        .Color = Buffer.ReadLong
        .Type = Buffer.ReadLong
        .Scroll = 1
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        
        Select Case .Type
            Case ACTIONMSG_STATIC
                .Created = GetTickCount + 1500
                
            Case ACTIONMSG_SCROLL
                .Created = GetTickCount + 1500
                
            Case ACTIONMSG_SCREEN
                .Created = GetTickCount + 3000
                ' This will kill any action screen Messages that there in the system
                For i = 1 To MAX_BYTE
                    If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                        If i <> ActionMsgIndex Then
                            ClearActionMsg i
                        End If
                    End If
                Next
        End Select
    End With
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    AnimIndex = AnimIndex + 1
    If AnimIndex > MAX_BYTE - 1 Then AnimIndex = 1
    
    With Anim(AnimIndex)
        .Anim = 1
        .AnimNum = Buffer.ReadByte
        .Created = GetTickCount
        .CurrFrame = 0
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .MaxFrames = Animation(.AnimNum).AnimationFrames
        .Speed = Animation(.AnimNum).AnimationSpeed
        .Size = Animation(.AnimNum).AnimationSize
        .Layer = Animation(.AnimNum).AnimationLayer
    End With
End Sub

Private Sub HandlePlayerGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    
    Player(i).GuildName = Buffer.ReadString
    Player(i).GuildAbbreviation = Buffer.ReadString
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Player(MyIndex).Exp = Buffer.ReadLong
    Player(MyIndex).NextLevel = Buffer.ReadLong
    
    frmMainGame.PicExp.Width = Abs((Current_Exp(MyIndex) / Current_NextLevel(MyIndex)) * 178)
End Sub

Private Sub HandleCancelSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    CastingSpell = 0
    CastTime = 0
End Sub

Private Sub HandleSpellReady(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim SpellSlot As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SpellSlot = Buffer.ReadLong
    Player(MyIndex).Spell(SpellSlot).Cooldown = 0
    AddText Trim$(Spell(Player(MyIndex).Spell(SpellSlot).SpellNum).Name) & " is ready to be cast again!", White
End Sub

Private Sub HandleSpellCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim SpellSlot As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SpellSlot = Buffer.ReadLong
    Player(MyIndex).Spell(SpellSlot).Cooldown = GetTickCount + (Spell(Player(MyIndex).Spell(SpellSlot).SpellNum).Cooldown * 1000)
End Sub

Private Sub HandleLeftGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    
    ClearPlayer i
    UpdateMapPlayers
End Sub

Private Sub HandlePlayerDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim Timer As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    Timer = Buffer.ReadLong + GetTickCount
    
    Update_IsDead i, True
    Update_IsDeadTimer i, Timer
End Sub

Private Sub HandlePlayerGold(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    
    'TODO
End Sub

Private Sub HandlePlayerRevival(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    AddText Name & " has offered to revive you. Type /revive to accept.", White
    AlertMessage Name & " has offered to revive you. Do you accept?", AddressOf Revive_Click, False
End Sub

Private Sub HandleQuestEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    QuestEditor
End Sub

Private Sub HandleUpdateQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim QuestNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    QuestNum = Buffer.ReadLong
    
    If QuestNum <= 0 Then Exit Sub
    If QuestNum > MAX_QUESTS Then Exit Sub
    
    ' Set the Quest from the byte array
    Set_QuestData QuestNum, Buffer.ReadBytes(QuestSize)
    
    ' Update the Npcs Quest List
    Update_Npcs_Quest Quest(QuestNum).StartNPC
    Update_Npcs_Quest Quest(QuestNum).EndNPC
End Sub

Private Sub HandleUpdateQuests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim QuestCount As Long
Dim QuestNum As Long
Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    Buffer.DecompressBuffer
    
    QuestCount = Buffer.ReadLong
    
    For i = 1 To QuestCount
        QuestNum = Buffer.ReadLong
        
        If QuestNum <= 0 Then Exit Sub
        If QuestNum > MAX_QUESTS Then Exit Sub
        
        ' Set the Quest from the byte array
        Set_QuestData QuestNum, Buffer.ReadBytes(QuestSize)
    Next
    
    ' Update the Npc quest list
    Update_Npc_Quests
End Sub

Private Sub HandleEditQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim QuestNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    QuestNum = Buffer.ReadLong
    
    If QuestNum <= 0 Then Exit Sub
    If QuestNum > MAX_QUESTS Then Exit Sub
    
    ' Set the Quest from the byte array
    Set_QuestData QuestNum, Buffer.ReadBytes(QuestSize)
    
    QuestEditorInit
End Sub

Private Sub HandleAvailableQuests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim QuestNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    QuestMapNpcNum = Buffer.ReadLong
    
    frmMainGame.lstNpcQuests.Clear
    Do While Buffer.Length > 0
        QuestNum = Buffer.ReadLong
        frmMainGame.lstNpcQuests.AddItem Trim$(Quest(QuestNum).Name)
        frmMainGame.lstNpcQuests.ItemData(i) = QuestNum
        i = i + 1
    Loop
    If i Then
        frmMainGame.lstNpcQuests.ListIndex = 0
        frmMainGame.picNpcQuests.Visible = True
        frmMainGame.picNpcQuestInfo.Visible = True
        frmMainGame.picTurnInQuest.Visible = False
    End If
    
'    Set Buffer = New clsBuffer
'    Buffer.WriteBytes Data()
'
'    QuestMapNpcNum = Buffer.ReadLong
'
'    frmQuests.lblQuestType.Caption = "NPC Quests"
'    frmQuests.lstNpcQuests.Clear
'    Do While Buffer.Length > 0
'        QuestNum = Buffer.ReadLong
'        frmQuests.lstNpcQuests.AddItem Trim$(Quest(QuestNum).Name)
'        frmQuests.lstNpcQuests.ItemData(i) = QuestNum
'        i = i + 1
'    Loop
'    If i Then
'        frmQuests.lstNpcQuests.ListIndex = 0
'        frmQuests.Visible = True
'        frmQuests.lstNpcQuests.Visible = True
'        frmQuests.frmNpcQuest.Visible = False
'    End If
End Sub

Private Sub HandlePlayerQuests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim n As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_QUESTS
        Player(MyIndex).QuestProgress(i).QuestNum = Buffer.ReadLong
        For n = 1 To MAX_QUEST_NEEDS
            Player(MyIndex).QuestProgress(i).Progress(n) = Buffer.ReadLong
        Next
    Next
    
    ' Update the GUI
    UpdateQuestList
End Sub

Private Sub HandlePlayerQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim n As Long
Dim QuestProgressNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    QuestProgressNum = Buffer.ReadLong
    Player(MyIndex).QuestProgress(QuestProgressNum).QuestNum = Buffer.ReadLong
    For n = 1 To MAX_QUEST_NEEDS
        Player(MyIndex).QuestProgress(QuestProgressNum).Progress(n) = Buffer.ReadLong
    Next
    
    ' Update the GUI
    UpdateQuestList
End Sub

