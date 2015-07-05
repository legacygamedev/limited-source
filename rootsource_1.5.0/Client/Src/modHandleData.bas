Attribute VB_Name = "DataHandling"
Option Explicit

' ******************************************
' **               rootSource               **
' ** Parses and handles String packets    **
' ******************************************

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SAllChars) = GetAddress(AddressOf HandleAllChars)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerSp) = GetAddress(AddressOf HandlePlayerSp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SEditItem) = GetAddress(AddressOf HandleEditItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SEditNpc) = GetAddress(AddressOf HandleEditNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SEditShop) = GetAddress(AddressOf HandleEditShop)
    HandleDataSub(SREditor) = GetAddress(AddressOf HandleRefresh)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SEditSpell) = GetAddress(AddressOf HandleEditSpell)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SCastSpell) = GetAddress(AddressOf HandleSpellCast)
    HandleDataSub(SDoor) = GetAddress(AddressOf HandleDoor)
    HandleDataSub(SSendMaxes) = GetAddress(AddressOf HandleMaxes)
    HandleDataSub(SSync) = GetAddress(AddressOf HandleSync)
    HandleDataSub(SMapRevs) = GetAddress(AddressOf HandleMapRevs)
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Integer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadInteger
    
    If MsgType < 0 Or MsgType >= SMSG_COUNT Then
        MsgBox "Packet Error.", vbOKOnly
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

' ::::::::::::::::::::::::::
' :: Alert message packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim msg As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    
    frmSendGetData.Visible = False
    frmMainMenu.Visible = True
    
    Call MsgBox(msg, vbOKOnly, GAME_NAME)
End Sub

' :::::::::::::::::::::::::::
' :: All characters packet ::
' :::::::::::::::::::::::::::
Private Sub HandleAllChars(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Level As Long
Dim Name As String
Dim msg As String
Dim Buffer As clsBuffer

    ReDim CharSprites(1 To MAX_CHARS) As Long

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
        
    With frmMainMenu
       .mnuChars.Visible = True
       
       frmSendGetData.Visible = False
       
       .lstChars.Clear
       
       For i = 1 To MAX_CHARS
           CharSprites(i) = Buffer.ReadLong
           Name = Buffer.ReadString
           msg = Buffer.ReadString
           Level = Buffer.ReadByte
           
           If Trim$(Name) = vbNullString Then
               .lstChars.AddItem "Free Character Slot"
           Else
               .lstChars.AddItem Name & " a level " & Level & " " & msg
           End If
       Next
       
       DrawSelChar (1)
       
       .lstChars.ListIndex = 0
    End With
End Sub

' :::::::::::::::::::::::::::::::::
' :: Login was successful packet ::
' :::::::::::::::::::::::::::::::::
Private Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    
    frmSendGetData.Visible = True
    frmMainMenu.Visible = False
    frmMainMenu.mnuChars.Visible = False
    
    Call SetStatus("Receiving game data...")
End Sub

' :::::::::::::::::::::::::::::::::::::::
' :: New character classes data packet ::
' :::::::::::::::::::::::::::::::::::::::
Private Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
     
    ' Max classes
    Max_Classes = Buffer.ReadByte
    ReDim Class(1 To Max_Classes)
        
    For i = 1 To Max_Classes
        With Class(i)
            .Name = Buffer.ReadString
            .Sprite = Buffer.ReadLong
            For n = 1 To Vitals.Vital_Count - 1
                .Vital(n) = Buffer.ReadLong
            Next
            For n = 1 To Stats.Stat_Count - 1
                .Stat(n) = Buffer.ReadByte
            Next
        End With
    Next
    
    ' Used for if the player is creating a new character
    With frmMainMenu
       .mnuNewCharacter.Visible = True
       .PreviewTimer = True
       
       frmSendGetData.Visible = False
       
       .cmbClass.Clear
       
       For i = 1 To Max_Classes
           .cmbClass.AddItem Trim$(Class(i).Name)
       Next
       
       .cmbClass.ListIndex = 0
       
       n = .cmbClass.ListIndex + 1
       
       .lblHP.Caption = CStr(Class(n).Vital(Vitals.HP))
       .lblMP.Caption = CStr(Class(n).Vital(Vitals.MP))
       .lblSP.Caption = CStr(Class(n).Vital(Vitals.SP))
    
       '.lblStrength.Caption = CStr(Class(n).Stat(Stats.Strength))
       '.lblDefense.Caption = CStr(Class(n).Stat(Stats.Defense))
       '.lblSpeed.Caption = CStr(Class(n).Stat(Stats.Speed))
       '.lblMagic.Caption = CStr(Class(n).Stat(Stats.Magic))
    End With
End Sub

' :::::::::::::::::::::::::
' :: Classes data packet ::
' :::::::::::::::::::::::::
Private Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Max classes
    Max_Classes = Buffer.ReadByte
    ReDim Class(1 To Max_Classes)
    
    For i = 1 To Max_Classes
        With Class(i)
            .Name = Buffer.ReadString
            .Sprite = Buffer.ReadLong
            For n = 1 To Vitals.Vital_Count - 1
                .Vital(n) = Buffer.ReadLong
            Next
            For n = 1 To Stats.Stat_Count - 1
                .Stat(n) = Buffer.ReadByte
            Next
        End With
    Next
End Sub

' ::::::::::::::::::::
' :: In game packet ::
' ::::::::::::::::::::
Private Sub HandleInGame()
     InGame = True
     Call GameInit
     Call GameLoop
End Sub

' :::::::::::::::::::::::::::::
' :: Player inventory packet ::
' :::::::::::::::::::::::::::::
Private Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemDur(MyIndex, i, Buffer.ReadLong)
    Next
    Call UpdateInventory
 End Sub

' ::::::::::::::::::::::::::::::::::::
' :: Player inventory update packet ::
' ::::::::::::::::::::::::::::::::::::
Private Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong)
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong)
    Call SetPlayerInvItemDur(MyIndex, n, Buffer.ReadLong)
    Call UpdateInventory
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player worn equipment packet ::
' ::::::::::::::::::::::::::::::::::
Private Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    For i = 1 To Equipment.Equipment_Count - 1
        Call SetPlayerEquipmentSlot(MyIndex, Buffer.ReadByte, i)
    Next
    Call UpdateInventory
End Sub

' ::::::::::::::::::::::
' :: Player hp packet ::
' ::::::::::::::::::::::
Private Sub HandlePlayerHp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Player(MyIndex).MaxHP = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)
    
    If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
        frmMainGame.lblHP.Caption = Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%"
    End If
End Sub

' ::::::::::::::::::::::
' :: Player mp packet ::
' ::::::::::::::::::::::
Private Sub HandlePlayerMp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Player(MyIndex).MaxMP = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)
    
    If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
        frmMainGame.lblMP.Caption = Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%"
    End If
End Sub

' ::::::::::::::::::::::
' :: Player sp packet ::
' ::::::::::::::::::::::
Private Sub HandlePlayerSp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Player(MyIndex).MaxSP = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.SP, Buffer.ReadLong)
    
    If GetPlayerMaxVital(MyIndex, Vitals.SP) > 0 Then
        frmMainGame.lblSP.Caption = Int(GetPlayerVital(MyIndex, Vitals.SP) / GetPlayerMaxVital(MyIndex, Vitals.SP) * 100) & "%"
    End If
End Sub

' :::::::::::::::::::::::::
' :: Player stats packet ::
' :::::::::::::::::::::::::
Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    For i = 1 To Stats.Stat_Count - 1
        Call SetPlayerStat(MyIndex, i, Buffer.ReadLong)
    Next
End Sub

' ::::::::::::::::::::::::
' :: Player data packet ::
' ::::::::::::::::::::::::
Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tempmap As Long
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
      
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
        
    tempmap = GetPlayerMap(i)
    
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Player(i).Guild = Buffer.ReadString
    Player(i).GuildAccess = Buffer.ReadLong
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    
    If i <> MyIndex Then
        If GetPlayerMap(i) <> tempmap Then
            Select Case GetPlayerDir(i)
                Case DIR_UP
                    Player(i).YOffset = PIC_Y
                Case DIR_DOWN
                    Player(i).YOffset = -1 * PIC_Y
                Case DIR_LEFT
                    Player(i).XOffset = PIC_X
                Case DIR_RIGHT
                    Player(i).XOffset = -1 * PIC_X
            End Select
                Player(i).Moving = 1
        End If
    End If
    
    ' Check if the player is the client player, and if so reset directions
    If i = MyIndex Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
    End If
    
    ' Make sure they aren't walking
    'Player(i).Moving = 0
    'Player(i).XOffset = 0
    'Player(i).YOffset = 0
    
    Call GetPlayersOnMap
End Sub

' ::::::::::::::::::::::::::::
' :: Player movement packet ::
' ::::::::::::::::::::::::::::
Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim x As Long
Dim y As Long
Dim Dir As Long
Dim n As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    n = Buffer.ReadLong
    
    Call SetPlayerX(i, x)
    Call SetPlayerY(i, y)
    Call SetPlayerDir(i, Dir)
            
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    Player(i).Moving = n
    
    Select Case GetPlayerDir(i)
        Case DIR_UP
            Player(i).YOffset = PIC_Y
        Case DIR_DOWN
            Player(i).YOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(i).XOffset = PIC_X
        Case DIR_RIGHT
            Player(i).XOffset = PIC_X * -1
    End Select
End Sub

' :::::::::::::::::::::::::
' :: Npc movement packet ::
' :::::::::::::::::::::::::
Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim x As Long
Dim y As Long
Dim Dir As Long
Dim Movement As Long
Dim map As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    map = Buffer.ReadLong
    MapNpcNum = Buffer.ReadInteger
    x = Buffer.ReadByte
    y = Buffer.ReadByte
    Dir = Buffer.ReadInteger
    Movement = Buffer.ReadLong

    With MapNpc(MapNpcNum, map)
        .x = x
        .y = y
        .Dir = Dir
        .XOffset = 0
        .YOffset = 0
        .Moving = Movement
        
        Select Case .Dir
            Case DIR_UP
                .YOffset = PIC_Y
            Case DIR_DOWN
                .YOffset = PIC_Y * -1
            Case DIR_LEFT
                .XOffset = PIC_X
            Case DIR_RIGHT
                .XOffset = PIC_X * -1
        End Select
    End With
End Sub

' :::::::::::::::::::::::::::::
' :: Player direction packet ::
' :::::::::::::::::::::::::::::
Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    Dir = Buffer.ReadLong
    
    Call SetPlayerDir(i, Dir)
    
    With Player(i)
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With
End Sub

' ::::::::::::::::::::::::::
' :: NPC direction packet ::
' ::::::::::::::::::::::::::
Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Dir As Byte
Dim Buffer As clsBuffer
Dim map As Long

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    map = Buffer.ReadLong
    i = Buffer.ReadInteger
    Dir = Buffer.ReadInteger
    
    With MapNpc(i, map)
       .Dir = Dir
    
       .XOffset = 0
       .YOffset = 0
       .Moving = 0
    End With
End Sub

' :::::::::::::::::::::::::::::::
' :: Player XY location packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()

     x = Buffer.ReadLong
     y = Buffer.ReadLong
     
     Call SetPlayerX(MyIndex, x)
     Call SetPlayerY(MyIndex, y)
     
     ' Make sure they aren't walking
     Player(MyIndex).Moving = 0
     Player(MyIndex).XOffset = 0
     Player(MyIndex).YOffset = 0
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
     
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount
End Sub

' :::::::::::::::::::::::
' :: NPC attack packet ::
' :::::::::::::::::::::::
Private Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
   
    i = Buffer.ReadLong

    ' Set player to attacking
    MapNpc(i, GetPlayerMap(MyIndex)).Attacking = 1
    MapNpc(i, GetPlayerMap(MyIndex)).AttackTimer = GetTickCount
End Sub

' ::::::::::::::::::::::::::
' :: Check for map packet ::
' ::::::::::::::::::::::::::
Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
Dim i As Long
Dim NeedMap As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
     ' Erase all players except self
     For i = 1 To High_Index
         If i <> MyIndex Then
             Call SetPlayerMap(i, 0)
         End If
     Next

     ' Erase all temporary tile values
    Call ClearTempTile
    
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap

     ' Get map num
     x = Buffer.ReadLong
     
     ' Get revision
     y = Buffer.ReadLong
     
    NeedMap = 1
     
     If FileExist(MAP_PATH & "map" & x & MAP_EXT, False) Then
         Call LoadMaps(x)
     
         ' Check to see if the revisions match
         NeedMap = 1
         If map(5).Revision = y Then
             ' We do so we dont need the map
             'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
             NeedMap = 0
         End If
     End If
     
     ' Either the revisions didn't match or we dont have the map, so we need it
     Set Buffer = New clsBuffer
     Buffer.PreAllocate 3
     Buffer.WriteInteger CNeedMap
     Buffer.WriteByte NeedMap
     Call SendData(Buffer.ToArray())
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Private Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim MapNum As Long
Dim Buffer As clsBuffer
Dim MapSize As Long
Dim MapData() As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
     
    MapNum = Buffer.ReadLong
    
    If MapNum < 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
             
    MapSize = LenB(map(5))
    ReDim MapData(MapSize - 1)
    MapData = Buffer.ReadBytes(MapSize)
    CopyMemory ByVal VarPtr(map(5)), ByVal VarPtr(MapData(0)), MapSize
    
    ' Save the map
    Call SaveMap(MapNum)
    
    If InGame Then
        If Player(MyIndex).map = MapNum Then
            Call LoadMaps(MapNum)
            CalcTilePositions
        End If
    End If
    
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If Editor = EDITOR_MAP Then
        Editor = EDITOR_NONE
        frmMainGame.picMapEditor.Visible = False
        
        If frmMapWarp.Visible Then
            Unload frmMapWarp
        End If
        
        If frmMapProperties.Visible Then
            Unload frmMapProperties
        End If
    End If
End Sub

' :::::::::::::::::::::::::::
' :: Map items data packet ::
' :::::::::::::::::::::::::::
Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer
Dim MapNum As Long
    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    MapNum = Buffer.ReadLong
    For i = 1 To MAX_MAP_ITEMS
       With MapItem(i, MapNum)
           .Num = Buffer.ReadByte
           .Value = Buffer.ReadLong
           .Dur = Buffer.ReadInteger
           .x = Buffer.ReadByte
           .y = Buffer.ReadByte
       End With
    Next
End Sub

' :::::::::::::::::::::::::
' :: Map npc data packet ::
' :::::::::::::::::::::::::
Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Dim map As Long
    map = Buffer.ReadLong
    
    For i = 1 To MAX_MAP_NPCS
       With MapNpc(i, map)
           .Num = Buffer.ReadInteger
           .x = Buffer.ReadByte
           .y = Buffer.ReadByte
           .Dir = Buffer.ReadInteger
       End With
    Next
End Sub

' :::::::::::::::::::::::::::::::
' :: Map send completed packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandleMapDone()
Dim i As Long
Dim MusicFile As String

    MusicFile = Trim$(CStr(map(5).Music)) & ".mid"

    ' Get high NPC index
    High_Npc_Index = 0
    For i = 1 To MAX_MAP_NPCS
        If map(5).Npc(i) > 0 Then
            High_Npc_Index = High_Npc_Index + 1
        Else
            Exit For
        End If
    Next
    
    ' Play music
    If map(5).Music > 0 Then
        If MusicFile <> CurrentMusic Then
            StopMusic
            PlayMusic (MusicFile)
            CurrentMusic = MusicFile
        End If
    Else
        StopMusic
        CurrentMusic = 0
    End If
    
    Call UpdateDrawMapName
    
    If map(5).Shop = 0 Then
        frmMainGame.picTradeButton.Visible = False
    Else
        frmMainGame.picTradeButton.Visible = True
    End If
    
    Call CalcTilePositions

    GettingMap = False
    CanMoveNow = True
End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    Color = Buffer.ReadByte
    
    Call AddText(msg, Color)
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    Color = Buffer.ReadByte
    
    Call AddText(msg, Color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    Color = Buffer.ReadByte
    
    Call AddText(msg, Color)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    Color = Buffer.ReadByte
    
    Call AddText(msg, Color)
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    Color = Buffer.ReadByte
    
    Call AddText(msg, Color)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    msg = Buffer.ReadString
    Color = Buffer.ReadByte
    
    Call AddText(msg, Color)
End Sub

' :::::::::::::::::::::::::::
' :: Refresh editor packet ::
' :::::::::::::::::::::::::::
Private Sub HandleRefresh()
    Dim i As Long
    
    frmIndex.lstIndex.Clear
    
    Select Case Editor
        Case EDITOR_ITEM
            For i = 1 To MAX_ITEMS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
            Next
        Case EDITOR_NPC
            For i = 1 To MAX_NPCS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
            Next
        Case EDITOR_SHOP
            For i = 1 To MAX_SHOPS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
            Next
        Case EDITOR_SPELL
            For i = 1 To MAX_SPELLS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
            Next
    End Select
     
    frmIndex.lstIndex.ListIndex = 0
     
End Sub

' :::::::::::::::::::::::
' :: Item spawn packet ::
' :::::::::::::::::::::::
Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer
Dim MapNum As Long

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    MapNum = Buffer.ReadLong
    n = Buffer.ReadLong
    
    With MapItem(n, MapNum)
       .Num = Buffer.ReadLong
       .Value = Buffer.ReadLong
       .Dur = Buffer.ReadLong
       .x = Buffer.ReadLong
       .y = Buffer.ReadLong
    End With
End Sub

' ::::::::::::::::::::::::
' :: Item editor packet ::
' ::::::::::::::::::::::::
Private Sub HandleItemEditor()
Dim i As Long

    With frmIndex
        .Caption = "Item Index"
        Editor = EDITOR_ITEM
        
        .lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next
        
        .lstIndex.ListIndex = 0
        .Show vbModal
    End With
End Sub

' ::::::::::::::::::::::::
' :: Update item packet ::
' ::::::::::::::::::::::::
Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    ' Update the item
    With Item(n)
       .Name = Buffer.ReadString
       .Pic = Buffer.ReadInteger
       .Type = Buffer.ReadByte
       .Data1 = Buffer.ReadInteger
       .Data2 = Buffer.ReadInteger
       .Data3 = Buffer.ReadInteger
    End With
End Sub

' ::::::::::::::::::::::
' :: Edit item packet ::
' ::::::::::::::::::::::
Private Sub HandleEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ItemNum As Long
Dim ItemSize As Long
Dim ItemData() As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    ItemNum = Buffer.ReadLong
     
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    ' Update the item
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(ItemNum)), ByVal VarPtr(ItemData(0)), ItemSize
     
    ' Initialize the item editor
    Call ItemEditorInit
End Sub

' ::::::::::::::::::::::
' :: Npc spawn packet ::
' ::::::::::::::::::::::
Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n, map As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    map = Buffer.ReadLong
    n = Buffer.ReadLong
    
    With MapNpc(n, map)
       .Num = Buffer.ReadInteger
       .x = Buffer.ReadByte
       .y = Buffer.ReadByte
       .Dir = Buffer.ReadInteger
       
       ' Client use only
       .XOffset = 0
       .YOffset = 0
       .Moving = 0
    End With
End Sub

' :::::::::::::::::::::
' :: Npc dead packet ::
' :::::::::::::::::::::
 Private Sub HandleNpcDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
 Dim n As Long
 Dim Buffer As clsBuffer
 Dim tMap As Long
 
    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    tMap = Buffer.ReadLong
    n = Buffer.ReadLong
    Call ClearMapNpc(n, tMap)
    
End Sub

' :::::::::::::::::::::::
' :: Npc editor packet ::
' :::::::::::::::::::::::
Private Sub HandleNpcEditor()
Dim i As Long

    With frmIndex
        .Caption = "NPC Index"
        Editor = EDITOR_NPC
         
        .lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
        Next
        
        .lstIndex.ListIndex = 0
        .Show vbModal
    End With
End Sub

' :::::::::::::::::::::::
' :: Update npc packet ::
' :::::::::::::::::::::::
Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    ' Update the item
    With Npc(n)
       .Name = Buffer.ReadString
       .AttackSay = vbNullString
       .Sprite = Buffer.ReadInteger
       .SpawnSecs = 0
       .Behavior = 0
       .Range = 0
       .DropChance = 0
       .DropItem = 0
       .DropItemValue = 0
       .Stat(Stats.Strength) = 0
       .Stat(Stats.Defense) = 0
       .Stat(Stats.Speed) = 0
       .Stat(Stats.Magic) = 0
    End With
End Sub

' :::::::::::::::::::::
' :: Edit npc packet ::
' :::::::::::::::::::::
Private Sub HandleEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim NpcNum As Long
Dim NpcSize As Long
Dim NpcData() As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    NpcNum = Buffer.ReadLong
     
    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Exit Sub
    End If
    
    ' Update the Npc
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(NpcNum)), ByVal VarPtr(NpcData(0)), NpcSize
    
    ' Initialize the npc editor
    Call NpcEditorInit
End Sub

' ::::::::::::::::::::
' :: Map key packet ::
' ::::::::::::::::::::
Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim x As Long
Dim y As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    n = Buffer.ReadLong
    
    TempTile(x, y).DoorOpen = n
End Sub

' :::::::::::::::::::::
' :: Edit map packet ::
' :::::::::::::::::::::
Private Sub HandleEditMap()
     Call MapEditorInit
End Sub

' ::::::::::::::::::::::::
' :: Shop editor packet ::
' ::::::::::::::::::::::::
Private Sub HandleShopEditor()
Dim i As Long

    With frmIndex
        .Caption = "Shop Index"
         Editor = EDITOR_SHOP
         
         .lstIndex.Clear
         
         ' Add the names
         For i = 1 To MAX_SHOPS
             .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
         Next
         
         .lstIndex.ListIndex = 0
         .Show vbModal
     End With
End Sub

' ::::::::::::::::::::::::
' :: Update shop packet ::
' ::::::::::::::::::::::::
Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
     
    ' Update the shop name
    Shop(n).Name = Buffer.ReadString
End Sub

' ::::::::::::::::::::::
' :: Edit shop packet ::
' ::::::::::::::::::::::
Private Sub HandleEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ShopNum As Long
Dim ShopSize As Long
Dim ShopData() As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    ShopNum = Buffer.ReadLong
     
    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Exit Sub
    End If
    
    ' Update the Shop
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize
     
    ' Initialize the shop editor
    Call ShopEditorInit
     
End Sub

' :::::::::::::::::::::::::
' :: Spell editor packet ::
' :::::::::::::::::::::::::
Private Sub HandleSpellEditor()
Dim i As Long

    With frmIndex
        .Caption = "Spell Index"
        Editor = EDITOR_SPELL
        
        .lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
        Next
        
        .lstIndex.ListIndex = 0
        .Show vbModal
    End With
End Sub

' ::::::::::::::::::::::::
' :: Update spell packet ::
' ::::::::::::::::::::::::
Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    ' Update the spell name
    With Spell(n)
       .Name = Buffer.ReadString
       .MPReq = Buffer.ReadInteger
       .Pic = Buffer.ReadInteger
    End With
End Sub

' :::::::::::::::::::::::
' :: Edit spell packet ::
Private Sub HandleEditSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim SpellSize As Long
Dim SpellData() As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    spellnum = Buffer.ReadLong
     
    If spellnum < 0 Or spellnum > MAX_SPELLS Then
        Exit Sub
    End If
    
    ' Update the Spell
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    
    ' Initialize the spell editor
    Call SpellEditorInit
End Sub

' ::::::::::::::::::
' :: Trade packet ::
' ::::::::::::::::::
 Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
 Dim i As Long
 Dim ShopNum As Long
 Dim GiveItem As Long
 Dim GiveValue As Long
 Dim GetItem As Long
 Dim GetValue As Long
 Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()

    ShopNum = Buffer.ReadLong
    
    With frmTrade
       If Buffer.ReadByte = 1 Then
           .lblFixItem.Visible = True
       Else
           .lblFixItem.Visible = False
       End If
       
       For i = 1 To MAX_TRADES
           GiveItem = Buffer.ReadLong
           GiveValue = Buffer.ReadLong
           GetItem = Buffer.ReadLong
           GetValue = Buffer.ReadLong
           
           If GiveItem > 0 Then
               If GetItem > 0 Then
                   .lstTrade.AddItem "Give " & Trim$(Shop(ShopNum).Name) & " " & GiveValue & " " & Trim$(Item(GiveItem).Name) & " for " & GetValue & " " & Trim$(Item(GetItem).Name)
               End If
           End If
       Next
       
       If .lstTrade.ListCount > 0 Then
           .lstTrade.ListIndex = 0
       End If
       .Show vbModal
    End With
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Private Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim j As Long
Dim n As Long
Dim k As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    With frmMainGame
        .picSpellsList.Visible = True
        .lstSpells.Clear
                
        For i = 1 To MAX_PLAYER_SPELLS
           k = Buffer.ReadLong
           PlayerSpells(k) = Buffer.ReadLong
        Next
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            If PlayerSpells(i) <> 0 Then
                .lstSpells.AddItem i & ": " & Trim$(Spell(PlayerSpells(i)).Name)
            Else
                .lstSpells.AddItem "<free spells slot>"
            End If
        Next
        
        .lstSpells.ListIndex = 0
    End With
End Sub

' ::::::::::::::::::::::
' :: Left game packet ::
' ::::::::::::::::::::::
Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Call ClearPlayer(Buffer.ReadLong)
    Call GetPlayersOnMap
End Sub

' ::::::::::::::::::::::
' :: HighIndex packet ::
' ::::::::::::::::::::::
Private Sub HandleHighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    High_Index = Buffer.ReadLong
End Sub

' :::::::::::::::::::::::
' :: Spell Cast packet ::
' :::::::::::::::::::::::
Private Sub HandleSpellCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim TargetType As Byte
Dim n As Long
Dim spellnum As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    TargetType = Buffer.ReadByte
    n = Buffer.ReadLong
    spellnum = Buffer.ReadLong
    
    If n = 0 Or spellnum = 0 Then
        Exit Sub
    End If
    
    Select Case TargetType
        Case TARGET_TYPE_PLAYER
        
            For i = 1 To MAX_SPELLANIM
                With Player(n).SpellAnimations(i)
                    If .spellnum = 0 Then
                        .spellnum = spellnum
                        .Timer = GetTickCount + 120
                        .FramePointer = 0
                        Exit For
                    End If
                End With
            Next
            
        Case TARGET_TYPE_NPC
        
            For i = 1 To MAX_SPELLANIM
                With MapNpc(n).SpellAnimations(i)
                    If .spellnum = 0 Then
                        .spellnum = spellnum
                        .Timer = GetTickCount + 120
                        .FramePointer = 0
                        Exit For
                    End If
                End With
            Next
        
        Case TARGET_TYPE_NONE

    End Select

End Sub

Private Sub HandleDoor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    TempTile(Buffer.ReadLong, Buffer.ReadLong).DoorOpen = YES
End Sub

Private Sub HandleMaxes(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    MAX_PLAYERS = Buffer.ReadInteger
    MAX_ITEMS = Buffer.ReadInteger
    MAX_NPCS = Buffer.ReadInteger
    MAX_SHOPS = Buffer.ReadInteger
    MAX_SPELLS = Buffer.ReadInteger
    MAX_MAPS = Buffer.ReadInteger
    
    ReDim MapNpc(1 To MAX_MAP_NPCS, 1 To MAX_MAPS)
    ReDim MapItem(1 To MAX_MAP_ITEMS, 1 To MAX_MAPS)
    ReDim Player(1 To MAX_PLAYERS) As PlayerRec
    ReDim Item(1 To MAX_ITEMS) As ItemRec
    ReDim Npc(1 To MAX_NPCS) As NpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    
End Sub

Private Sub HandleSync(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim tx, ty, tm As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    tx = Buffer.ReadByte
    ty = Buffer.ReadByte
    tm = Buffer.ReadLong
    
    If SyncX = tx Then
        If SyncY = ty Then
            If SyncMap = tm Then
                SentSync = False
                Exit Sub
            End If
        End If
    End If
    
    Player(MyIndex).x = tx
    Player(MyIndex).y = ty
    Player(MyIndex).map = tm
    
    CalcTilePositions
    SentSync = False
End Sub


Private Sub HandleMapRevs(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim Buffer2 As clsBuffer
    
    Set Buffer2 = New clsBuffer
    
    Buffer2.PreAllocate MAX_MAPS * 1 + 2
    
    Buffer2.WriteInteger CMapReqs

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_MAPS
        If CheckMapRevision(i, Buffer.ReadLong) = False Then
            Buffer2.WriteByte 1
        Else
            Buffer2.WriteByte 0
        End If
    Next
    
    Call SendData(Buffer2.ToArray())
End Sub

