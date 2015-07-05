Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
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
    'HandleDataSub(SDoor) = GetAddress(AddressOf HandleDoor)
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        MsgBox "Packet Error.", vbOKOnly
        DestroyGame
        Exit Sub
    End If

    If MsgType >= CMSG_COUNT Then
        MsgBox "Packet Error: MsgType = " & MsgType & ".", vbOKOnly
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.Length - 4 + 1), 0, 0
    
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::
' :: Alert message packet ::
' ::::::::::::::::::::::::::
Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

    frmSendGetData.Visible = False
    frmMainMenu.Visible = True
     
    Msg = Buffer.ReadString 'Parse(1)
    
    Set Buffer = Nothing
    
    Call MsgBox(Msg, vbOKOnly, GAME_NAME)
End Sub

' :::::::::::::::::::::::::::
' :: All characters packet ::
' :::::::::::::::::::::::::::
Sub HandleAllChars(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Level As Long
Dim Name As String
Dim Msg As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

     n = 1
     
     frmChars.Visible = True
     frmSendGetData.Visible = False
     
     frmChars.lstChars.Clear
     
     For i = 1 To MAX_CHARS
         Name = Buffer.ReadString
         Msg = Buffer.ReadString
         Level = Buffer.ReadLong
         
         If Trim$(Name) = vbNullString Then
             frmChars.lstChars.AddItem "Free Character Slot"
         Else
             frmChars.lstChars.AddItem Name & " a level " & Level & " " & Msg
         End If
         
         n = n + 3
     Next
     
     Set Buffer = Nothing
     
     frmChars.lstChars.ListIndex = 0
End Sub

' :::::::::::::::::::::::::::::::::
' :: Login was successful packet ::
' :::::::::::::::::::::::::::::::::
Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
     ' Now we can receive game data
     MyIndex = Buffer.ReadLong 'CLng(Parse(1))
     
     Set Buffer = Nothing
     
     frmSendGetData.Visible = True
     frmChars.Visible = False
     
     Call SetStatus("Receiving game data...")
     
End Sub

' :::::::::::::::::::::::::::::::::::::::
' :: New character classes data packet ::
' :::::::::::::::::::::::::::::::::::::::
Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
     
     n = 1
     
     ' Max classes
     Max_Classes = Buffer.ReadLong
     ReDim Class(1 To Max_Classes)
     
     n = n + 1
     
     For i = 1 To Max_Classes
         With Class(i)
             .Name = Buffer.ReadString
             
             .Vital(Vitals.HP) = Buffer.ReadLong 'CLng(Parse(n + 1))
             .Vital(Vitals.MP) = Buffer.ReadLong 'CLng(Parse(n + 2))
             .Vital(Vitals.SP) = Buffer.ReadLong 'CLng(Parse(n + 3))
             
             .Stat(Stats.Strength) = Buffer.ReadLong 'CLng(Parse(n + 4))
             .Stat(Stats.Defense) = Buffer.ReadLong 'CLng(Parse(n + 5))
             .Stat(Stats.SPEED) = Buffer.ReadLong 'CLng(Parse(n + 6))
             .Stat(Stats.Magic) = Buffer.ReadLong 'CLng(Parse(n + 7))
         End With
         
         n = n + 8
     Next
     
     Set Buffer = Nothing
     
     ' Used for if the player is creating a new character
     frmNewChar.Visible = True
     frmSendGetData.Visible = False

     frmNewChar.cmbClass.Clear

     For i = 1 To Max_Classes
         frmNewChar.cmbClass.AddItem Trim$(Class(i).Name)
     Next

     frmNewChar.cmbClass.ListIndex = 0
     
     n = frmNewChar.cmbClass.ListIndex + 1
     
     frmNewChar.lblHP.Caption = CStr(Class(n).Vital(Vitals.HP))
     frmNewChar.lblMP.Caption = CStr(Class(n).Vital(Vitals.MP))
     frmNewChar.lblSP.Caption = CStr(Class(n).Vital(Vitals.SP))
 
     frmNewChar.lblStrength.Caption = CStr(Class(n).Stat(Stats.Strength))
     frmNewChar.lblDefense.Caption = CStr(Class(n).Stat(Stats.Defense))
     frmNewChar.lblSpeed.Caption = CStr(Class(n).Stat(Stats.SPEED))
     frmNewChar.lblMagic.Caption = CStr(Class(n).Stat(Stats.Magic))
End Sub

' :::::::::::::::::::::::::
' :: Classes data packet ::
' :::::::::::::::::::::::::
Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
     
     n = 1
     
     ' Max classes
     Max_Classes = Buffer.ReadLong 'CByte(Parse(n))
     ReDim Class(1 To Max_Classes)
     
     n = n + 1
     
     For i = 1 To Max_Classes
         With Class(i)
             .Name = Buffer.ReadString 'Trim$(Parse(n))
             
             .Vital(Vitals.HP) = Buffer.ReadLong 'CLng(Parse(n + 1))
             .Vital(Vitals.MP) = Buffer.ReadLong 'CLng(Parse(n + 2))
             .Vital(Vitals.SP) = Buffer.ReadLong 'CLng(Parse(n + 3))
             
             .Stat(Stats.Strength) = Buffer.ReadLong 'CLng(Parse(n + 4))
             .Stat(Stats.Defense) = Buffer.ReadLong 'CLng(Parse(n + 5))
             .Stat(Stats.SPEED) = Buffer.ReadLong 'CLng(Parse(n + 6))
             .Stat(Stats.Magic) = Buffer.ReadLong 'CLng(Parse(n + 7))
         End With
         
         n = n + 8
     Next
     
     Set Buffer = Nothing
     
End Sub

' ::::::::::::::::::::
' :: In game packet ::
' ::::::::::::::::::::
Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
     InGame = True
     Call GameInit
     Call GameLoop
End Sub

' :::::::::::::::::::::::::::::
' :: Player inventory packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

     n = 1
     For i = 1 To MAX_INV
         Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
         Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
         Call SetPlayerInvItemDur(MyIndex, i, Buffer.ReadLong)
         
         n = n + 3
     Next
     
     Set Buffer = Nothing
     
     Call UpdateInventory
 End Sub

' ::::::::::::::::::::::::::::::::::::
' :: Player inventory update packet ::
' ::::::::::::::::::::::::::::::::::::
Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

     n = Buffer.ReadLong 'CLng(Parse(1))
     
     Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(2)))
     Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(3)))
     Call SetPlayerInvItemDur(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(4)))
     Call UpdateInventory
     
     Set Buffer = Nothing
     
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player worn equipment packet ::
' ::::::::::::::::::::::::::::::::::
Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
     Call SetPlayerEquipmentSlot(MyIndex, Buffer.ReadLong, Armor)
     Call SetPlayerEquipmentSlot(MyIndex, Buffer.ReadLong, Weapon)
     Call SetPlayerEquipmentSlot(MyIndex, Buffer.ReadLong, Helmet)
     Call SetPlayerEquipmentSlot(MyIndex, Buffer.ReadLong, Shield)
     Call UpdateInventory
     
     Set Buffer = Nothing
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
        frmMirage.lblHP.Caption = Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%"
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
        frmMirage.lblMP.Caption = Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%"
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
        frmMirage.lblSP.Caption = Int(GetPlayerVital(MyIndex, Vitals.SP) / GetPlayerMaxVital(MyIndex, Vitals.SP) * 100) & "%"
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
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    
    ' Check if the player is the client player, and if so reset directions
    If i = MyIndex Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
    End If
    
    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    
    Call GetPlayersOnMap
End Sub

' ::::::::::::::::::::::::::::
' :: Player movement packet ::
' ::::::::::::::::::::::::::::
Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim n As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    n = Buffer.ReadLong
    
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
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
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Movement As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    MapNpcNum = Buffer.ReadLong
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
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

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    i = Buffer.ReadLong
    Dir = Buffer.ReadLong
    
    With MapNpc(i)
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
Dim X As Long
Dim Y As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()

     X = Buffer.ReadLong
     Y = Buffer.ReadLong
     
     Call SetPlayerX(MyIndex, X)
     Call SetPlayerY(MyIndex, Y)
     
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
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount
End Sub

' ::::::::::::::::::::::::::
' :: Check for map packet ::
' ::::::::::::::::::::::::::
Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
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
     X = Buffer.ReadLong
     
     ' Get revision
     Y = Buffer.ReadLong
     
     If FileExist(MAP_PATH & "map" & X & MAP_EXT, False) Then
         Call LoadMap(X)
     
         ' Check to see if the revisions match
         NeedMap = 1
         If Map.Revision = Y Then
             ' We do so we dont need the map
             'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
             NeedMap = 0
         End If
     End If
     
     ' Either the revisions didn't match or we dont have the map, so we need it
     Set Buffer = New clsBuffer
     'buffer.preallocate 3
     Buffer.WriteLong CNeedMap
     Buffer.WriteLong NeedMap
     SendData Buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim Buffer As clsBuffer
Dim MapNum As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

    Buffer.DecompressBuffer
    
     n = 1
     
     MapNum = Buffer.ReadLong
     Map.Name = Buffer.ReadString
     Map.Revision = Buffer.ReadLong 'CLng(Parse(n + 2))
     Map.Moral = Buffer.ReadLong 'CByte(Parse(n + 3))
     Map.TileSet = Buffer.ReadLong 'CInt(Parse(n + 4))
     Map.Up = Buffer.ReadLong 'CInt(Parse(n + 5))
     Map.Down = Buffer.ReadLong 'CInt(Parse(n + 6))
     Map.Left = Buffer.ReadLong 'CInt(Parse(n + 7))
     Map.Right = Buffer.ReadLong 'CInt(Parse(n + 8))
     Map.Music = Buffer.ReadLong 'CByte(Parse(n + 9))
     Map.BootMap = Buffer.ReadLong 'CInt(Parse(n + 10))
     Map.BootX = Buffer.ReadLong 'CByte(Parse(n + 11))
     Map.BootY = Buffer.ReadLong 'CByte(Parse(n + 12))
     Map.Shop = Buffer.ReadLong 'CByte(Parse(n + 13))
     Map.MaxX = Buffer.ReadLong 'CByte(Parse(n + 14))
     Map.MaxY = Buffer.ReadLong 'CByte(Parse(n + 15))
     
     ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
     
     n = n + 16
     For X = 0 To Map.MaxX
         For Y = 0 To Map.MaxY
             Map.Tile(X, Y).Ground = Buffer.ReadLong 'CInt(Parse(n))
             Map.Tile(X, Y).Mask = Buffer.ReadLong 'CInt(Parse(n + 1))
             Map.Tile(X, Y).Mask2 = Buffer.ReadLong 'CInt(Parse(n + 2))
             Map.Tile(X, Y).Anim = Buffer.ReadLong 'CInt(Parse(n + 3))
             Map.Tile(X, Y).Fringe = Buffer.ReadLong 'CInt(Parse(n + 4))
             Map.Tile(X, Y).Fringe2 = Buffer.ReadLong 'CInt(Parse(n + 5))
             Map.Tile(X, Y).Type = Buffer.ReadLong 'CByte(Parse(n + 6))
             Map.Tile(X, Y).Data1 = Buffer.ReadLong 'CInt(Parse(n + 7))
             Map.Tile(X, Y).Data2 = Buffer.ReadLong 'CInt(Parse(n + 8))
             Map.Tile(X, Y).Data3 = Buffer.ReadLong 'CInt(Parse(n + 9))
             n = n + 10
         Next
     Next
     
     For X = 1 To MAX_MAP_NPCS
         Map.Npc(X) = Buffer.ReadLong 'CByte(Parse(n))
         n = n + 1
     Next
     
     ClearTempTile
             
     ' Save the map
     Call SaveMap(Buffer.ReadLong) 'CLng(Parse(1)))
     
     Set Buffer = Nothing
     
     ' Check if we get a map from someone else and if we were editing a map cancel it out
     If InMapEditor Then
         InMapEditor = False
         frmMirage.picMapEditor.Visible = False
         
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

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
     
    For i = 1 To MAX_MAP_ITEMS
       With MapItem(i)
           .Num = Buffer.ReadLong
           .Value = Buffer.ReadLong
           .Dur = Buffer.ReadLong
           .X = Buffer.ReadLong
           .Y = Buffer.ReadLong
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
     
    For i = 1 To MAX_MAP_NPCS
       With MapNpc(i)
           .Num = Buffer.ReadLong
           .X = Buffer.ReadLong
           .Y = Buffer.ReadLong
           .Dir = Buffer.ReadLong
       End With
    Next
End Sub

' :::::::::::::::::::::::::::::::
' :: Map send completed packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandleMapDone()
Dim i As Long
Dim MusicFile As String

    MusicFile = Trim$(CStr(Map.Music)) & ".mid"
'
'    ' get high NPC index
'    High_Npc_Index = 0
'    For i = 1 To MAX_MAP_NPCS
'        If Map.Npc(i) > 0 Then
'            High_Npc_Index = High_Npc_Index + 1
'        Else
'            Exit For
'        End If
'    Next
    
    ' Play music
    If Map.Music > 0 Then
        If MusicFile <> CurrentMusic Then
            DirectMusic_StopMidi
            Call DirectMusic_PlayMidi(MusicFile)
            CurrentMusic = MusicFile
        End If
    Else
        DirectMusic_StopMidi
        CurrentMusic = 0
    End If
    
    Call UpdateDrawMapName
    
    If Map.Shop = 0 Then
        frmMirage.picTradeButton.Visible = False
    Else
        frmMirage.picTradeButton.Visible = True
    End If
    
    'Call CalcTilePositions
    
    Call InitTileSurf(Map.TileSet)
    
    GettingMap = False
    CanMoveNow = True
End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    
    Call AddText(Msg, Color)
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    
    Call AddText(Msg, Color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    
    Call AddText(Msg, Color)
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    
    Call AddText(Msg, Color)
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    
    Call AddText(Msg, Color)
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim Color As Byte

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    
    Call AddText(Msg, Color)
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

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    With MapItem(n)
       .Num = Buffer.ReadLong
       .Value = Buffer.ReadLong
       .Dur = Buffer.ReadLong
       .X = Buffer.ReadLong
       .Y = Buffer.ReadLong
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
       .Pic = Buffer.ReadLong
       .Type = Buffer.ReadLong
       .Data1 = Buffer.ReadLong
       .Data2 = Buffer.ReadLong
       .Data3 = Buffer.ReadLong
    End With
End Sub

' ::::::::::::::::::::::
' :: Edit item packet ::
' ::::::::::::::::::::::
Sub HandleEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

     Set Buffer = New clsBuffer
     
     Buffer.WriteBytes Data()

     n = Buffer.ReadLong
     
     ' Update the item
     Item(n).Name = Buffer.ReadString
     Item(n).Pic = Buffer.ReadLong
     Item(n).Type = Buffer.ReadLong
     Item(n).Data1 = Buffer.ReadLong
     Item(n).Data2 = Buffer.ReadLong
     Item(n).Data3 = Buffer.ReadLong
     'Item(n).ClassReq = buffer.readlong
     'Item(n).data4 = Buffer.ReadLong
     'Item(n).Proficiency = Buffer.ReadLong
     
     ' Initialize the item editor
     Call ItemEditorInit
     
     Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Npc spawn packet ::
' ::::::::::::::::::::::
Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    
    With MapNpc(n)
       .Num = Buffer.ReadLong
       .X = Buffer.ReadLong
       .Y = Buffer.ReadLong
       .Dir = Buffer.ReadLong
       
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

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    Call ClearMapNpc(n)
    
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
       .Sprite = Buffer.ReadLong
       .SpawnSecs = 0
       .Behavior = 0
       .Range = 0
       .DropChance = 0
       .DropItem = 0
       .DropItemValue = 0
       .Stat(Stats.Strength) = 0
       .Stat(Stats.Defense) = 0
       .Stat(Stats.SPEED) = 0
       .Stat(Stats.Magic) = 0
    End With
End Sub

' :::::::::::::::::::::
' :: Edit npc packet ::
' :::::::::::::::::::::
Sub HandleEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

     Set Buffer = New clsBuffer
     
     Buffer.WriteBytes Data()

     n = Buffer.ReadLong
     
     ' Update the npc
     Npc(n).Name = Buffer.ReadString
     Npc(n).AttackSay = Buffer.ReadString
     'Npc(n).MaxHP = Buffer.ReadLong
     'Npc(n).GiveEXP = Buffer.ReadLong
     Npc(n).Sprite = Buffer.ReadLong
     Npc(n).SpawnSecs = Buffer.ReadLong
     Npc(n).Behavior = Buffer.ReadLong
     Npc(n).Range = Buffer.ReadLong
     Npc(n).DropChance = Buffer.ReadLong
     Npc(n).DropItem = Buffer.ReadLong
     Npc(n).DropItemValue = Buffer.ReadLong
     Npc(n).Stat(Stats.Strength) = Buffer.ReadLong
     Npc(n).Stat(Stats.Defense) = Buffer.ReadLong
     Npc(n).Stat(Stats.SPEED) = Buffer.ReadLong
     Npc(n).Stat(Stats.Magic) = Buffer.ReadLong
     'Npc(n).TrainPro = buffer.readlong
     
     ' Initialize the npc editor
     Call NpcEditorInit
     
     Set Buffer = Nothing
End Sub

' ::::::::::::::::::::
' :: Map key packet ::
' ::::::::::::::::::::
Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    n = Buffer.ReadLong
    
    TempTile(X, Y).DoorOpen = n
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
Sub HandleEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
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
     
     ' Update the shop
     Shop(ShopNum).Name = Buffer.ReadString
     Shop(ShopNum).JoinSay = Buffer.ReadString
     Shop(ShopNum).LeaveSay = Buffer.ReadString
     Shop(ShopNum).FixesItems = Buffer.ReadLong
     
     n = 6
     For i = 1 To MAX_TRADES
         
         GiveItem = Buffer.ReadLong
         GiveValue = Buffer.ReadLong
         GetItem = Buffer.ReadLong
         GetValue = Buffer.ReadLong
         
         Shop(ShopNum).TradeItem(i).GiveItem = GiveItem
         Shop(ShopNum).TradeItem(i).GiveValue = GiveValue
         Shop(ShopNum).TradeItem(i).GetItem = GetItem
         Shop(ShopNum).TradeItem(i).GetValue = GetValue
         
         n = n + 4
     Next
     
     ' Initialize the shop editor
     Call ShopEditorInit
     
     Set Buffer = Nothing
     
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
       .MPReq = Buffer.ReadLong
       .Pic = Buffer.ReadLong
    End With
End Sub

' :::::::::::::::::::::::
' :: Edit spell packet ::
' :::::::::::::::::::::::
Sub HandleEditSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long, i As Byte, t As Byte
Dim Buffer As clsBuffer

     Set Buffer = New clsBuffer
     
     Buffer.WriteBytes Data()

     n = Buffer.ReadLong
     
     ' Update the spell
     Spell(n).Name = Buffer.ReadString
     Spell(n).Pic = Buffer.ReadLong
     Spell(n).ClassReq = Buffer.ReadLong
     Spell(n).Type = Buffer.ReadLong
     
     t = 6
     
     'For i = 1 To MAX_SPELL_LEVEL
        Spell(n).MPReq = Buffer.ReadLong
        Spell(n).LevelReq = Buffer.ReadLong
        Spell(n).Data1 = Buffer.ReadLong
        't = t + 3
     'Next
     
     Spell(n).Data2 = Buffer.ReadLong
     Spell(n).Data3 = Buffer.ReadLong
     
     'For i = 1 To MAX_SPELL_LEVEL
        'Spell(n).Range = Buffer.ReadLong
     'Next
     
     ' Initialize the spell editor
     Call SpellEditorInit
     
     Set Buffer = Nothing
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
       If Buffer.ReadLong = 1 Then
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
Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim j As Long
Dim n As Long
Dim k As Long
Dim Buffer As clsBuffer

     Set Buffer = New clsBuffer
     
     Buffer.WriteBytes Data()
     
     frmMirage.picSpellsList.Visible = True
     frmMirage.lstSpells.Clear
     
     n = 1
     
     For i = 1 To MAX_PLAYER_SPELLS
        'k = Buffer.ReadLong
        PlayerSpells(i) = Buffer.ReadLong
        'PlayerSpellLevels(i) = Buffer.ReadLong
        n = n + 3
     Next
     
     ' Put spells known in player record
     For i = 1 To MAX_PLAYER_SPELLS
         If PlayerSpells(i) <> 0 Then
             frmMirage.lstSpells.AddItem i & ": " & Trim$(Spell(PlayerSpells(i)).Name) '& " Lvl. " & PlayerSpellLevels(i)
         Else
             frmMirage.lstSpells.AddItem "[free spells slot]"
         End If
     Next
     
     frmMirage.lstSpells.ListIndex = 0
     
     Set Buffer = Nothing
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
Dim SpellNum As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
     
    Buffer.WriteBytes Data()
    
    TargetType = Buffer.ReadLong
    n = Buffer.ReadLong
    SpellNum = Buffer.ReadLong
    
    If n = 0 Or SpellNum = 0 Then
        Exit Sub
    End If
    
    Select Case TargetType
        Case TARGET_TYPE_PLAYER
        
            For i = 1 To MAX_SPELLANIM
                With Player(n).SpellAnimations(i)
                    If .SpellNum = 0 Then
                        .SpellNum = SpellNum
                        .Timer = GetTickCount + 120
                        .FramePointer = 0
                        Exit For
                    End If
                End With
            Next
            
        Case TARGET_TYPE_NPC
        
            For i = 1 To MAX_SPELLANIM
                With MapNpc(n).SpellAnimations(i)
                    If .SpellNum = 0 Then
                        .SpellNum = SpellNum
                        .Timer = GetTickCount + 120
                        .FramePointer = 0
                        Exit For
                    End If
                End With
            Next
        
        Case TARGET_TYPE_NONE

    End Select

End Sub

