Attribute VB_Name = "modServerTCP"
Option Explicit

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
Dim i As Long

    If Index = 0 Then
        i = FindOpenPlayerSlot
        
        If i <> 0 Then
            ' Whoho, we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            SocketConnected i
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    ' Are they trying to connect more then one connection?
    'If Not IsMultiIPOnline(Current_IP(Index)) Then
        If Not IsBanned(Current_IP(Index)) Then
            AddText frmServer.txtText, "Received connection from " & Current_IP(Index) & "."
        Else
            SendAlertMsg Index, "Your account is banished."
        End If
    'Else
       ' Tried multiple connections
    '    SendAlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
    'End If
End Sub

Sub CloseSocket(ByVal Index As Long)
    ' Make sure player was/is playing the game, and if so, save'm.
    If Index > 0 Then
        LeftGame Index
    
        AddText frmServer.txtText, "Connection from " & Current_IP(Index) & " has been terminated."
        
        frmServer.Socket(Index).Close
        
        UpdateCaption
    End If
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then IsConnected = True
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If Player(Index).InGame = True Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If Trim$(Player(Index).Login) <> vbNullString Then
            IsLoggedIn = True
        End If
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) Then
            If LCase$(Current_Login(i)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If
    Next
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
Dim i As Long
Dim n As Long
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) Then
            If Trim$(Current_IP(i)) = Trim$(IP) Then
                n = n + 1
                
                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Function IsBanned(ByVal IP As String) As Boolean
Dim FileName As String, fIP As String, fName As String
Dim f As Long

    IsBanned = False
    
    FileName = App.Path & "\Data\banlist.txt"
    
    ' Check if file exists
    If Not FileExist(FileName, True) Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open FileName For Input As #f
    
    Do While Not EOF(f)
        Input #f, fIP
        Input #f, fName
    
        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If
    Loop
    
    Close #f
End Function

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    If Current_Access(Index) <= 0 Then
        ' Check for data flooding
        If Player(Index).DataBytes > 1000 Then
            HackingAttempt Index, "Data Flooding"
            Exit Sub
        End If
    
        ' Check for packet flooding
        If Player(Index).DataPackets > 25 Then
            HackingAttempt Index, "Packet Flooding"
            Exit Sub
        End If
    End If
            
    ' Check if elapsed time has passed
    Player(Index).DataBytes = Player(Index).DataBytes + DataLength
    If GetTickCount >= Player(Index).DataTimer Then
        Player(Index).DataTimer = GetTickCount + 1000
        Player(Index).DataBytes = 0
        Player(Index).DataPackets = 0
    End If
    
    
    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    
    Player(Index).Buffer.WriteBytes Buffer()
    
    If Player(Index).Buffer.Length >= 4 Then
        pLength = Player(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            HackingAttempt Index, "Hacking attempt."
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= Player(Index).Buffer.Length - 4
        If pLength <= Player(Index).Buffer.Length - 4 Then
            Player(Index).DataPackets = Player(Index).DataPackets + 1
            Player(Index).Buffer.ReadLong
            HandleData Index, Player(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If Player(Index).Buffer.Length >= 4 Then
            pLength = Player(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                HackingAttempt Index, "Hacking attempt."
                Exit Sub
            End If
        End If
    Loop
            
    Player(Index).Buffer.Trim
    
    
End Sub

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim tempData() As Byte

    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        tempData = Data
        If EncryptPackets Then
            Encryption_XOR_EncryptByte tempData(), PacketKeys(Player(Index).PacketOutIndex)
            Player(Index).PacketOutIndex = Player(Index).PacketOutIndex + 1
            If Player(Index).PacketOutIndex > PacketEncKeys - 1 Then Player(Index).PacketOutIndex = 0
        End If
        
        Buffer.PreAllocate 4 + (UBound(tempData) - LBound(tempData)) + 1
        Buffer.WriteLong (UBound(tempData) - LBound(tempData)) + 1
        Buffer.WriteBytes tempData()
        
        frmServer.Socket(Index).SendData Buffer.ToArray()
        
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
Dim i As Long

    For i = 1 To OnlinePlayersCount
        SendDataTo OnlinePlayers(i), Data()
    Next
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To OnlinePlayersCount
        If OnlinePlayers(i) <> Index Then
            SendDataTo OnlinePlayers(i), Data()
        End If
    Next
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To MapData(MapNum).MapPlayersCount
        SendDataTo MapData(MapNum).MapPlayers(i), Data()
    Next
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Long

'    For i = 1 To OnlinePlayersCount
'        If Current_Map(OnlinePlayers(i)) = MapNum Then
'            If OnlinePlayers(i) <> Index Then
'                SendDataTo OnlinePlayers(i), Data()
'            End If
'        End If
'    Next
    For i = 1 To MapData(MapNum).MapPlayersCount
        If MapData(MapNum).MapPlayers(i) <> Index Then
            SendDataTo MapData(MapNum).MapPlayers(i), Data()
        End If
    Next
End Sub

Sub SendDataToParty(PartyIndex As Long, ByRef Data() As Byte)
Dim i As Long
Dim n As Long
    For i = 1 To MAX_PLAYER_PARTY
        If LenB(Party(PartyIndex).PartyPlayers(i)) > 0 Then
            n = FindPlayer(Party(PartyIndex).PartyPlayers(i))
            If n > 0 Then
                SendDataTo n, Data
            End If
        End If
    Next
End Sub

Sub SendAlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Msg) + 8
    Buffer.WriteLong CMsgAlertMsg
    Buffer.WriteString Msg
    
    SendDataTo Index, Buffer.ToArray()
    DoEvents
    CloseSocket Index
End Sub

Sub SendClientMsg(ByVal Index As Long, ByVal Msg As String, ByVal MenuState As MenuStates)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Msg) + 12
    Buffer.WriteLong CMsgClientMsg
    Buffer.WriteString Msg
    Buffer.WriteLong MenuState
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendChars(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim FileName As String
Dim nFileNum As Integer
Dim Char As PlayerRec

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgAllChars

    FileName = AccountPath & "\" & Trim$(Player(Index).Login) & ".acc" 'Cool file extention

    nFileNum = FreeFile
    Open FileName For Binary As #nFileNum
        For i = 1 To MAX_CHARS
            Char.Name = vbNullString
            Char.GuildName = vbNullString
            Get #nFileNum, (NAME_LENGTH * 2) + (LenB(Char) * (i - 1)), Char
            Buffer.WriteString Trim$(Char.Name)
            Buffer.WriteString GetClassName(Char.Class)
            Buffer.WriteLong Char.Level
        Next
    Close #nFileNum
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendLoginOk(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgLoginOk
    Buffer.WriteLong Index
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long, n As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgNewCharClasses
    Buffer.WriteByte MAX_CLASSES
    For i = 0 To MAX_CLASSES
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxHP(i)
        Buffer.WriteLong GetClassMaxMP(i)
        Buffer.WriteLong GetClassMaxSP(i)
        For n = 1 To Stats.Stat_Count
            Buffer.WriteLong Class(i).Stat(n)
        Next
        Buffer.WriteString Class(i).MaleSprite
        Buffer.WriteString Class(i).FemaleSprite
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendClassesData(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long, n As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgClassesData
    Buffer.WriteByte MAX_CLASSES
    For i = 0 To MAX_CLASSES
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxHP(i)
        Buffer.WriteLong GetClassMaxMP(i)
        Buffer.WriteLong GetClassMaxSP(i)
        
        For n = 1 To Stats.Stat_Count
            Buffer.WriteLong Class(i).Stat(n)
        Next
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendInGame(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgInGame
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendPlayerInv(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_INV * 8) + 4
    Buffer.WriteLong CMsgPlayerInv
    For i = 1 To MAX_INV
       Buffer.WriteLong Current_InvItemNum(Index, i)
       Buffer.WriteLong Current_InvItemValue(Index, i)
       Buffer.WriteByte Current_InvItemBound(Index, i)
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendPlayerInvUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 16
    Buffer.WriteLong CMsgPlayerInvUpdate
    Buffer.WriteLong InvSlot
    Buffer.WriteLong Current_InvItemNum(Index, InvSlot)
    Buffer.WriteLong Current_InvItemValue(Index, InvSlot)
    Buffer.WriteByte Current_InvItemBound(Index, InvSlot)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendPlayerWornEq(ByVal Index As Long, ByVal EquipmentSlot As Slots)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong CMsgPlayerWornEq
    Buffer.WriteLong EquipmentSlot
    Buffer.WriteLong Current_EquipmentSlot(Index, EquipmentSlot)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong CMsgPlayerVital
    Buffer.WriteLong Vital
    Buffer.WriteLong Current_MaxVital(Index, Vital)
    Buffer.WriteLong Current_BaseVital(Index, Vital)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendStats(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate ((Stats.Stat_Count) * 8) + 12
    Buffer.WriteLong CMsgPlayerStats
    Buffer.WriteLong Current_Level(Index)
    Buffer.WriteLong Current_Points(Index)
    For i = 1 To Stats.Stat_Count
        Buffer.WriteLong Current_BaseStat(Index, i)
        Buffer.WriteLong Current_ModStat(Index, i)
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Public Sub SendPlayerData(ByVal Index As Long)
    SendDataTo Index, PlayerData(Index)
End Sub

Public Sub SendPlayerMove(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 20
    Buffer.WriteLong CMsgPlayerMove
    Buffer.WriteLong Index
    Buffer.WriteLong Current_X(Index)
    Buffer.WriteLong Current_Y(Index)
    Buffer.WriteLong Current_Dir(Index)

    SendDataToMapBut Index, Current_Map(Index), Buffer.ToArray()
End Sub

Sub SendNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Movement As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 24
    Buffer.WriteLong CMsgNpcMove
    Buffer.WriteLong MapNpcNum
    Buffer.WriteLong MapData(MapNum).MapNpc(MapNpcNum).X
    Buffer.WriteLong MapData(MapNum).MapNpc(MapNpcNum).Y
    Buffer.WriteLong MapData(MapNum).MapNpc(MapNpcNum).Dir
    Buffer.WriteLong Movement
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendPlayerDir(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong CMsgPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong Current_Dir(Index)
    
    SendDataToMapBut Index, Current_Map(Index), Buffer.ToArray()
End Sub

Sub SendNpcDir(ByVal MapNum As Long, MapNpcNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong CMsgNpcDir
    Buffer.WriteLong MapNpcNum
    Buffer.WriteLong MapData(MapNum).MapNpc(MapNpcNum).Dir
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendPlayerXY(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong CMsgPlayerXY
    Buffer.WriteLong Current_X(Index)
    Buffer.WriteLong Current_Y(Index)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendAttack(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgAttack
    Buffer.WriteLong Index
    
    SendDataToMapBut Index, Current_Map(Index), Buffer.ToArray()

End Sub

Sub SendNpcAttack(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgNpcAttack
    Buffer.WriteLong MapNpcNum
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendCheckForMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgCheckForMap
    Buffer.WriteLong MapNum
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    SendDataTo Index, MapCache(MapNum).Data()
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_MAP_ITEMS * 10) + 4
    Buffer.WriteLong CMsgMapItemData
    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteLong MapData(MapNum).MapItem(i).Num
        Buffer.WriteLong MapData(MapNum).MapItem(i).Value
        Buffer.WriteByte MapData(MapNum).MapItem(i).X
        Buffer.WriteByte MapData(MapNum).MapItem(i).Y
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_MAP_ITEMS * 10) + 4
    Buffer.WriteLong CMsgMapItemData
    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteLong MapData(MapNum).MapItem(i).Num
        Buffer.WriteLong MapData(MapNum).MapItem(i).Value
        Buffer.WriteByte MapData(MapNum).MapItem(i).X
        Buffer.WriteByte MapData(MapNum).MapItem(i).Y
    Next
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MapData(MapNum).NpcCount * 7) + 4
    Buffer.WriteLong CMsgMapNpcData
    Buffer.WriteLong MapData(MapNum).NpcCount
    For i = 1 To MapData(MapNum).NpcCount
        Buffer.WriteLong MapData(MapNum).MapNpc(i).Num
        Buffer.WriteByte MapData(MapNum).MapNpc(i).X
        Buffer.WriteByte MapData(MapNum).MapNpc(i).Y
        Buffer.WriteByte MapData(MapNum).MapNpc(i).Dir
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MapData(MapNum).NpcCount * 4) + 4
    Buffer.WriteLong CMsgMapNpcData
    Buffer.WriteLong MapData(MapNum).NpcCount
    For i = 1 To MapData(MapNum).NpcCount
        Buffer.WriteLong MapData(MapNum).MapNpc(i).Num
        Buffer.WriteByte MapData(MapNum).MapNpc(i).X
        Buffer.WriteByte MapData(MapNum).MapNpc(i).Y
        Buffer.WriteByte MapData(MapNum).MapNpc(i).Dir
    Next
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendMapDone(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgMapDone
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendGlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Msg) + 7
    Buffer.WriteLong CMsgChatMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendAdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Msg) + 7
    Buffer.WriteLong CMsgChatMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color

    For i = 1 To OnlinePlayersCount
        If Current_Access(OnlinePlayers(i)) > 0 Then
            SendDataTo OnlinePlayers(i), Buffer.ToArray()
        End If
    Next
'    For i = 1 To MAX_PLAYERS
'        If IsPlaying(i) Then
'            If Current_Access(i) > 0 Then
'                SendDataTo i, Buffer.ToArray()
'            End If
'        End If
'    Next
End Sub

Sub SendPlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Msg) + 7
    Buffer.WriteLong CMsgChatMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendMapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Msg) + 7
    Buffer.WriteLong CMsgChatMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendGuildMsg(ByVal GuildNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Msg) + 7
    Buffer.WriteLong CMsgChatMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color

    For i = 1 To OnlinePlayersCount
        If Current_Guild(OnlinePlayers(i)) = GuildNum Then
            SendDataTo OnlinePlayers(i), Buffer.ToArray()
        End If
    Next
End Sub

Sub SendPartyMsg(ByVal PartyIndex As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Msg) + 7
    Buffer.WriteLong CMsgChatMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color

    SendDataToParty PartyIndex, Buffer.ToArray()
End Sub

Sub SendSpawnItem(ByVal MapNum As Long, ByVal MapItemSlot As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 24
    Buffer.WriteLong CMsgSpawnItem
    Buffer.WriteLong MapItemSlot
    Buffer.WriteLong MapData(MapNum).MapItem(MapItemSlot).Num
    Buffer.WriteLong MapData(MapNum).MapItem(MapItemSlot).Value
    Buffer.WriteLong MapData(MapNum).MapItem(MapItemSlot).X
    Buffer.WriteLong MapData(MapNum).MapItem(MapItemSlot).Y
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

'////////////////////
'//  Item Packets  //
'////////////////////

Sub SendItemEditor(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgItemEditor
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendItems(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(ItemsCache) - LBound(ItemsCache)) + 3
    Buffer.WriteLong CMsgUpdateItems
    Buffer.WriteBytes ItemsCache()
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate ItemSize + 8
    Buffer.WriteLong CMsgUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes Get_ItemData(ItemNum)
    
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate ItemSize + 8
    Buffer.WriteLong CMsgUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes Get_ItemData(ItemNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate ItemSize + 8
    Buffer.WriteLong CMsgEditItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes Get_ItemData(ItemNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

'///////////////////////////
'//  Emotioicons Packets  //
'///////////////////////////

Sub SendEmoticonEditor(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgEmoticonEditor
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendEmoticons(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(EmoticonsCache) - LBound(EmoticonsCache)) + 3
    Buffer.WriteLong CMsgUpdateEmoticons
    Buffer.WriteBytes EmoticonsCache()
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal EmoticonNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate EmoticonSize + 8
    Buffer.WriteLong CMsgUpdateEmoticon
    Buffer.WriteLong EmoticonNum
    Buffer.WriteBytes Get_EmoticonData(EmoticonNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendUpdateEmoticonToAll(ByVal EmoticonNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate EmoticonSize + 8
    Buffer.WriteLong CMsgUpdateEmoticon
    Buffer.WriteLong EmoticonNum
    Buffer.WriteBytes Get_EmoticonData(EmoticonNum)
    
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendEditEmoticonTo(ByVal Index As Long, ByVal EmoticonNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
      
    Buffer.PreAllocate EmoticonSize + 8
    Buffer.WriteLong CMsgEditEmoticon
    Buffer.WriteLong EmoticonNum
    Buffer.WriteBytes Get_EmoticonData(EmoticonNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendCheckEmoticon(ByVal Index As Long, ByVal MapNum As Long, ByVal EmoticonNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgCheckEmoticon
    Buffer.WriteLong Index
    Buffer.WriteLong EmoticonNum
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendTarget(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteLong CMsgNewTarget
    Buffer.WriteByte Player(Index).Target
    Buffer.WriteByte Player(Index).TargetType

    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendSpawnNpc(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 24
    Buffer.WriteLong CMsgSpawnNpc
    Buffer.WriteLong MapNpcNum
    Buffer.WriteLong MapData(MapNum).MapNpc(MapNpcNum).Num
    Buffer.WriteLong MapData(MapNum).MapNpc(MapNpcNum).X
    Buffer.WriteLong MapData(MapNum).MapNpc(MapNpcNum).Y
    Buffer.WriteLong MapData(MapNum).MapNpc(MapNpcNum).Dir
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendNpcDead(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgNpcDead
    Buffer.WriteLong MapNpcNum
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

'///////////////////
'//  NPC Packets  //
'///////////////////

Sub SendNpcEditor(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgNpcEditor
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendNpcs(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(NpcsCache) - LBound(NpcsCache)) + 3
    Buffer.WriteLong CMsgUpdateNpcs
    Buffer.WriteBytes NpcsCache()
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteInteger Npc(NpcNum).Sprite
    Buffer.WriteByte Npc(NpcNum).Behavior
    Buffer.WriteByte Npc(NpcNum).MovementSpeed
    Buffer.WriteByte Npc(NpcNum).MovementFrequency
    
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong CMsgUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteInteger Npc(NpcNum).Sprite
    Buffer.WriteByte Npc(NpcNum).Behavior
    Buffer.WriteByte Npc(NpcNum).MovementSpeed
    Buffer.WriteByte Npc(NpcNum).MovementFrequency

    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate NpcSize + 8
    Buffer.WriteLong CMsgEditNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes Get_NpcData(NpcNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendMapKey(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Key As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 13
    Buffer.WriteLong CMsgMapKey
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte Key
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendEditMap(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgEditMap
    
    SendDataTo Index, Buffer.ToArray()
End Sub

'////////////////////
'//  Shop Packets  //
'////////////////////

Sub SendShopEditor(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgShopEditor
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendShops(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(ShopsCache) - LBound(ShopsCache)) + 5
    Buffer.WriteLong CMsgUpdateShops
    Buffer.WriteBytes ShopsCache()
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate ShopSize + 8
    Buffer.WriteLong CMsgUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteBytes Get_ShopData(ShopNum)
    
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate ShopSize + 8
    Buffer.WriteLong CMsgUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteBytes Get_ShopData(ShopNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate ShopSize + 8
    Buffer.WriteLong CMsgEditShop
    Buffer.WriteLong ShopNum
    Buffer.WriteBytes Get_ShopData(ShopNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

'/////////////////////
'//  Spell Packets  //
'/////////////////////

Sub SendSpellEditor(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgSpellEditor
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendSpells(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(SpellsCache) - LBound(SpellsCache)) + 3
    Buffer.WriteLong CMsgUpdateSpells
    Buffer.WriteBytes SpellsCache()
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate SpellSize + 8
    Buffer.WriteLong CMsgUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes Get_SpellData(SpellNum)
    
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate SpellSize + 8
    Buffer.WriteLong CMsgUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes Get_SpellData(SpellNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate SpellSize + 8
    Buffer.WriteLong CMsgEditSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes Get_SpellData(SpellNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

'/////////////////////////
'//  Animation Packets  //
'/////////////////////////

Sub SendAnimationEditor(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgAnimationEditor
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendAnimations(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(AnimationsCache) - LBound(AnimationsCache)) + 3
    Buffer.WriteLong CMsgUpdateAnimations
    Buffer.WriteBytes AnimationsCache()
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate AnimationSize + 8
    Buffer.WriteLong CMsgUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes Get_AnimationData(AnimationNum)
    
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate AnimationSize + 8
    Buffer.WriteLong CMsgUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes Get_AnimationData(AnimationNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendEditAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate AnimationSize + 8
    Buffer.WriteLong CMsgEditAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes Get_AnimationData(AnimationNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendTrade(ByVal Index As Long, ByVal MapNpcNum As Byte, ByVal ShopNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 9
    Buffer.WriteLong CMsgTrade
    Buffer.WriteByte MapNpcNum
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    Buffer.WriteString Trim$(Shop(ShopNum).JoinSay)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_PLAYER_SPELLS * 8) + 4
    Buffer.WriteLong CMsgSpells
    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong Current_Spell(Index, i)
        Buffer.WriteLong Current_SpellCooldown(Index, i)
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendActionMsg(ByVal MapNum As Long, ByVal Message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal X As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgActionMsg
    Buffer.WriteString Message
    Buffer.WriteLong Color
    Buffer.WriteLong MsgType
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
End Sub

Sub SendAnimation(ByVal MapNum As Long, ByVal AnimationNum As Byte, ByVal X As Long, ByVal Y As Long)
Dim Buffer As clsBuffer
    
    If AnimationNum <= 0 Then Exit Sub
    If AnimationNum > MAX_ANIMATIONS Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 13
    Buffer.WriteLong CMsgAnimation
    Buffer.WriteByte AnimationNum
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendPlayerGuild(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgPlayerGuild
    Buffer.WriteLong Index
    Buffer.WriteString Current_GuildName(Index)
    Buffer.WriteString GetGuildAbbreviation(Current_Guild(Index))
    
    SendDataToMap Current_Map(Index), Buffer.ToArray()
End Sub

Public Sub SendPlayerExp(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong CMsgPlayerExp
    Buffer.WriteLong Current_Exp(Index)
    Buffer.WriteLong Current_NextLevel(Index)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendCancelSpell(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgCancelSpell
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Public Sub SendSpellReady(ByVal Index As Long, ByVal SpellSlot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgSpellReady
    Buffer.WriteLong SpellSlot

    SendDataTo Index, Buffer.ToArray()
End Sub

Public Sub SendSpellCooldown(ByVal Index As Long, ByVal SpellSlot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgSpellCooldown
    Buffer.WriteLong SpellSlot
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgLeftGame
    Buffer.WriteLong Index
    
    SendDataToMapBut Index, MapNum, Buffer.ToArray()
End Sub

Sub SendLeftGame(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgLeftGame
    Buffer.WriteLong Index
    
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendPlayerDead(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgPlayerDead
    Buffer.WriteLong Index
    Buffer.WriteLong Current_IsDeadTimer(Index) - GetTickCount
    
    SendDataToMap Current_Map(Index), Buffer.ToArray()
End Sub

Sub SendPlayerGold(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong CMsgPlayerGold
    Buffer.WriteLong Current_Gold(Index)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendPlayerRevival(ByVal Index As Long, ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong CMsgPlayerRevival
    Buffer.WriteString Name
    
    SendDataTo Index, Buffer.ToArray()
End Sub


Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            SendGlobalMsg "[Realm Event] " & Current_Login(Index) & "/" & Current_Name(Index) & " has been cast from the realm. (" & Reason & ")", Yellow
        End If
    
        CloseSocket Index
    End If
End Sub

Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim i As Long

    For i = 1 To OnlinePlayersCount
        If OnlinePlayers(i) <> Index Then
            s = s & Current_Name(OnlinePlayers(i)) & ", "
        End If
    Next
            
    If OnlinePlayersCount - 1 = 0 Then
        s = "There are no other adventurers in the realm."
    Else
        s = "There are " & OnlinePlayersCount - 1 & " other adventurers in the realm: " & Left$(s, Len(s) - 2) & "."
    End If
        
    SendPlayerMsg Index, s, AlertColor
End Sub

Sub SendJoinMap(ByVal Index As Long)
Dim i As Long
Dim MapNum As Long
    
    MapNum = Current_Map(Index)
    
    ' Send all players on current map to index
    For i = 1 To MapData(MapNum).MapPlayersCount
        If MapData(MapNum).MapPlayers(i) <> Index Then
            SendDataTo Index, PlayerData(MapData(MapNum).MapPlayers(i))
        End If
    Next
'    For i = 1 To OnlinePlayersCount
'        If OnlinePlayers(i) <> Index Then
'            If Current_Map(OnlinePlayers(i)) = Current_Map(Index) Then
'                SendDataTo Index, PlayerData(OnlinePlayers(i))
'            End If
'        End If
'    Next
    
    SendDataToMap MapNum, PlayerData(Index)
End Sub

Sub SendWelcome(ByVal Index As Long)

    ' Send them MOTD
    If Trim$(GameMOTD) <> vbNullString Then
        SendPlayerMsg Index, "[Realm News] " & GameMOTD, Yellow
    End If
    
    ' Send Guild MOTD if in guild
    If Current_Guild(Index) > 0 Then
        If Trim$(GetGuildGMOTD(Current_Guild(Index))) <> vbNullString Then
            SendPlayerMsg Index, "[" & Current_GuildName(Index) & "] " & Trim$(GetGuildGMOTD(Current_Guild(Index))), BrightGreen
        End If
    End If
    
    ' Send whos online
    SendWhosOnline Index
End Sub

'/////////////////////
'//  Quest Packets  //
'/////////////////////

Sub SendQuestEditor(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong CMsgQuestEditor
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendQuests(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(QuestsCache) - LBound(QuestsCache)) + 3
    Buffer.WriteLong CMsgUpdateQuests
    Buffer.WriteBytes QuestsCache()
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate QuestSize + 8
    Buffer.WriteLong CMsgUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes Get_QuestData(QuestNum)
    
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal QuestNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.PreAllocate QuestSize + 8
    Buffer.WriteLong CMsgUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes Get_QuestData(QuestNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendEditQuestTo(ByVal Index As Long, ByVal QuestNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate QuestSize + 8
    Buffer.WriteLong CMsgEditQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes Get_QuestData(QuestNum)
    
    SendDataTo Index, Buffer.ToArray()
End Sub



