Attribute VB_Name = "modServerTCP"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Sub UpdateCaption()
    frmServer.Caption = "Mirage Source Server <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Sub CreateFullMapCache()
    Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next
    
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If LenB(Trim$(Player(Index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
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
            If Trim$(GetPlayerIP(i)) = IP Then
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
Dim FileName As String
Dim fIP As String
Dim fName As String
Dim F As Long

    FileName = App.Path & "\data\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    F = FreeFile
    Open FileName For Input As #F
    
    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName
    
        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If
    Loop
    
    Close #F
End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer

    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data()
        
        'Buffer.CompressBuffer
        
        frmServer.Socket(Index).SendData Buffer.ToArray()
    
        Set Buffer = Nothing
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> Index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    'buffer.preallocate Len(Msg) + 5
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    
    SendDataToAll Buffer.ToArray
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    'buffer.preallocate Len(Msg) + 5
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    
    For i = 1 To TotalPlayersOnline
        If GetPlayerAccess(PlayersOnline(i)) > 0 Then
            SendDataTo PlayersOnline(i), Buffer.ToArray
        End If
    Next
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    'buffer.preallocate Len(Msg) + 5
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    
    SendDataTo Index, Buffer.ToArray
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    'buffer.preallocate Len(Msg) + 5
    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    
    SendDataToMap MapNum, Buffer.ToArray
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    'buffer.preallocate Len(Msg) + 4
    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    
    SendDataTo Index, Buffer.ToArray
    DoEvents
    Call CloseSocket(Index)
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
Dim i As Long

    If (Index = 0) Then
        i = FindOpenPlayerSlot
        
        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    If Index <> 0 Then
        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(Index)) Then
                Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".")
            Else
                Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
            End If
        'Else
           ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim Data() As Byte
Dim pLength As Long

    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    
    'IncomingBytes = IncomingBytes + DataLength
    
    TempPlayer(Index).Buffer.WriteBytes Buffer()
    
    'TempPlayer(Index).Buffer.DecompressBuffer
    
    If TempPlayer(Index).Buffer.Length >= 4 Then
        pLength = TempPlayer(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            HackingAttempt Index, "Hacking attempt."
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.Length - 4
        If pLength <= TempPlayer(Index).Buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            
            TempPlayer(Index).Buffer.ReadLong
            Data() = TempPlayer(Index).Buffer.ReadBytes(pLength + 1)
            
            'If EncryptPackets Then
            '    Encryption_XOR_DecryptByte Data(), PacketKeys(Player(Index).PacketInIndex)
            '    Player(Index).PacketInIndex = Player(Index).PacketInIndex + 1
            '    If Player(Index).PacketInIndex > PacketEncKeys - 1 Then Player(Index).PacketInIndex = 0
            'End If

            HandleData Index, Data()
        End If
        
        pLength = 0
        If TempPlayer(Index).Buffer.Length >= 4 Then
            pLength = TempPlayer(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                HackingAttempt Index, "Hacking attempt."
                Exit Sub
            End If
        End If
    Loop
    
    If GetPlayerAccess(Index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then
            HackingAttempt Index, "Data Flooding"
            Exit Sub
        End If
    
        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 25 Then
            HackingAttempt Index, "Packet Flooding"
            Exit Sub
        End If
    End If
            
    ' Check if elapsed time has passed
    'Player(Index).DataBytes = Player(Index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = GetTickCount + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If

    If TempPlayer(Index).Buffer.Length <= 1 Then TempPlayer(Index).Buffer.Flush
End Sub

Sub CloseSocket(ByVal Index As Long)

    If Index > 0 Then
        Call LeftGame(Index)
    
        Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        
        frmServer.Socket(Index).Close
            
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If
End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'Buffer.WriteLong SMapData
    Buffer.WriteLong MapNum
    Buffer.WriteString Trim$(Map(MapNum).Name)
    Buffer.WriteLong Map(MapNum).Revision
    Buffer.WriteLong Map(MapNum).Moral
    Buffer.WriteLong Map(MapNum).TileSet
    Buffer.WriteLong Map(MapNum).Up
    Buffer.WriteLong Map(MapNum).Down
    Buffer.WriteLong Map(MapNum).Left
    Buffer.WriteLong Map(MapNum).Right
    Buffer.WriteLong Map(MapNum).Music
    Buffer.WriteLong Map(MapNum).BootMap
    Buffer.WriteLong Map(MapNum).BootX
    Buffer.WriteLong Map(MapNum).BootY
    Buffer.WriteLong Map(MapNum).Shop
    Buffer.WriteLong Map(MapNum).MaxX
    Buffer.WriteLong Map(MapNum).MaxY
    
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            With Map(MapNum).Tile(x, y)
                Buffer.WriteLong .Ground
                Buffer.WriteLong .Mask
                Buffer.WriteLong .Anim
                Buffer.WriteLong .Mask2
                Buffer.WriteLong .Fringe
                Buffer.WriteLong .Fringe2
                Buffer.WriteLong .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
            End With
        Next
    Next
    
    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(MapNum).Npc(x)
    Next
    
    Buffer.CompressBuffer

    MapCache(MapNum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************

Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim n As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> Index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If
    Next
            
    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If
        
    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub SendChars(ByVal Index As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAllChars
    
    For i = 1 To MAX_CHARS
        Buffer.WriteString Trim$(Player(Index).Char(i).Name)
        Buffer.WriteString Trim$(Class(Player(Index).Char(i).Class).Name)
        Buffer.WriteLong Player(Index).Char(i).Level
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
Dim Buffer As clsBuffer

    If Index > MAX_PLAYERS Then Exit Function

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString GetPlayerName(Index)
    'Buffer.WriteLong GetPlayerLevel(Index)
    Buffer.WriteLong GetPlayerSprite(Index)
    Buffer.WriteLong GetPlayerMap(Index)
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    'Buffer.WriteLong GetPlayerClass(Index)
    
    PlayerData = Buffer.ToArray()
    
    Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal Index As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    SendDataTo Index, PlayerData(i)
                End If
            End If
        End If
    Next
    
    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong Index
    
    SendDataToMapBut Index, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal Index As Long)
Dim Packet As String
    
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)

End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(MapNum).Data()
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).Dur
        Buffer.WriteLong MapItem(MapNum, i).x
        Buffer.WriteLong MapItem(MapNum, i).y
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData
    
    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).Dur
        Buffer.WriteLong MapItem(MapNum, i).x
        Buffer.WriteLong MapItem(MapNum, i).y
    Next
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        'Packet = Packet & SEP_CHAR & mapnpc(mapnum).Npc(i).Num & SEP_CHAR & mapnpc(mapnum).Npc(i).x & SEP_CHAR & mapnpc(mapnum).Npc(i).y & SEP_CHAR & mapnpc(mapnum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
        Buffer.WriteLong MapNpc(MapNum).Npc(i).y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        'Packet = Packet & SEP_CHAR & MapNpc(MapNum).Npc(i).Num & SEP_CHAR & MapNpc(MapNum).Npc(i).x & SEP_CHAR & MapNpc(MapNum).Npc(i).y & SEP_CHAR & MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
        Buffer.WriteLong MapNpc(MapNum).Npc(i).y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
    Next
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendItems(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_ITEMS
        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If
    Next
End Sub

Sub SendNpcs(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_NPCS
        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If
    Next
End Sub

Sub SendInventory(ByVal Index As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, i)
        Buffer.WriteLong GetPlayerInvItemValue(Index, i)
        Buffer.WriteLong GetPlayerInvItemDur(Index, i)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong InvSlot
    Buffer.WriteLong GetPlayerInvItemNum(Index, InvSlot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, InvSlot)
    Buffer.WriteLong GetPlayerInvItemDur(Index, InvSlot)
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendWornEquipment(ByVal Index As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipmentSlot(Index, Armor)
    Buffer.WriteLong GetPlayerEquipmentSlot(Index, Weapon)
    Buffer.WriteLong GetPlayerEquipmentSlot(Index, Helmet)
    Buffer.WriteLong GetPlayerEquipmentSlot(Index, Shield)
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
        Case SP
            Buffer.WriteLong SPlayerSp
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.SP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.SP)
    End Select
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendStats(ByVal Index As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerStats
    Buffer.WriteLong GetPlayerStat(Index, Stats.Strength)
    Buffer.WriteLong GetPlayerStat(Index, Stats.Defense)
    Buffer.WriteLong GetPlayerStat(Index, Stats.Speed)
    Buffer.WriteLong GetPlayerStat(Index, Stats.Magic)
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendWelcome(ByVal Index As Long)
    ' Send them welcome
    Call PlayerMsg(Index, "Type /help for help on commands.  Use arrow keys to move, hold down shift to run, and use ctrl to attack.", Cyan)
    
    ' Send them MOTD
    If LenB(MOTD) > 0 Then
        Call PlayerMsg(Index, "MOTD: " & MOTD, BrightCyan)
    End If
    
    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Sub SendClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.SP)
        Buffer.WriteLong Class(i).Stat(Stats.Strength)
        Buffer.WriteLong Class(i).Stat(Stats.Defense)
        Buffer.WriteLong Class(i).Stat(Stats.Speed)
        Buffer.WriteLong Class(i).Stat(Stats.Magic)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Sub SendNewCharClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes
    
    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.SP)
        Buffer.WriteLong Class(i).Stat(Stats.Strength)
        Buffer.WriteLong Class(i).Stat(Stats.Defense)
        Buffer.WriteLong Class(i).Stat(Stats.Speed)
        Buffer.WriteLong Class(i).Stat(Stats.Magic)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendLeftGame(ByVal Index As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0

    SendDataToAllBut Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendPlayerXY(ByVal Index As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    
    SendDataTo Index, Buffer.ToArray()

    Set Buffer = Nothing
    
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteString Trim$(Item(ItemNum).Name)
    Buffer.WriteLong Item(ItemNum).Pic
    Buffer.WriteLong Item(ItemNum).Type
    Buffer.WriteLong Item(ItemNum).Data1
    Buffer.WriteLong Item(ItemNum).Data2
    Buffer.WriteLong Item(ItemNum).Data3

    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteString Trim$(Item(ItemNum).Name)
    Buffer.WriteLong Item(ItemNum).Pic
    Buffer.WriteLong Item(ItemNum).Type
    Buffer.WriteLong Item(ItemNum).Data1
    Buffer.WriteLong Item(ItemNum).Data2
    Buffer.WriteLong Item(ItemNum).Data3
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SEditItem
    Buffer.WriteLong ItemNum
    Buffer.WriteString Trim$(Item(ItemNum).Name)
    Buffer.WriteLong Item(ItemNum).Pic
    Buffer.WriteLong Item(ItemNum).Type
    Buffer.WriteLong Item(ItemNum).Data1
    Buffer.WriteLong Item(ItemNum).Data2
    Buffer.WriteLong Item(ItemNum).Data3

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteLong Npc(NpcNum).Sprite

    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteLong Npc(NpcNum).Sprite

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SEditNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteString Trim$(Npc(NpcNum).AttackSay)
    Buffer.WriteLong Npc(NpcNum).Sprite
    Buffer.WriteLong Npc(NpcNum).SpawnSecs
    Buffer.WriteLong Npc(NpcNum).Behavior
    Buffer.WriteLong Npc(NpcNum).Range
    Buffer.WriteLong Npc(NpcNum).DropChance
    Buffer.WriteLong Npc(NpcNum).DropItem
    Buffer.WriteLong Npc(NpcNum).DropItemValue
    Buffer.WriteLong Npc(NpcNum).Stat(Stats.Strength)
    Buffer.WriteLong Npc(NpcNum).Stat(Stats.Defense)
    Buffer.WriteLong Npc(NpcNum).Stat(Stats.Speed)
    Buffer.WriteLong Npc(NpcNum).Stat(Stats.Magic)
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendShops(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If
    Next
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SEditShop
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    Buffer.WriteString Trim$(Shop(ShopNum).JoinSay)
    Buffer.WriteString Trim$(Shop(ShopNum).LeaveSay)
    Buffer.WriteLong Shop(ShopNum).FixesItems

    For i = 1 To MAX_TRADES
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GiveItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GiveValue
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GetItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GetValue
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Sub SendSpells(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If
    Next
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteString Trim$(Spell(SpellNum).Name)
    Buffer.WriteLong Spell(SpellNum).MPReq
    Buffer.WriteLong Spell(SpellNum).Pic

    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteString Trim$(Spell(SpellNum).Name)
    Buffer.WriteLong Spell(SpellNum).MPReq
    Buffer.WriteLong Spell(SpellNum).Pic

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SEditSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteString Trim$(Spell(SpellNum).Name)
    Buffer.WriteLong Spell(SpellNum).Pic
    Buffer.WriteLong Spell(SpellNum).MPReq
    Buffer.WriteLong Spell(SpellNum).ClassReq
    Buffer.WriteLong Spell(SpellNum).LevelReq
    Buffer.WriteLong Spell(SpellNum).Type
    Buffer.WriteLong Spell(SpellNum).Data1
    Buffer.WriteLong Spell(SpellNum).Data2
    Buffer.WriteLong Spell(SpellNum).Data3
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long
Dim x As Long
Dim y As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong STrade
    Buffer.WriteLong ShopNum
    Buffer.WriteLong Shop(ShopNum).FixesItems

    For i = 1 To MAX_TRADES
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GiveItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GiveValue
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GetItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GetValue
        
        ' Item #
        x = Shop(ShopNum).TradeItem(i).GetItem
        
        If x > 0 And x <= MAX_ITEMS Then
        
            If Item(x).Type = ITEM_TYPE_SPELL Then
                ' Spell class requirement
                y = Spell(Item(x).Data1).ClassReq
                
                If y = 0 Then
                    Call PlayerMsg(Index, Trim$(Item(x).Name) & " can be used by all classes.", Yellow)
                Else
                    Call PlayerMsg(Index, Trim$(Item(x).Name) & " can only be used by a " & GetClassName(y - 1) & ".", Yellow)
                End If
            End If
            
        End If
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
Dim Packet As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) > 0 Then
            Buffer.WriteLong i
            Buffer.WriteLong GetPlayerSpell(Index, i)
        End If
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

