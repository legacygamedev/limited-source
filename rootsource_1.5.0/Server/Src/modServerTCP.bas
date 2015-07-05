Attribute VB_Name = "modServerTCP"
Option Explicit

' ********************************************
' **               rootSource               **
' ********************************************


Public Sub UpdateCaption()
    frmServer.Caption = "Mirage Source Server <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalPlayersOnline & ")"
End Sub

Public Sub CreateFullMapCache()
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

    For i = 1 To TotalPlayersOnline
        If LCase$(Trim$(Player(PlayersOnline(i)).Login)) = LCase$(Login) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next
    
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
Dim i As Long
Dim n As Long

    For i = 1 To TotalPlayersOnline
        If Trim$(GetPlayerIP(PlayersOnline(i))) = IP Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
            
        End If
    Next
End Function

Private Function IsBanned(ByVal IP As String) As Boolean
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

Public Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteInteger (UBound(Data) - LBound(Data)) + 1 ' Writes the length of the packet
        Buffer.WriteBytes Data()            ' Writes the data to the packet
        frmServer.Socket(Index).SendData Buffer.ToArray()
    End If
End Sub

Public Sub SendDataToAll(ByRef Data() As Byte)
Dim i As Long

    For i = 1 To TotalPlayersOnline
        Call SendDataTo(PlayersOnline(i), Data)
    Next
End Sub

Public Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To TotalPlayersOnline
        If PlayersOnline(i) <> Index Then
            Call SendDataTo(PlayersOnline(i), Data)
        End If
    Next
End Sub

Public Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To TotalPlayersOnline
        'If GetPlayerMap(PlayersOnline(i)) = MapNum Then
            Call SendDataTo(PlayersOnline(i), Data)
        'End If
    Next
End Sub

Public Sub SendDataToMaps(ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Long
Dim n As Long
Dim tMap(1 To 9) As Long

    tMap(5) = MapNum
    tMap(2) = Map(tMap(5)).Up
    tMap(8) = Map(tMap(5)).Down
    tMap(4) = Map(tMap(5)).Left
    tMap(6) = Map(tMap(5)).Right
    
    If tMap(4) <> 0 Then
        tMap(7) = Map(tMap(4)).Down
        tMap(1) = Map(tMap(4)).Up
    End If
    
    If tMap(6) <> 0 Then
        tMap(9) = Map(tMap(6)).Down
        tMap(3) = Map(tMap(6)).Up
    End If
    
    For i = 1 To TotalPlayersOnline
        For n = 1 To 9
            If GetPlayerMap(PlayersOnline(i)) = tMap(n) Then
                Call SendDataTo(PlayersOnline(i), Data)
            End If
        Next
    Next
End Sub

Public Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To TotalPlayersOnline
        'If GetPlayerMap(PlayersOnline(i)) = MapNum Then
            If PlayersOnline(i) <> Index Then
                Call SendDataTo(PlayersOnline(i), Data)
            End If
        'End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 5
    Buffer.WriteInteger SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    
    Call SendDataToAll(Buffer.ToArray)
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 5
    Buffer.WriteInteger SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    
    For i = 1 To TotalPlayersOnline
        If GetPlayerAccess(PlayersOnline(i)) > 0 Then
            Call SendDataTo(PlayersOnline(i), Buffer.ToArray)
        End If
    Next
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 5
    Buffer.WriteInteger SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 5
    Buffer.WriteInteger SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    
    Call SendDataToMaps(MapNum, Buffer.ToArray)
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.PreAllocate Len(Msg) + 4
    Buffer.WriteInteger SAlertMsg
    Buffer.WriteString Msg
    
    Call SendDataTo(Index, Buffer.ToArray)
    DoEvents
    Call CloseSocket(Index)
End Sub

Public Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Public Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
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

Public Sub SocketConnected(ByVal Index As Long)
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
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    ' Get the data as an array
    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    
    ' Write the bytes to the byte array
    TempPlayer(Index).Buffer.WriteBytes Buffer()
    
    ' Check if we have enough in the buffer
    If TempPlayer(Index).Buffer.Length >= 2 Then
        pLength = TempPlayer(Index).Buffer.ReadInteger(False)
    
        ' If the plength is less than 0 then we know there was something odd
        ' hacking attempt is usually what happened
        If pLength < 0 Then
            HackingAttempt Index, "Hacking attempt."
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.Length - 2
        If pLength <= TempPlayer(Index).Buffer.Length - 2 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            
            ' Remove the "size" off the packet now that we have the full packet
            TempPlayer(Index).Buffer.ReadInteger
            ' Handle the packet data
            HandleData Index, TempPlayer(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(Index).Buffer.Length >= 2 Then
            pLength = TempPlayer(Index).Buffer.ReadInteger(False)
        
            If pLength < 0 Then
                HackingAttempt Index, "Hacking attempt."
                Exit Sub
            End If
        End If
    Loop
            
    ' Trim down the packet
    TempPlayer(Index).Buffer.Trim
    
    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    
    If GetTickCount >= TempPlayer(Index).DataTimer + 1000 Then
        TempPlayer(Index).DataTimer = GetTickCount
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
        Exit Sub
    End If
    
    If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 2000 Then
            Call HackingAttempt(Index, "Data Flooding")
            Exit Sub
        End If
        
        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 55 Then
            Call HackingAttempt(Index, "Packet Flooding")
            Exit Sub
        End If
    End If
    
End Sub

Public Sub CloseSocket(ByVal Index As Long)

    If Index > 0 Then
        Call LeftGame(Index)
    
        Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        
        frmServer.Socket(Index).Close
            
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If
End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
Dim MapSize As Long
Dim MapData() As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    MapSize = LenB(Map(MapNum))
    ReDim MapData(MapSize - 1)
    CopyMemory MapData(0), ByVal VarPtr(Map(MapNum)), MapSize
    
    Buffer.PreAllocate MapSize + 6
    Buffer.WriteInteger SMapData
    Buffer.WriteLong MapNum
    Buffer.WriteBytes MapData
    
    MapCache(MapNum).Cache = Buffer.ToArray()
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************

Public Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim n As Long
Dim i As Long

    For i = 1 To TotalPlayersOnline
        If PlayersOnline(i) <> Index Then
            s = s & GetPlayerName(PlayersOnline(i)) & ", "
            n = n + 1
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

Public Sub SendChars(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteInteger SAllChars
    For i = 1 To MAX_CHARS
        Buffer.WriteLong Player(Index).Char(i).Sprite
        Buffer.WriteString Trim$(Player(Index).Char(i).Name)
        Buffer.WriteString Trim$(Class(Player(Index).Char(i).Class).Name)
        Buffer.WriteByte Player(Index).Char(i).Level
    Next
    
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Public Sub SendMaxes(ByVal Index As Long)
Dim Buffer As New clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 14
    Buffer.WriteInteger SSendMaxes
    Buffer.WriteInteger MAX_PLAYERS
    Buffer.WriteInteger MAX_ITEMS
    Buffer.WriteInteger MAX_NPCS
    Buffer.WriteInteger MAX_SHOPS
    Buffer.WriteInteger MAX_SPELLS
    Buffer.WriteInteger MAX_MAPS
    
    Call SendDataTo(Index, Buffer.ToArray)
    
End Sub

Public Sub SendJoinMap(ByVal Index As Long)
Dim i As Long
    ' Send all players on current map to index
    For i = 1 To TotalPlayersOnline
        If PlayersOnline(i) <> Index Then
            'If GetPlayerMap(PlayersOnline(i)) = GetPlayerMap(Index) Then
                Call SendDataTo(Index, PlayerData(PlayersOnline(i)))
            'End If
        End If
    Next
    
    ' Send index's player data to everyone on the map including himself
    Call SendDataToAll(PlayerData(Index))
End Sub

Public Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger SLeft
    Buffer.WriteLong Index

    Call SendDataToMapBut(Index, MapNum, Buffer.ToArray())
End Sub

Public Sub SendPlayerData(ByVal Index As Long)
    ' Send index's player data to everyone including himself on th emap
    Call SendDataToMap(GetPlayerMap(Index), PlayerData(Index))
End Sub

Public Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Call SendDataTo(Index, MapCache(MapNum).Cache)
End Sub

Public Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_MAP_ITEMS * 9) + 2 + 4
    Buffer.WriteInteger SMapItemData
    Buffer.WriteLong MapNum
    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteByte MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteInteger MapItem(MapNum, i).Dur
        Buffer.WriteByte MapItem(MapNum, i).X
        Buffer.WriteByte MapItem(MapNum, i).y
    Next
    
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Public Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_MAP_ITEMS * 9) + 2 + 4
    Buffer.WriteInteger SMapItemData
    Buffer.WriteLong MapNum
    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteByte MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteInteger MapItem(MapNum, i).Dur
        Buffer.WriteByte MapItem(MapNum, i).X
        Buffer.WriteByte MapItem(MapNum, i).y
    Next
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_MAP_NPCS * 6) + 6
    Buffer.WriteInteger SMapNpcData
    Buffer.WriteLong MapNum
    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteInteger MapNpc(MapNum, i).Num
        Buffer.WriteByte MapNpc(MapNum, i).X
        Buffer.WriteByte MapNpc(MapNum, i).y
        Buffer.WriteInteger MapNpc(MapNum, i).Dir
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_MAP_NPCS * 6) + 6
    Buffer.WriteInteger SMapNpcData
    Buffer.WriteLong MapNum
    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteInteger MapNpc(MapNum, i).Num
        Buffer.WriteByte MapNpc(MapNum, i).X
        Buffer.WriteByte MapNpc(MapNum, i).y
        Buffer.WriteInteger MapNpc(MapNum, i).Dir
    Next
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendMapRevs(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_MAPS * 4) + 2
    Buffer.WriteInteger SMapRevs
    For i = 1 To MAX_MAPS
        Buffer.WriteLong Map(i).Revision
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendItems(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_ITEMS
        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If
    Next
End Sub

Public Sub SendNpcs(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_NPCS
        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If
    Next
    
    For i = 1 To MAX_MAPS
        Call SendMapNpcsTo(Index, i)
    Next
End Sub

Public Sub SendInventory(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_INV * 12) + 2
    Buffer.WriteInteger SPlayerInv
    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, i)
        Buffer.WriteLong GetPlayerInvItemValue(Index, i)
        Buffer.WriteLong GetPlayerInvItemDur(Index, i)
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 18
    Buffer.WriteInteger SPlayerInvUpdate
    Buffer.WriteLong InvSlot
    Buffer.WriteLong GetPlayerInvItemNum(Index, InvSlot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, InvSlot)
    Buffer.WriteLong GetPlayerInvItemDur(Index, InvSlot)
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendWornEquipment(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (Equipment.Equipment_Count - 1) + 2
    Buffer.WriteInteger SPlayerWornEq
    For i = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteByte GetPlayerEquipmentSlot(Index, i)
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Select Case Vital
        Case HP
            Buffer.WriteInteger SPlayerHp
        Case MP
            Buffer.WriteInteger SPlayerMp
        Case SP
            Buffer.WriteInteger SPlayerSp
    End Select
    
    Buffer.WriteLong GetPlayerMaxVital(Index, Vital)
    Buffer.WriteLong GetPlayerVital(Index, Vital)
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendStats(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate ((Stats.Stat_Count - 1) * 4) + 2
    Buffer.WriteInteger SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(Index, i)
    Next
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendWelcome(ByVal Index As Long)
    ' Send them welcome
    Call PlayerMsg(Index, "Type /help for help on commands.  Use arrow keys to move, hold down shift to run, and use ctrl to attack.", Cyan)
    
    ' Send them MOTD
    If LenB(MOTD) > 0 Then
        Call PlayerMsg(Index, "MOTD: " & MOTD, BrightCyan)
    End If
    
    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Public Sub SendClasses(ByVal Index As Long)
Dim i As Long
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteInteger SClassesData
    Buffer.WriteByte Max_Classes
    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong Class(i).Sprite
        For n = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong GetClassMaxVital(i, n)
        Next
        For n = 1 To Stats.Stat_Count - 1
            Buffer.WriteByte Class(i).Stat(n)
        Next
    Next
    
    Call SendDataTo(Index, Buffer.ToArray)
End Sub

Public Sub SendNewCharClasses(ByVal Index As Long)
Dim i As Long
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteInteger SNewCharClasses
    Buffer.WriteByte Max_Classes
    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong Class(i).Sprite
        For n = 1 To Vitals.Vital_Count - 1
            Buffer.WriteLong GetClassMaxVital(i, n)
        Next
        For n = 1 To Stats.Stat_Count - 1
            Buffer.WriteByte Class(i).Stat(n)
        Next
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendLeftGame(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger SLeft
    Buffer.WriteLong Index
    
    Call SendDataToAllBut(Index, Buffer.ToArray())
End Sub

Public Sub SendPlayerXY(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 10
    Buffer.WriteInteger SPlayerXY
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Item(ItemNum)) + 2
    Buffer.WriteInteger SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteString Trim$(Item(ItemNum).Name)
    Buffer.WriteInteger Item(ItemNum).Pic
    Buffer.WriteByte Item(ItemNum).Type
    Buffer.WriteInteger Item(ItemNum).Data1
    Buffer.WriteInteger Item(ItemNum).Data2
    Buffer.WriteInteger Item(ItemNum).Data3
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Item(ItemNum)) + 2
    Buffer.WriteInteger SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteString Trim$(Item(ItemNum).Name)
    Buffer.WriteInteger Item(ItemNum).Pic
    Buffer.WriteByte Item(ItemNum).Type
    Buffer.WriteInteger Item(ItemNum).Data1
    Buffer.WriteInteger Item(ItemNum).Data2
    Buffer.WriteInteger Item(ItemNum).Data3
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Buffer As clsBuffer
Dim ItemData() As Byte
Dim ItemSize As Long

    Set Buffer = New clsBuffer
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)

    Buffer.PreAllocate ItemSize + 6
    Buffer.WriteInteger SEditItem
    Buffer.WriteLong ItemNum
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteBytes ItemData
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Npc(NpcNum).Name)) + 8
    Buffer.WriteInteger SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteInteger Npc(NpcNum).Sprite
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Npc(NpcNum).Name)) + 8
    Buffer.WriteInteger SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteInteger Npc(NpcNum).Sprite
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Buffer As clsBuffer
Dim NpcData() As Byte
Dim NpcSize As Long

    Set Buffer = New clsBuffer
    
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize)

    Buffer.PreAllocate NpcSize + 6
    Buffer.WriteInteger SEditNpc
    Buffer.WriteLong NpcNum
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    Buffer.WriteBytes NpcData
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendShops(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If
    Next
End Sub

Public Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Shop(ShopNum).Name)) + 8
    Buffer.WriteInteger SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Shop(ShopNum).Name)) + 8
    Buffer.WriteInteger SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Buffer As clsBuffer
Dim ShopData() As Byte
Dim ShopSize As Long

    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize)

    Buffer.PreAllocate ShopSize + 6
    Buffer.WriteInteger SEditShop
    Buffer.WriteLong ShopNum
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    Buffer.WriteBytes ShopData
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendSpells(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If
    Next
End Sub

Public Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Spell(SpellNum).Name)) + 12
    Buffer.WriteInteger SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteString Trim$(Spell(SpellNum).Name)
    Buffer.WriteInteger Spell(SpellNum).MPReq
    Buffer.WriteInteger Spell(SpellNum).Pic
    
    Call SendDataToAll(Buffer.ToArray())
End Sub

Public Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Spell(SpellNum).Name)) + 12
    Buffer.WriteInteger SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteString Trim$(Spell(SpellNum).Name)
    Buffer.WriteInteger Spell(SpellNum).MPReq
    Buffer.WriteInteger Spell(SpellNum).Pic
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Buffer As clsBuffer
Dim SpellData() As Byte
Dim SpellSize As Long

    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize)

    Buffer.PreAllocate SpellSize + 6
    Buffer.WriteInteger SEditSpell
    Buffer.WriteLong SpellNum
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    Buffer.WriteBytes SpellData
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim i As Long
Dim X As Long
Dim y As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_TRADES * 16) + 7
    Buffer.WriteInteger STrade
    Buffer.WriteLong ShopNum
    Buffer.WriteByte Shop(ShopNum).FixesItems
    
    For i = 1 To MAX_TRADES
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GiveItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GiveValue
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GetItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).GetValue
        
        ' Item #
        X = Shop(ShopNum).TradeItem(i).GetItem
        
        If X > 0 And X <= MAX_ITEMS Then
        
            If Item(X).Type = ITEM_TYPE_SPELL Then
                ' Spell class requirement
                y = Spell(Item(X).Data1).ClassReq
                
                If y = 0 Then
                    Call PlayerMsg(Index, Trim$(Item(X).Name) & " can be used by all classes.", Yellow)
                Else
                    Call PlayerMsg(Index, Trim$(Item(X).Name) & " can only be used by a " & GetClassName(y - 1) & ".", Yellow)
                End If
            End If
            
        End If
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendPlayerSpells(ByVal Index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (MAX_PLAYER_SPELLS * 8) + 2
    Buffer.WriteInteger SSpells
    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong i
        Buffer.WriteLong GetPlayerSpell(Index, i)
    Next
    
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

Public Sub SendDoorData(ByVal Index As Long)
Dim X As Long
Dim y As Long
Dim Buffer As clsBuffer
    
    For X = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            If TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = YES Then
                Set Buffer = New clsBuffer
    
                Buffer.PreAllocate 10
                Buffer.WriteInteger SDoor
                Buffer.WriteLong X
                Buffer.WriteLong y
                
                Call SendDataTo(Index, Buffer.ToArray())
            End If
        Next
    Next

End Sub

Public Function PlayerData(ByVal Index As Long) As Byte()
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteInteger SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerSprite(Index)
    Buffer.WriteLong GetPlayerMap(Index)
    Buffer.WriteLong GetPlayerX(Index)
    Buffer.WriteLong GetPlayerY(Index)
    Buffer.WriteString GetPlayerGuild(Index)
    Buffer.WriteLong GetPlayerGAccess(Index)
    Buffer.WriteLong GetPlayerDir(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    PlayerData = Buffer.ToArray()
End Function
