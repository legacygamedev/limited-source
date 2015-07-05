Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateTitle()
    frmServer.caption = GAME_NAME & " (" & frmServer.Socket(0).LocalIP & ":" & CStr(GameServer.LocalPort) & ") - Eclipse Stable Server"
End Sub

Sub UpdateTOP()
    frmServer.TPO.caption = "Total Players Online: " & TotalOnlinePlayers
End Sub

Sub SendPlayerNewXY(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERNEWXY" & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR)
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    On Error Resume Next

    If GameServer.Sockets.Item(Index).Socket Is Nothing Then
        IsConnected = False
    Else
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    If IsConnected(Index) Then
        If Player(Index).InGame Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    If IsConnected(Index) Then
        If Trim$(Player(Index).Login) <> vbNullString Then
            IsLoggedIn = True
        End If
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsConnected(I) Then
            If LCase$(Trim$(Player(I).Login)) = LCase$(Trim$(Login)) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If
    Next I
End Function

Function IsBanned(ByVal IPAddr As String) As Boolean
    Dim FileName As String
    Dim FileIP As String
    Dim FileID As Long

    FileName = App.Path & "\BanList.txt"

    FileID = FreeFile

    ' Check if file exists
    If Not FileExists("BanList.txt") Then
        Open FileName For Output As #FileID
        Close #FileID
    End If

    Open FileName For Input As #FileID
        Do While Not EOF(FileID)
            Line Input #FileID, FileIP
    
            If FileIP = IPAddr Then
                IsBanned = True
                Exit Do
            End If
        Loop
    Close #FileID
End Function

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
    Dim dbytes() As Byte

    dbytes = StrConv(Data, vbFromUnicode)
    If IsConnected(Index) Then
        GameServer.Sockets.Item(Index).WriteBytes dbytes
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendDataTo(I, Data)
        End If
    Next I
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If I <> Index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                If I <> Index Then
                    Call SendDataTo(I, Data)
                End If
            End If
        End If
    Next I
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToAll("GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataTo(Index, "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim packet As String
    Dim I As Long

    packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerAccess(I) > 0 Then
                Call SendDataTo(I, packet)
            End If
        End If
    Next I
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToMap(MapNum, "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Call SendDataTo(Index, "ALERTMSG" & SEP_CHAR & Msg & END_CHAR)
    Call CloseSocket(Index)
End Sub

Sub PlainMsg(ByVal Index As Long, ByVal Msg As String, ByVal num As Long)
    Call SendDataTo(Index, "PLAINMSG" & SEP_CHAR & Msg & SEP_CHAR & num & END_CHAR)
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Sub
    End If

    If IsPlaying(Index) Then
        Call AdminMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", BRIGHTRED)
        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub AcceptConnection(Socket As JBSOCKETSERVERLib.ISocket)
    Dim Index As Long

    Index = FindOpenPlayerSlot

    If Index > 0 Then
        Socket.UserData = Index
        Set GameServer.Sockets.Item(CStr(Index)).Socket = Socket
        Call SocketConnected(Index)
        Socket.RequestRead
    Else
        Socket.Close
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Sub
    End If

    If Not IsBanned(GetPlayerIP(Index)) Then
        Call TextAdd(frmServer.txtText(0), "Received connection from " & GetPlayerIP(Index) & ".", True)
    Else
        Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
    End If
End Sub

Sub IncomingData(Socket As JBSOCKETSERVERLib.ISocket, Data As JBSOCKETSERVERLib.IData)
    Dim dbytes() As Byte
    Dim Buffer As String
    Dim Start As Long
    Dim packet As String
    Dim Index As Long

    dbytes = Data.Read

    Socket.RequestRead

    Buffer = StrConv(dbytes(), vbUnicode)

    Index = CLng(Socket.UserData)

    Player(Index).Buffer = Player(Index).Buffer & Buffer

    ' Check if elapsed time has passed
    Player(Index).DataBytes = Player(Index).DataBytes + Len(Buffer)
    If GetTickCount >= Player(Index).DataTimer + 1000 Then
        If Player(Index).CharNum <> 0 Then
            Player(Index).DataTimer = GetTickCount
            Player(Index).DataBytes = 0
            Player(Index).DataPackets = 0
        End If
    End If

    ' Check for data flooding
    If Player(Index).DataBytes > 1000 Then
        If GetPlayerAccess(Index) <= 0 Then
            Call HackingAttempt(Index, "Data Flooding")
            Exit Sub
        End If
    End If

    ' Check for packet flooding
    If Player(Index).DataPackets > 50 Then
        If GetPlayerAccess(Index) = 0 Then
            Call HackingAttempt(Index, "Packet Flooding")
            Exit Sub
        End If
    End If

    Start = InStr(Player(Index).Buffer, END_CHAR)
    
    Do While Start > 0
        packet = Mid$(Player(Index).Buffer, 1, Start - 1)
        Player(Index).Buffer = Mid$(Player(Index).Buffer, Start + 1, Len(Player(Index).Buffer))
        Player(Index).DataPackets = Player(Index).DataPackets + 1
        Start = InStr(Player(Index).Buffer, END_CHAR)
        If LenB(packet) <> 0 Then
            Call HandleData(Index, packet)
        End If
    Loop
End Sub

Sub CloseSocket(ByVal Index As Long)
    If Index > 0 Then
        Call LeftGame(Index)
       
        Call TextAdd(frmServer.txtText(0), "Connection from " & GetPlayerIP(Index) & " has been terminated.", True)

        Call GameServer.Sockets.Item(Index).ShutDown(ShutdownBoth)

        Set GameServer.Sockets.Item(Index).Socket = Nothing
       
        Call UpdateTOP
        Call ClearPlayer(Index)
    End If
End Sub

Sub SendWhosOnline(ByVal Index As Long)
    Dim PlayerNames As String
    Dim PlayerCount As Long
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If I <> Index Then
                PlayerNames = PlayerNames & GetPlayerName(I) & ", "
                PlayerCount = PlayerCount + 1
            End If
        End If
    Next I

    If PlayerCount = 0 Then
        PlayerNames = "There are no other players online."
    Else
        PlayerNames = Mid$(PlayerNames, 1, Len(PlayerNames) - 2)
        PlayerNames = "There are " & PlayerCount & " other players online: " & PlayerNames & "."
    End If

    Call PlayerMsg(Index, PlayerNames, WhoColor)
End Sub

Sub SendOnlineList()
    Dim packet As String
    Dim PlayerCount As Long
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            packet = packet & SEP_CHAR & GetPlayerName(I) & SEP_CHAR
            PlayerCount = PlayerCount + 1
        End If
    Next I

    Call SendDataToAll("ONLINELIST" & SEP_CHAR & PlayerCount & packet & END_CHAR)
End Sub

Sub SendChars(ByVal Index As Long)
    Dim packet As String
    Dim I As Long

    packet = "ALLCHARS" & SEP_CHAR
    For I = 1 To MAX_CHARS
        packet = packet & Trim$(Player(Index).Char(I).Name) & SEP_CHAR & Trim$(ClassData(Player(Index).Char(I).Class).Name) & SEP_CHAR & Player(Index).Char(I).LEVEL & SEP_CHAR
    Next I
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendJoinMap(ByVal Index As Long)
    Dim packet As String
    Dim I As Long
    Dim j As Long

    packet = vbNullString

    ' Send all players on current map to index
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            'If I <> index Then
                If GetPlayerMap(I) = GetPlayerMap(Index) Then
                    packet = "PLAYERDATA" & SEP_CHAR
                    packet = packet & I & SEP_CHAR
                    packet = packet & GetPlayerName(I) & SEP_CHAR
                    packet = packet & GetPlayerSprite(I) & SEP_CHAR
                    packet = packet & GetPlayerMap(I) & SEP_CHAR
                    packet = packet & GetPlayerX(I) & SEP_CHAR
                    packet = packet & GetPlayerY(I) & SEP_CHAR
                    packet = packet & GetPlayerDir(I) & SEP_CHAR
                    packet = packet & GetPlayerAccess(I) & SEP_CHAR
                    packet = packet & GetPlayerPK(I) & SEP_CHAR
                    packet = packet & GetPlayerGuild(I) & SEP_CHAR
                    packet = packet & GetPlayerGuildAccess(I) & SEP_CHAR
                    packet = packet & GetPlayerClass(I) & SEP_CHAR
                    packet = packet & GetPlayerHead(I) & SEP_CHAR
                    packet = packet & GetPlayerBody(I) & SEP_CHAR
                    packet = packet & GetPlayerleg(I) & SEP_CHAR
                    packet = packet & GetPlayerPaperdoll(I) & SEP_CHAR
                    packet = packet & GetPlayerLevel(I) & SEP_CHAR
                    packet = packet & END_CHAR
                    Call SendDataTo(Index, packet)
                End If
            'End If
        End If
    Next I

    packet = packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String

    packet = "leave" & SEP_CHAR & Index & END_CHAR
    Call SendDataToMapBut(Index, MapNum, packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String
    Dim j As Long

    ' Send index's player data to everyone including himself on th emap
    packet = "PLAYERDATA" & SEP_CHAR
    packet = packet & Index & SEP_CHAR
    packet = packet & GetPlayerName(Index) & SEP_CHAR
    packet = packet & GetPlayerSprite(Index) & SEP_CHAR
    packet = packet & GetPlayerMap(Index) & SEP_CHAR
    packet = packet & GetPlayerX(Index) & SEP_CHAR
    packet = packet & GetPlayerY(Index) & SEP_CHAR
    packet = packet & GetPlayerDir(Index) & SEP_CHAR
    packet = packet & GetPlayerAccess(Index) & SEP_CHAR
    packet = packet & GetPlayerPK(Index) & SEP_CHAR
    packet = packet & GetPlayerGuild(Index) & SEP_CHAR
    packet = packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    packet = packet & GetPlayerClass(Index) & SEP_CHAR
    packet = packet & GetPlayerHead(Index) & SEP_CHAR
    packet = packet & GetPlayerBody(Index) & SEP_CHAR
    packet = packet & GetPlayerleg(Index) & SEP_CHAR
    packet = packet & GetPlayerPaperdoll(Index) & SEP_CHAR
    packet = packet & GetPlayerLevel(Index) & SEP_CHAR

    packet = packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Public Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    If LenB(MapCache(MapNum)) = 0 Then
        Call MapCache_Create(MapNum)
    End If

    Call SendDataTo(Index, MapCache(MapNum))
End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim X As Long
    Dim Y As Long

    MapData = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR & Map(MapNum).Weather & SEP_CHAR

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map(MapNum).Tile(X, Y)
                MapData = MapData & (.Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .Light & SEP_CHAR)
                MapData = MapData & (.GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR)
            End With
        Next X
    Next Y

    For X = 1 To MAX_MAP_NPCS
        MapData = MapData & (Map(MapNum).NPC(X) & SEP_CHAR & Map(MapNum).SpawnX(X) & SEP_CHAR & Map(MapNum).SpawnY(X) & SEP_CHAR)
    Next X

    MapData = MapData & END_CHAR

    MapCache(MapNum) = MapData
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim I As Long

    packet = "MAPITEMDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_ITEMS
        If MapNum > 0 Then
            packet = packet & (MapItem(MapNum, I).num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).X & SEP_CHAR & MapItem(MapNum, I).Y & SEP_CHAR)
        End If
    Next I
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
    Dim packet As String
    Dim I As Long

    packet = "MAPITEMDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_ITEMS
        packet = packet & (MapItem(MapNum, I).num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).X & SEP_CHAR & MapItem(MapNum, I).Y & SEP_CHAR)
    Next I
    packet = packet & END_CHAR

    Call SendDataToMap(MapNum, packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim I As Long

    packet = "MAPNPCDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_NPCS
        If MapNum > 0 Then
            packet = packet & (MapNPC(MapNum, I).num & SEP_CHAR & MapNPC(MapNum, I).X & SEP_CHAR & MapNPC(MapNum, I).Y & SEP_CHAR & MapNPC(MapNum, I).Dir & SEP_CHAR)
        End If
    Next I
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
    Dim packet As String
    Dim I As Long

    packet = "MAPNPCDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_NPCS
        packet = packet & (MapNPC(MapNum, I).num & SEP_CHAR & MapNPC(MapNum, I).X & SEP_CHAR & MapNPC(MapNum, I).Y & SEP_CHAR & MapNPC(MapNum, I).Dir & SEP_CHAR)
    Next I
    packet = packet & END_CHAR

    Call SendDataToMap(MapNum, packet)
End Sub

Sub SendItems(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_ITEMS
        If Trim$(Item(I).Name) <> vbNullString Then
            Call SendUpdateItemTo(Index, I)
        End If
    Next I
End Sub

Sub SendElements(ByVal Index As Long)
    Dim I As Long

    For I = 0 To MAX_ELEMENTS
        If Trim$(Element(I).Name) <> vbNullString Then
            Call SendUpdateElementTo(Index, I)
        End If
    Next I
End Sub
Sub SendEmoticons(ByVal Index As Long)
    Dim I As Long

    For I = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(I).Command) <> vbNullString Then
            Call SendUpdateEmoticonTo(Index, I)
        End If
    Next I
End Sub

Sub SendArrows(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_ARROWS
        If Trim$(Arrows(I).Name) <> vbNullString Then
            Call SendUpdateArrowTo(Index, I)
        End If
    Next I
End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_NPCS
        If Trim$(NPC(I).Name) <> vbNullString Then
            Call SendUpdateNpcTo(Index, I)
        End If
    Next I
End Sub
Sub SendBank(ByVal Index As Long)
    Dim packet As String
    Dim I As Integer

    packet = "PLAYERBANK" & SEP_CHAR
    For I = 1 To MAX_BANK
        packet = packet & (GetPlayerBankItemNum(Index, I) & SEP_CHAR & GetPlayerBankItemValue(Index, I) & SEP_CHAR & GetPlayerBankItemDur(Index, I) & SEP_CHAR)
    Next I
    packet = packet & END_CHAR
    
    ' Dump the packet as-is into a file.
    Call PacketDump(packet, "Logs\PacketDump.txt")

    Call SendDataTo(Index, packet)
End Sub

Sub SendBankUpdate(ByVal Index As Long, ByVal BankSlot As Long)
    Call SendDataTo(Index, "PLAYERBANKUPDATE" & SEP_CHAR & BankSlot & SEP_CHAR & GetPlayerBankItemNum(Index, BankSlot) & SEP_CHAR & GetPlayerBankItemValue(Index, BankSlot) & SEP_CHAR & GetPlayerBankItemDur(Index, BankSlot) & END_CHAR)
End Sub
Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim I As Long

    packet = "PLAYERINV" & SEP_CHAR & Index & SEP_CHAR
    For I = 1 To MAX_INV
        packet = packet & (GetPlayerInvItemNum(Index, I) & SEP_CHAR & GetPlayerInvItemValue(Index, I) & SEP_CHAR & GetPlayerInvItemDur(Index, I) & SEP_CHAR)
    Next I
    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Call SendDataToMap(GetPlayerMap(Index), "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & END_CHAR)
End Sub

Sub SendIndexInventoryFromMap(ByVal Index As Long)
    Dim packet As String
    Dim n As Long
    Dim I As Long
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = GetPlayerMap(Index) Then
                packet = "PLAYERINV" & SEP_CHAR & I & SEP_CHAR
                For n = 1 To MAX_INV
                    packet = packet & (GetPlayerInvItemNum(I, n) & SEP_CHAR & GetPlayerInvItemValue(I, n) & SEP_CHAR & GetPlayerInvItemDur(I, n) & SEP_CHAR)
                Next n
                packet = packet & END_CHAR

                Call SendDataTo(Index, packet)
            End If
        End If
    Next I
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim packet As String
    If IsPlaying(Index) Then
        packet = "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).ArmorSlot & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).WeaponSlot & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).HelmetSlot & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).ShieldSlot & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).LegsSlot & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).RingSlot & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).NecklaceSlot & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), packet)
    End If
End Sub

Sub GetMapWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim I As Long
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If Player(I).Char(Player(I).CharNum).Map = Player(Index).Char(Player(Index).CharNum).Map Then
                packet = "PLAYERWORNEQ" & SEP_CHAR & I & SEP_CHAR & Player(I).Char(Player(I).CharNum).ArmorSlot & SEP_CHAR & Player(I).Char(Player(I).CharNum).WeaponSlot & SEP_CHAR & Player(I).Char(Player(I).CharNum).HelmetSlot & SEP_CHAR & Player(I).Char(Player(I).CharNum).ShieldSlot & SEP_CHAR & Player(I).Char(Player(I).CharNum).LegsSlot & SEP_CHAR & Player(I).Char(Player(I).CharNum).RingSlot & SEP_CHAR & Player(I).Char(Player(I).CharNum).NecklaceSlot & END_CHAR
                Call SendDataTo(Index, packet)
            End If
        End If
    Next I
End Sub

Sub SendHP(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), "PLAYERHP" & SEP_CHAR & Index & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR)
End Sub

Sub SendMP(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR)
End Sub

Sub SendSP(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR)
End Sub

Sub SendPTS(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR)
End Sub

Sub SendEXP(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYEREXP" & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & END_CHAR)
End Sub

Sub SendStats(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR)
End Sub

Sub SendPlayerLevelToAll(ByVal Index As Long)
    Call SendDataToAll("PLAYERLEVEL" & SEP_CHAR & Index & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR)
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim packet As String
    Dim I As Long

    packet = "CLASSESDATA" & SEP_CHAR & MAX_CLASSES & SEP_CHAR
    For I = 0 To MAX_CLASSES
        packet = packet & (GetClassName(I) & SEP_CHAR & GetClassMaxHP(I) & SEP_CHAR & GetClassMaxMP(I) & SEP_CHAR & GetClassMaxSP(I) & SEP_CHAR & ClassData(I).STR & SEP_CHAR & ClassData(I).DEF & SEP_CHAR & ClassData(I).Speed & SEP_CHAR & ClassData(I).Magi & SEP_CHAR & ClassData(I).Locked & SEP_CHAR & ClassData(I).Desc & SEP_CHAR)
    Next I
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim packet As String
    Dim I As Long

    packet = "NEWCHARCLASSES" & SEP_CHAR & MAX_CLASSES & SEP_CHAR & CLASSES & SEP_CHAR
    For I = 0 To MAX_CLASSES
        packet = packet & (GetClassName(I) & SEP_CHAR & GetClassMaxHP(I) & SEP_CHAR & GetClassMaxMP(I) & SEP_CHAR & GetClassMaxSP(I) & SEP_CHAR & ClassData(I).STR & SEP_CHAR & ClassData(I).DEF & SEP_CHAR & ClassData(I).Speed & SEP_CHAR & ClassData(I).Magi & SEP_CHAR & ClassData(I).MaleSprite & SEP_CHAR & ClassData(I).FemaleSprite & SEP_CHAR & ClassData(I).Locked & SEP_CHAR & ClassData(I).Desc & SEP_CHAR)
    Next I
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Call SendDataToAllBut(Index, "left" & SEP_CHAR & Index & END_CHAR)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), "PLAYERXY" & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim packet As String

    ' Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & END_CHAR
    packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    packet = packet & Item(ItemNum).addHP & SEP_CHAR & Item(ItemNum).addMP & SEP_CHAR & Item(ItemNum).addSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    packet = packet & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim packet As String

    ' Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & END_CHAR
    packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    packet = packet & Item(ItemNum).addHP & SEP_CHAR & Item(ItemNum).addMP & SEP_CHAR & Item(ItemNum).addSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    packet = packet & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim packet As String

    packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    packet = packet & Item(ItemNum).addHP & SEP_CHAR & Item(ItemNum).addMP & SEP_CHAR & Item(ItemNum).addSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    packet = packet & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
    Call SendDataToAll("UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & END_CHAR)
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal ItemNum As Long)
    Call SendDataTo(Index, "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & END_CHAR)
End Sub

Sub SendEditEmoticonTo(ByVal Index As Long, ByVal EmoNum As Long)
    Call SendDataTo(Index, "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & END_CHAR)
End Sub

Sub SendUpdateElementToAll(ByVal ElementNum As Long)
    Call SendDataToAll("UPDATEELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendUpdateElementTo(ByVal Index As Long, ByVal ElementNum As Long)
    Call SendDataTo(Index, "UPDATEELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendEditElementTo(ByVal Index As Long, ByVal ElementNum As Long)
    Call SendDataTo(Index, "EDITELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
    Call SendDataToAll("UPDATEARROW" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & END_CHAR)
End Sub

Sub SendUpdateArrowTo(ByVal Index As Long, ByVal ItemNum As Long)
    Call SendDataTo(Index, "UPDATEARROW" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & END_CHAR)
End Sub

Sub SendEditArrowTo(ByVal Index As Long, ByVal EmoNum As Long)
    Call SendDataTo(Index, "EDITARROW" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Arrows(EmoNum).Name) & END_CHAR)
End Sub

Sub SendUpdateNpcToAll(ByVal NPCnum As Long)
    Call SendDataToAll("UPDATENPC" & SEP_CHAR & NPCnum & SEP_CHAR & Trim$(NPC(NPCnum).Name) & SEP_CHAR & NPC(NPCnum).Sprite & SEP_CHAR & NPC(NPCnum).SPRITESIZE & SEP_CHAR & NPC(NPCnum).Big & SEP_CHAR & NPC(NPCnum).MAXHP & END_CHAR)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NPCnum As Long)
    Call SendDataTo(Index, "UPDATENPC" & SEP_CHAR & NPCnum & SEP_CHAR & Trim$(NPC(NPCnum).Name) & SEP_CHAR & NPC(NPCnum).Sprite & SEP_CHAR & NPC(NPCnum).SPRITESIZE & SEP_CHAR & NPC(NPCnum).Big & SEP_CHAR & NPC(NPCnum).MAXHP & END_CHAR)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NPCnum As Long)
    Dim packet As String
    Dim I As Long

    packet = "EDITNPC" & SEP_CHAR & NPCnum & SEP_CHAR & Trim$(NPC(NPCnum).Name) & SEP_CHAR & Trim$(NPC(NPCnum).AttackSay) & SEP_CHAR & NPC(NPCnum).Sprite & SEP_CHAR & NPC(NPCnum).SpawnSecs & SEP_CHAR & NPC(NPCnum).Behavior & SEP_CHAR & NPC(NPCnum).Range & SEP_CHAR & NPC(NPCnum).STR & SEP_CHAR & NPC(NPCnum).DEF & SEP_CHAR & NPC(NPCnum).Speed & SEP_CHAR & NPC(NPCnum).Magi & SEP_CHAR & NPC(NPCnum).Big & SEP_CHAR & NPC(NPCnum).MAXHP & SEP_CHAR & NPC(NPCnum).Exp & SEP_CHAR & NPC(NPCnum).SpawnTime & SEP_CHAR & NPC(NPCnum).Element & SEP_CHAR & NPC(NPCnum).SPRITESIZE & SEP_CHAR
    For I = 1 To MAX_NPC_DROPS
        packet = packet & (NPC(NPCnum).ItemNPC(I).Chance & SEP_CHAR & NPC(NPCnum).ItemNPC(I).ItemNum & SEP_CHAR & NPC(NPCnum).ItemNPC(I).ItemValue & SEP_CHAR)
    Next I
    packet = packet & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SendShops(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_SHOPS
        If Trim$(Shop(I).Name) <> vbNullString Then
            Call SendUpdateShopTo(Index, I)
        End If
    Next I
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim packet As String
    Dim I As Integer

    packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).CurrencyItem & SEP_CHAR
    For I = 1 To MAX_SHOP_ITEMS
        packet = packet & (Shop(ShopNum).ShopItem(I).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(I).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(I).Price & SEP_CHAR)
    Next I
    packet = packet & END_CHAR

    Call SendDataToAll(packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum)
    Dim packet As String
    Dim I As Integer

    packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).CurrencyItem & SEP_CHAR
    For I = 1 To MAX_SHOP_ITEMS
        packet = packet & (Shop(ShopNum).ShopItem(I).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(I).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(I).Price & SEP_CHAR)
    Next I
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
    Dim packet As String
    Dim z As Integer

    packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).CurrencyItem & SEP_CHAR
    For z = 1 To MAX_SHOP_ITEMS
        packet = packet & (Shop(ShopNum).ShopItem(z).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(z).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(z).Price & SEP_CHAR)
    Next z
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_SPELLS
        If Trim$(Spell(I).Name) <> vbNullString Then
            Call SendUpdateSpellTo(Index, I)
        End If
    Next I
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Call SendDataToAll("UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & END_CHAR)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Call SendDataTo(Index, "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & END_CHAR)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Call SendDataTo(Index, "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).Element & END_CHAR)
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
    Call SendDataTo(Index, "GOSHOP" & SEP_CHAR & ShopNum & END_CHAR)
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim I As Long

    packet = "SPELLS" & SEP_CHAR
    For I = 1 To MAX_PLAYER_SPELLS
        packet = packet & (GetPlayerSpell(Index, I) & SEP_CHAR)
    Next I
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendWeatherTo(ByVal Index As Long)
    If WeatherLevel <= 0 Then
        WeatherLevel = 1
    End If

    Call SendDataTo(Index, "WEATHER" & SEP_CHAR & WeatherType & SEP_CHAR & WeatherLevel & END_CHAR)
End Sub

Sub SendWeatherToAll()
    Dim I As Long
    Dim Weather As String

    Select Case WeatherType
        Case 0
            Weather = "None"
        Case 1
            Weather = "Rain"
        Case 2
            Weather = "Snow"
        Case 3
            Weather = "Thunder"
    End Select

    frmServer.Label5.caption = "Current Weather: " & Weather

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendWeatherTo(I)
        End If
    Next I
End Sub

Sub SendGameClockTo(ByVal Index As Long)
    Call SendDataTo(Index, "GAMECLOCK" & SEP_CHAR & Seconds & SEP_CHAR & Minutes & SEP_CHAR & Hours & SEP_CHAR & Gamespeed & END_CHAR)
End Sub

Sub SendGameClockToAll()
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendGameClockTo(I)
        End If
    Next I
End Sub
Sub SendNewsTo(ByVal Index As Long)
    Dim packet As String
    Dim RED As Integer
    Dim GREEN As Integer
    Dim BLUE As Integer

    On Error GoTo NewsError
    RED = Val(ReadINI("COLOR", "Red", App.Path & "\News.ini", "255"))
    GREEN = Val(ReadINI("COLOR", "Green", App.Path & "\News.ini", "255"))
    BLUE = Val(ReadINI("COLOR", "Blue", App.Path & "\News.ini", "255"))

    packet = "NEWS" & SEP_CHAR & ReadINI("DATA", "NewsTitle", App.Path & "\News.ini", vbNullString) & SEP_CHAR
    packet = packet & RED & SEP_CHAR & GREEN & SEP_CHAR & BLUE & SEP_CHAR & ReadINI("DATA", "NewsBody", App.Path & "\News.ini", vbNullString) & END_CHAR

    Call SendDataTo(Index, packet)
    Exit Sub

NewsError:
    ' Error reading the news, so just send white
    RED = 255
    GREEN = 255
    BLUE = 255

    packet = "NEWS" & SEP_CHAR & ReadINI("DATA", "NewsTitle", App.Path & "\News.ini", vbNullString) & SEP_CHAR
    packet = packet & RED & SEP_CHAR & GREEN & SEP_CHAR & BLUE & SEP_CHAR & ReadINI("DATA", "NewsBody", App.Path & "\News.ini", vbNullString) & END_CHAR

    Call SendDataTo(Index, packet)
End Sub


Sub SendTimeTo(ByVal Index As Long)
    Call SendDataTo(Index, "TIME" & SEP_CHAR & GameTime & END_CHAR)
End Sub

Sub SendTimeToAll()
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendTimeTo(I)
        End If
    Next I

    Call SpawnAllMapNpcs
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal Index As Long)
    Call SendDataToMap(MapNum, "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & Index & END_CHAR)
End Sub

Sub DisabledTime()
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call DisabledTimeTo(I)
        End If
    Next I
End Sub

Sub DisabledTimeTo(ByVal Index As Long)
    Call SendDataTo(Index, "DTIME" & SEP_CHAR & TimeDisable & END_CHAR)
End Sub

Sub SendSprite(ByVal Index As Long, ByVal indexto As Long)
    Call SendDataTo(indexto, "cussprite" & SEP_CHAR & Index & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).Head & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).Body & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).Leg & END_CHAR)
End Sub

Sub GrapleHook(ByVal Index As Long)
    Dim X As Long, Y As Long, MapNum As Long
    MapNum = GetPlayerMap(Index)

    If Player(Index).HookShotX <> 0 Or Player(Index).HookShotY <> 0 Then
        If Player(Index).Locked = True Then
            Call PlayerMsg(Index, "You can only fire one grappleshot at the time", 1)
            Exit Sub
        End If
    End If

    Player(Index).Locked = True
    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)

    If GetPlayerDir(Index) = DIR_DOWN Then
        X = GetPlayerX(Index)
        Y = GetPlayerY(Index) + 1
        Do While Y <= MAX_MAPY
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Player(Index).HookShotX = X
                Player(Index).HookShotY = Y
                Exit Sub
            Else
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & END_CHAR)
                    Player(Index).HookShotX = X
                    Player(Index).HookShotY = Y
                    Exit Sub
                End If
            End If
            Y = Y + 1
        Loop
        Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & END_CHAR)
        Player(Index).HookShotX = X
        Player(Index).HookShotY = Y
        Exit Sub
    End If
    If GetPlayerDir(Index) = DIR_UP Then
        X = GetPlayerX(Index)
        Y = GetPlayerY(Index) - 1
        Do While Y >= 0
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Player(Index).HookShotX = X
                Player(Index).HookShotY = Y
                Exit Sub
            Else
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & END_CHAR)
                    Player(Index).HookShotX = X
                    Player(Index).HookShotY = Y
                    Exit Sub
                End If
            End If
            Y = Y - 1
        Loop
        Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & END_CHAR)
        Player(Index).HookShotX = X
        Player(Index).HookShotY = Y
        Exit Sub
    End If

    If GetPlayerDir(Index) = DIR_RIGHT Then
        X = GetPlayerX(Index) + 1
        Y = GetPlayerY(Index)
        Do While X <= MAX_MAPX
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Player(Index).HookShotX = X
                Player(Index).HookShotY = Y
                Exit Sub
            Else
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & END_CHAR)
                    Player(Index).HookShotX = X
                    Player(Index).HookShotY = Y
                    Exit Sub
                End If
            End If
            X = X + 1
        Loop
        Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & END_CHAR)
        Player(Index).HookShotX = X
        Player(Index).HookShotY = Y
        Exit Sub
    End If

    If GetPlayerDir(Index) = DIR_LEFT Then
        X = GetPlayerX(Index) - 1
        Y = GetPlayerY(Index)
        Do While X >= 0
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Player(Index).HookShotX = X
                Player(Index).HookShotY = Y
                Exit Sub
            Else
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & END_CHAR)
                    Player(Index).HookShotX = X
                    Player(Index).HookShotY = Y
                    Exit Sub
                End If
            End If
            X = X - 1
        Loop
        Call SendDataToMap(GetPlayerMap(Index), "hookshot" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & END_CHAR)
        Player(Index).HookShotX = X
        Player(Index).HookShotY = Y
        Exit Sub
    End If
End Sub
