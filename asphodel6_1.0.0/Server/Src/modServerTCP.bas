Attribute VB_Name = "modServerTCP"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

Public Sub UpdateCaption()
    frmServer.lblServer.Caption = "Asphodel Server <IP " & ACTUAL_IP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
    frmServer.Caption = frmServer.lblServer.Caption
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    IsConnected = (frmServer.Socket(Index).State = sckConnected)
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If Index < 1 Then Exit Function
    IsPlaying = (IsConnected(Index) And TempPlayer(Index).InGame)
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    IsLoggedIn = (IsConnected(Index) And LenB(Trim$(Player(Index).Login)) > 0)
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long

    IsMultiAccounts = False
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

    IsMultiIPOnline = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next
    
End Function

Function IsBanned(ByVal IPOrName As String, ByVal OnlyIP As Boolean) As Boolean
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\data\bans.ini"
    
    ' Check if file exists
    If Not FileExist("data\bans.ini") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    If OnlyIP Then
        If Val(GetVar(FileName, "IP", "Total")) < 1 Then Exit Function
        For F = 1 To Val(GetVar(FileName, "IP", "Total"))
            If IPOrName = GetVar(FileName, "IP", "IP" & F) Then IsBanned = True: Exit Function
        Next
    Else
        If Val(GetVar(FileName, "ACCOUNT", "Total")) < 1 Then Exit Function
        For F = 1 To Val(GetVar(FileName, "ACCOUNT", "Total"))
            If IPOrName = GetVar(FileName, "ACCOUNT", "Account" & F) Then IsBanned = True: Exit Function
        Next
    End If
    
End Function

Public Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
    If IsConnected(Index) Then
        frmServer.Socket(Index).SendData Data
        DoEvents
    End If
End Sub

Public Sub SendDataToAll(ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next
End Sub

Public Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Public Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Public Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
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

Public Sub SendMessage(ByVal Sender As Long, ByVal ChatType As Byte, ByVal ChatTag As String, ByVal Message As String, ByVal ChatColor As Long, Optional ByVal SendTo As Long = 0)
Dim SenderName As String
Dim LoopI As Long
Dim SenderColor As Long

    SenderName = GetPlayerName(Sender)
    
    Select Case GetPlayerAccess(Sender)
        Case 0
            SenderColor = Color.Brown
        Case 1
            SenderColor = Color.DarkGrey
        Case 2
            SenderColor = Color.Cyan
        Case 3
            SenderColor = Color.Blue
        Case 4
            SenderColor = Color.Pink
        Case Else
            SenderColor = Color.Black
    End Select
    
    If GetPlayerPK(Sender) > 0 Then SenderColor = Color.BrightRed
    
    Select Case ChatType
        Case E_ChatType.MapMsg_
            SendDataToMap GetPlayerMap(Sender), SMessage & SEP_CHAR & ChatTag & SEP_CHAR & SenderColor & SEP_CHAR & SenderName & SEP_CHAR & Message & SEP_CHAR & ChatColor & END_CHAR
        Case E_ChatType.EmoteMsg_
            SendDataToMap GetPlayerMap(Sender), SMessage & SEP_CHAR & ChatTag & SEP_CHAR & SenderColor & SEP_CHAR & SenderName & SEP_CHAR & Message & SEP_CHAR & ChatColor & END_CHAR
        Case E_ChatType.BroadcastMsg_
            SendDataToAll SMessage & SEP_CHAR & ChatTag & SEP_CHAR & SenderColor & SEP_CHAR & SenderName & SEP_CHAR & Message & SEP_CHAR & ChatColor & END_CHAR
        Case E_ChatType.GlobalMsg_
            SendDataToAll SMessage & SEP_CHAR & ChatTag & SEP_CHAR & SenderColor & SEP_CHAR & SenderName & SEP_CHAR & Message & SEP_CHAR & ChatColor & END_CHAR
        Case E_ChatType.AdminMsg_
            For LoopI = 1 To MAX_PLAYERS
                If IsPlaying(LoopI) Then
                    If GetPlayerAccess(LoopI) > 0 Then
                        SendDataTo LoopI, SMessage & SEP_CHAR & ChatTag & SEP_CHAR & SenderColor & SEP_CHAR & SenderName & SEP_CHAR & Message & SEP_CHAR & ChatColor & END_CHAR
                    End If
                End If
            Next
        Case E_ChatType.PrivateMsg_
            SendDataTo Sender, SMessage & SEP_CHAR & ChatTag & SEP_CHAR & SenderColor & SEP_CHAR & SenderName & SEP_CHAR & Message & SEP_CHAR & ChatColor & END_CHAR
            SendDataTo SendTo, SMessage & SEP_CHAR & ChatTag & SEP_CHAR & SenderColor & SEP_CHAR & SenderName & SEP_CHAR & Message & SEP_CHAR & ChatColor & END_CHAR
    End Select

End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = SGlobalMsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim i As Long

    Packet = SAdminMsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerAccess(i) > 0 Then
                Call SendDataTo(i, Packet)
            End If
        End If
    Next
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = SPlayerMsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
    
    Packet = SMapMsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Public Sub NormalMsg(ByVal Index As Long, ByVal Message As String, ByVal WindowState As Long)

    SendDataTo Index, SNormalMsg & SEP_CHAR & Message & SEP_CHAR & WindowState & END_CHAR
    
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = SAlertMsg & SEP_CHAR & Msg & END_CHAR
    
    Call SendDataTo(Index, Packet)
    Call CloseSocket(Index)
End Sub

Public Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", Color.White)
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
    If Index <> 0 Then
        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(Index), True) Then
                Call TextAdd(frmServer.txtText, "Received connection from " & GetPlayerIP(Index) & ".")
            Else
                Call AlertMsg(Index, "You are banned from " & GAME_NAME & "!" & vbNewLine & "Contact an admin at " & GAME_WEBSITE & ".")
            End If
        'Else
           ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    End If
End Sub

Public Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim start As Long

    If Index > 0 Then
    
        frmServer.Socket(Index).GetData Buffer, vbString, DataLength
            
        TempPlayer(Index).Buffer = TempPlayer(Index).Buffer & Buffer
        
        start = InStr(TempPlayer(Index).Buffer, END_CHAR)
        Do While start > 0
            Packet = Mid$(TempPlayer(Index).Buffer, 1, start - 1)
            TempPlayer(Index).Buffer = Mid$(TempPlayer(Index).Buffer, start + 1, Len(TempPlayer(Index).Buffer))
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            start = InStr(TempPlayer(Index).Buffer, END_CHAR)
            
            If Len(Packet) > 0 Then
                Call HandleData(Index, Packet)
            End If
        Loop
                
        ' Check if elapsed time has passed
        TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
        
        If TempPlayer(Index).DataTimer < GetTickCountNew Then
            TempPlayer(Index).DataTimer = GetTickCountNew + 1000
            TempPlayer(Index).DataBytes = 0
            TempPlayer(Index).DataPackets = 0
            Exit Sub
        End If
        
        If GetPlayerAccess(Index) <= StaffType.Monitor Then
            ' Check for data flooding
            If TempPlayer(Index).DataBytes > 1000 Then
                Call HackingAttempt(Index, "Data Flooding")
                Exit Sub
            End If
            
            ' Check for packet flooding
            If TempPlayer(Index).DataPackets > 25 Then
                Call HackingAttempt(Index, "Packet Flooding")
                Exit Sub
            End If
        End If
        
    End If
End Sub

Public Sub CloseSocket(ByVal Index As Long)

    If Index > 0 Then
        Call LeftGame(Index)
    
        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(Index) & " has been terminated.")
        
        frmServer.Socket(Index).Close
            
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If
    
End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
Dim MapData As String
Dim X As Long
Dim Y As Long
Dim i As Long

    MapData = SMapData & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY
    
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            With Map(MapNum).Tile(X, Y)
                For i = 0 To UBound(.Layer)
                    MapData = MapData & SEP_CHAR & .Layer(i)
                    MapData = MapData & SEP_CHAR & .LayerSet(i)
                Next
                MapData = MapData & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3
            End With
        Next
    Next
    
    MapData = MapData & SEP_CHAR & UBound(MapSpawn(MapNum).Npc)
    
    For X = 1 To UBound(MapSpawn(MapNum).Npc)
        MapData = MapData & SEP_CHAR & MapSpawn(MapNum).Npc(X).Num & SEP_CHAR & MapSpawn(MapNum).Npc(X).X & SEP_CHAR & MapSpawn(MapNum).Npc(X).Y
    Next
    
    MapData = MapData & END_CHAR
    
    MapCache(MapNum) = MapData
    
End Sub

' ******************************
' ** Outcoming Server Packets **
' ******************************

Public Sub SendWhosOnline(ByVal Index As Long)
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

Public Sub SendChars(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = SAllChars
    
    For i = 1 To MAX_CHARS
        Packet = Packet & SEP_CHAR & Trim$(Player(Index).Char(i).Name) & SEP_CHAR & Trim$(Class(Player(Index).Char(i).Class).Name) & SEP_CHAR & Player(Index).Char(i).Level & SEP_CHAR & Player(Index).Char(i).Sprite
    Next
    
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendPlayerGuildTo(ByVal Index As Long, ByVal TargetIndex As Long)

    If Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).Guild < 1 Then
        SendDataTo Index, SPlayerGuild & SEP_CHAR & TargetIndex & END_CHAR
    Else
        SendDataTo Index, SPlayerGuild & SEP_CHAR & TargetIndex & SEP_CHAR & Trim$(Guild(Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).Guild).Name) & SEP_CHAR & Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank & END_CHAR
    End If
    
End Sub

Public Sub SendPlayerGuildToAll(ByVal Index As Long)

    If Player(Index).Char(TempPlayer(Index).CharNum).Guild < 1 Then
        SendDataToAll SPlayerGuild & SEP_CHAR & Index & END_CHAR
    Else
        SendDataToAll SPlayerGuild & SEP_CHAR & Index & SEP_CHAR & Trim$(Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Name) & SEP_CHAR & Player(Index).Char(TempPlayer(Index).CharNum).GuildRank & END_CHAR
    End If
    
End Sub

Public Sub SendJoinMap(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    Packet = SPlayerData & SEP_CHAR & i & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & GetPlayerSprite(i) & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & GetPlayerX(i) & SEP_CHAR & GetPlayerY(i) & SEP_CHAR & GetPlayerDir(i) & SEP_CHAR & GetPlayerAccess(i) & SEP_CHAR & GetPlayerPK(i) & END_CHAR
                    Call SendDataTo(Index, Packet)
                    Packet = SPlayerHp & SEP_CHAR & GetPlayerMaxVital(i, Vitals.HP, True) & SEP_CHAR & GetPlayerVital(i, Vitals.HP) & SEP_CHAR & i & END_CHAR
                    SendDataTo Index, Packet
                End If
            End If
        End If
    Next
    
    ' Send index's player data to everyone on the map including himself
    Packet = SPlayerData & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
    Packet = SPlayerHp & SEP_CHAR & GetPlayerMaxVital(Index, Vitals.HP, True) & SEP_CHAR & GetPlayerVital(Index, Vitals.HP) & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
    
    For i = 1 To UBound(MapSpawn(GetPlayerMap(Index)).Npc)
        If MapNpc(GetPlayerMap(Index)).MapNpc(i).Num > 0 Then
            SendNPCVital GetPlayerMap(Index), i
        End If
    Next
    
End Sub

Public Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String

    Packet = SLeft & SEP_CHAR & Index & END_CHAR
    Call SendDataToMapBut(Index, MapNum, Packet)
    
End Sub

Public Sub SendPlayerData(ByVal Index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = SPlayerData & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
    
End Sub

Public Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Call SendDataTo(Index, MapCache(MapNum))
End Sub

Public Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = SMapItemData
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & SEP_CHAR & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y & SEP_CHAR & MapItem(MapNum, i).Anim
    Next
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = SMapItemData
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & SEP_CHAR & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y
    Next
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Public Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = SMapNpcData
    Packet = Packet & SEP_CHAR & UBound(MapSpawn(MapNum).Npc)
    For i = 1 To UBound(MapSpawn(MapNum).Npc)
        Packet = Packet & SEP_CHAR & MapNpc(MapNum).MapNpc(i).Num & SEP_CHAR & MapNpc(MapNum).MapNpc(i).X & SEP_CHAR & MapNpc(MapNum).MapNpc(i).Y & SEP_CHAR & MapNpc(MapNum).MapNpc(i).Dir
    Next
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = SMapNpcData
    Packet = Packet & SEP_CHAR & UBound(MapSpawn(MapNum).Npc)
    For i = 1 To UBound(MapSpawn(MapNum).Npc)
        Packet = Packet & SEP_CHAR & MapNpc(MapNum).MapNpc(i).Num & SEP_CHAR & MapNpc(MapNum).MapNpc(i).X & SEP_CHAR & MapNpc(MapNum).MapNpc(i).Y & SEP_CHAR & MapNpc(MapNum).MapNpc(i).Dir
    Next
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
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
End Sub

Public Sub SendInventory(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = SPlayerInv & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemDur(Index, i) & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = SPlayerInvUpdate & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendWornEquipment(ByVal Index As Long)
Dim Packet As String

    Packet = SPlayerWornEq & SEP_CHAR & GetPlayerEquipmentSlot(Index, Armor) & SEP_CHAR & GetPlayerEquipmentSlot(Index, Weapon) & SEP_CHAR & GetPlayerEquipmentSlot(Index, Helmet) & SEP_CHAR & GetPlayerEquipmentSlot(Index, Shield) & END_CHAR
    Call SendDataTo(Index, Packet)
    
    SendStats Index
    
End Sub

Public Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
Dim Packet As String

    Select Case Vital
        Case HP
            Packet = SPlayerHp & SEP_CHAR & GetPlayerMaxVital(Index, Vitals.HP, True) & SEP_CHAR & GetPlayerVital(Index, Vitals.HP) & SEP_CHAR & Index & END_CHAR
        Case MP
            Packet = SPlayerMp & SEP_CHAR & GetPlayerMaxVital(Index, Vitals.MP, True) & SEP_CHAR & GetPlayerVital(Index, Vitals.MP) & END_CHAR
        Case SP
            Packet = SPlayerSp & SEP_CHAR & GetPlayerMaxVital(Index, Vitals.SP, True) & SEP_CHAR & GetPlayerVital(Index, Vitals.SP) & END_CHAR
    End Select
    
    Call SendDataTo(Index, Packet)
    
    If Vital = HP Then SendDataToMap GetPlayerMap(Index), Packet
    
End Sub

Public Sub SendStats(ByVal Index As Long)
Dim Packet As String

    Packet = SPlayerStats & SEP_CHAR & GetPlayerStat(Index, Stats.Strength) & SEP_CHAR & GetPlayerStat(Index, Stats.Defense) & SEP_CHAR & GetPlayerStat(Index, Stats.SPEED) & SEP_CHAR & GetPlayerStat(Index, Stats.Magic) & END_CHAR
    Call SendDataTo(Index, Packet)
    
    SendStatBuffs Index
    
End Sub

Private Sub SendStatBuffs(ByVal Index As Long)
Dim Packet As String

    Packet = SPlayerStatBuffs & SEP_CHAR & GetPlayerStat_withBonus(Index, Stats.Strength) & SEP_CHAR & GetPlayerStat_withBonus(Index, Stats.Defense) & SEP_CHAR & GetPlayerStat_withBonus(Index, Stats.SPEED) & SEP_CHAR & GetPlayerStat_withBonus(Index, Stats.Magic) & END_CHAR
    Call SendDataTo(Index, Packet)
    
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
    
    If Player(Index).Char(TempPlayer(Index).CharNum).Muted Then
        PlayerMsg Index, "You are still muted for " & CInt((Player(Index).Char(TempPlayer(Index).CharNum).MuteTime - GetTickCountNew) / 60000) & " minutes!", Color.BrightRed
        AdminMsg "Staff: " & GetPlayerName(Index) & " (access " & GetPlayerAccess(Index) & ") is still muted for " & CInt((Player(Index).Char(TempPlayer(Index).CharNum).MuteTime - GetTickCountNew) / 60000) & " minutes!", Color.BrightRed
    End If
    
End Sub

Public Sub SendClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = SClassesData & SEP_CHAR & Max_Classes
    For i = 1 To Max_Classes
        Packet = Packet & SEP_CHAR & GetClassName(i) & SEP_CHAR & GetClassMaxVital(i, Vitals.HP) & SEP_CHAR & GetClassMaxVital(i, Vitals.MP) & SEP_CHAR & GetClassMaxVital(i, Vitals.SP) & SEP_CHAR & Class(i).Stat(Stats.Strength) & SEP_CHAR & Class(i).Stat(Stats.Defense) & SEP_CHAR & Class(i).Stat(Stats.SPEED) & SEP_CHAR & Class(i).Stat(Stats.Magic)
    Next
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendNewCharClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = SNewCharClasses & SEP_CHAR & Max_Classes
    For i = 1 To Max_Classes
        Packet = Packet & SEP_CHAR & GetClassName(i) & SEP_CHAR & GetClassMaxVital(i, Vitals.HP) & SEP_CHAR & GetClassMaxVital(i, Vitals.MP) & SEP_CHAR & GetClassMaxVital(i, Vitals.SP) & SEP_CHAR & Class(i).Stat(Stats.Strength) & SEP_CHAR & Class(i).Stat(Stats.Defense) & SEP_CHAR & Class(i).Stat(Stats.SPEED) & SEP_CHAR & Class(i).Stat(Stats.Magic)
    Next
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendLeftGame(ByVal Index As Long)
Dim Packet As String

    Packet = SPlayerData & SEP_CHAR & Index & SEP_CHAR & vbNullString & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR
    Call SendDataToAllBut(Index, Packet)
    
End Sub

Public Sub SendPlayerXY(ByVal Index As Long)
Dim Packet As String

    Packet = SPlayerXY & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String
Dim LoopI As Long

    For LoopI = 1 To MAX_PLAYERS
        If IsPlaying(LoopI) Then SendStats LoopI
    Next
    
    Packet = SUpdateItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Anim & SEP_CHAR & Item(ItemNum).CostItem & SEP_CHAR & Item(ItemNum).CostAmount & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3
    
    For LoopI = 1 To Stats.Stat_Count - 1
        Packet = Packet & SEP_CHAR & Item(ItemNum).BuffStats(LoopI)
    Next
    
    For LoopI = 1 To Vitals.Vital_Count - 1
        Packet = Packet & SEP_CHAR & Item(ItemNum).BuffVitals(LoopI)
    Next
    
    For LoopI = 1 To Item_Requires.Count - 1
        Packet = Packet & SEP_CHAR & Item(ItemNum).Required(LoopI)
    Next
    
    Packet = Packet & END_CHAR
    
    Call SendDataToAll(Packet)
    
End Sub

Public Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String
Dim LoopI As Long

    Packet = SUpdateItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Anim & SEP_CHAR & Item(ItemNum).CostItem & SEP_CHAR & Item(ItemNum).CostAmount & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3
    
    For LoopI = 1 To Stats.Stat_Count - 1
        Packet = Packet & SEP_CHAR & Item(ItemNum).BuffStats(LoopI)
    Next
    
    For LoopI = 1 To Vitals.Vital_Count - 1
        Packet = Packet & SEP_CHAR & Item(ItemNum).BuffVitals(LoopI)
    Next
    
    For LoopI = 1 To Item_Requires.Count - 1
        Packet = Packet & SEP_CHAR & Item(ItemNum).Required(LoopI)
    Next
    
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String
Dim LoopI As Long

    Packet = SEditItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Durability & SEP_CHAR & Item(ItemNum).Anim & SEP_CHAR & Item(ItemNum).CostItem & SEP_CHAR & Item(ItemNum).CostAmount
    
    For LoopI = 1 To Stats.Stat_Count - 1
        Packet = Packet & SEP_CHAR & Item(ItemNum).BuffStats(LoopI)
    Next
    
    For LoopI = 1 To Vitals.Vital_Count - 1
        Packet = Packet & SEP_CHAR & Item(ItemNum).BuffVitals(LoopI)
    Next
    
    For LoopI = 0 To Item_Requires.Count - 1
        Packet = Packet & SEP_CHAR & Item(ItemNum).Required(LoopI)
    Next
    
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Packet As String

    Packet = SUpdateNpc & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).HP & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Public Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String

    Packet = SUpdateNpc & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).HP & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String
Dim i As Long

    Packet = SEditNpc & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).Stat(Stats.Strength) & SEP_CHAR & Npc(NpcNum).Stat(Stats.Defense) & SEP_CHAR & Npc(NpcNum).Stat(Stats.SPEED) & SEP_CHAR & Npc(NpcNum).Stat(Stats.Magic) & SEP_CHAR & Npc(NpcNum).GivesGuild
    
    For i = 0 To UBound(Npc(NpcNum).Sound)
        Packet = Packet & SEP_CHAR & Npc(NpcNum).Sound(i)
    Next
    
    For i = 0 To UBound(Npc(NpcNum).Reflection)
        Packet = Packet & SEP_CHAR & Npc(NpcNum).Reflection(i)
    Next
    
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendShops(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If
    Next
End Sub

Public Sub SendUpdateAnimToAll(ByVal AnimNum As Long)
Dim Packet As String

    Packet = SUpdateAnim & SEP_CHAR & AnimNum & SEP_CHAR & Trim$(Animation(AnimNum).Name) & SEP_CHAR & Animation(AnimNum).Height & SEP_CHAR & Animation(AnimNum).Width & SEP_CHAR & Animation(AnimNum).Pic & SEP_CHAR & Animation(AnimNum).Delay & END_CHAR
    Call SendDataToAll(Packet)
    
End Sub

Public Sub SendUpdateAnimTo(ByVal Index As Long, ByVal AnimNum As Long)
Dim Packet As String

    Packet = SUpdateAnim & SEP_CHAR & AnimNum & SEP_CHAR & Trim$(Animation(AnimNum).Name) & SEP_CHAR & Animation(AnimNum).Height & SEP_CHAR & Animation(AnimNum).Width & SEP_CHAR & Animation(AnimNum).Pic & SEP_CHAR & Animation(AnimNum).Delay & END_CHAR
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String

    Packet = SUpdateShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Public Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String

    Packet = SUpdateShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = SEditShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems
    For i = 1 To MAX_TRADES
        Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue
    Next
    Packet = Packet & END_CHAR

    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendEditAnimTo(ByVal Index As Long, ByVal AnimNum As Long)
Dim Packet As String

    With Animation(AnimNum)
        Packet = SEditAnim & SEP_CHAR & AnimNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .Delay & SEP_CHAR & .Height & SEP_CHAR & .Width & END_CHAR
    End With
    
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendSpells(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If
    Next
End Sub

Public Sub SendSigns(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SIGNS
        If LenB(Trim$(Sign(i).Name)) > 0 Then
            Call SendUpdateSignTo(Index, i)
        End If
    Next
    
End Sub

Public Sub SendUpdateSignToAll(ByVal SignNum As Long)
Dim Packet As String

    Packet = SUpdateSign & SEP_CHAR & SignNum & SEP_CHAR & Trim$(Sign(SignNum).Name) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Public Sub SendUpdateSignTo(ByVal Index As Long, ByVal SignNum As Long)
Dim Packet As String

    Packet = SUpdateSign & SEP_CHAR & SignNum & SEP_CHAR & Trim$(Sign(SignNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String

    Packet = SUpdateSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).Anim & SEP_CHAR & Spell(SpellNum).Icon & SEP_CHAR & Spell(SpellNum).Timer & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).AOE & END_CHAR
    Call SendDataToAll(Packet)
    
End Sub

Public Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = SUpdateSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).Anim & SEP_CHAR & Spell(SpellNum).Icon & SEP_CHAR & Spell(SpellNum).Timer & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).AOE & END_CHAR
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    With Spell(SpellNum)
        Packet = SEditSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .MPReq & SEP_CHAR & .Type & SEP_CHAR & .Anim & SEP_CHAR & .Icon & SEP_CHAR & .Range & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & Trim$(.CastSound) & SEP_CHAR & .AOE & SEP_CHAR & .Timer & END_CHAR
    End With
    
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendEditSignTo(ByVal Index As Long, ByVal SignNum As Long)
Dim Packet As String
Dim LoopI As Long

    With Sign(SignNum)
        Packet = SEditSign & SEP_CHAR & SignNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & UBound(.Section)
        For LoopI = 0 To UBound(.Section)
            Packet = Packet & SEP_CHAR & Trim$(.Section(LoopI))
        Next
    End With
    
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Z As Long

    Packet = STrade & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay)
    
    For i = 1 To MAX_TRADES
        Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue
    Next
    
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
    
End Sub

Public Sub SendPlayerSpells(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = SSpells
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & SEP_CHAR & GetPlayerSpell(Index, i)
    Next
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Public Sub SendGameOptions(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = SGameOptions & SEP_CHAR & GAME_NAME & SEP_CHAR & GAME_WEBSITE & SEP_CHAR & SPRITE_OFFSET & SEP_CHAR & TOTAL_WALKFRAMES & SEP_CHAR & TOTAL_ATTACKFRAMES & SEP_CHAR & WALKANIM_SPEED & SEP_CHAR & TOTAL_ANIMFRAMES & SEP_CHAR & CONFIG_STANDFRAME
    
    For i = 0 To 3
        Packet = Packet & SEP_CHAR & Direction_Anim(i)
    Next
    
    Packet = Packet & SEP_CHAR & MAX_PLAYERS & SEP_CHAR & MAX_SHOPS & SEP_CHAR & MAX_SPELLS & SEP_CHAR & MAX_ITEMS & SEP_CHAR & MAX_NPCS & SEP_CHAR & MAX_MAPS & SEP_CHAR & MAX_SIGNS & SEP_CHAR & MAX_ANIMS & SEP_CHAR & GAME_NEWS
    
    If TOTAL_WALKFRAMES > 0 Then
        For i = 1 To TOTAL_WALKFRAMES
            Packet = Packet & SEP_CHAR & WalkFrame(i)
        Next
    End If
    
    If TOTAL_ATTACKFRAMES > 0 Then
        For i = 1 To TOTAL_ATTACKFRAMES
            Packet = Packet & SEP_CHAR & AttackFrame(i)
        Next
    End If
    
    Packet = Packet & END_CHAR
    
    SendDataTo Index, Packet
    
End Sub

Public Sub SendSound(ByVal MapNum As Long, ByVal Sound As String)
    SendDataToMap MapNum, SSoundPlay & SEP_CHAR & Sound & END_CHAR
End Sub

Public Sub SendNPCVital(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    SendDataToMap MapNum, SNpcHP & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum).MapNpc(MapNpcNum).Vital(Vitals.HP) & END_CHAR
End Sub
