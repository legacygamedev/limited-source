Attribute VB_Name = "modServerTCP"
Option Explicit
Public MuteBroadcast As Boolean

Sub UpdateCaption()
On Error GoTo ErrorHandler
    frmServer.Caption = GAME_NAME & " Server <IP " & frmServer.Socket(0).LocalIP & " Port " & STR(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "UpdateCaption", Err.Number, Err.Description
End Sub

Function IsConnected(ByVal index As Long) As Boolean
On Error GoTo ErrorHandler
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsConnected", Err.Number, Err.Description
End Function

Function IsPlaying(ByVal index As Long) As Boolean
On Error GoTo ErrorHandler
    If IsConnected(index) And Player(index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsPlaying", Err.Number, Err.Description
End Function

Function IsLoggedIn(ByVal index As Long) As Boolean
On Error GoTo ErrorHandler
    If IsConnected(index) And Trim(Player(index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsLoggedIn", Err.Number, Err.Description
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
On Error GoTo ErrorHandler
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase(Trim(Player(i).Login)) = LCase(Trim(Login)) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsMultiAccounts", Err.Number, Err.Description
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
On Error GoTo ErrorHandler
Dim i As Long
Dim n As Long

    n = 0
    IsMultiIPOnline = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And Trim(GetPlayerIP(i)) = Trim(IP) Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsMultiIPOnline", Err.Number, Err.Description
End Function

Function IsBanned(ByVal IP As String) As Boolean
On Error GoTo ErrorHandler

Dim FileName As String, fIP As String, fName As String
Dim f As Long

    IsBanned = False
    
    FileName = App.Path & "\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("banlist.txt") Then
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
        If Trim(LCase(fIP)) = Trim(LCase(Mid(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If
    Loop
    
    Close #f
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsBanned", Err.Number, Err.Description
End Function

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long, n As Long, startc As Long

    If IsConnected(index) Then
        frmServer.Socket(index).SendData Data
        DoEvents
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataTo", Err.Number, Err.Description
End Sub

Sub SendDataToAll(ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataToAll", Err.Number, Err.Description
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            Call SendDataTo(i, Data)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataToAllBut", Err.Number, Err.Description
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataToMap", Err.Number, Err.Description
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataToMapBut", Err.Number, Err.Description
End Sub

Sub BroadcastMsg(ByVal Msg As String, ByVal Color As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "broadcastmsg" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "BroadcastMsg", Err.Number, Err.Description
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "GlobalMsg", Err.Number, Err.Description
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "AdminMsg", Err.Number, Err.Description
End Sub

Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "PlayerMsg", Err.Number, Err.Description
End Sub

Sub SendParty(ByVal index As Long)
'On Error GoTo ErrorHandler
Dim i As Long
Dim n As Long
Dim Packet As String
    
    If Player(index).Party.Started = YES Then
        n = index
    Else
        n = Player(index).Party.PlayerNums(1)
    End If
    'MsgBox "n = " & n
    
    If n > 0 Then
        Packet = "partydisplay" & SEP_CHAR & GetPlayerName(n) & SEP_CHAR
        For i = 1 To MAX_PARTY_MEMS
            If Player(n).Party.PlayerNums(i) > 0 Then
            Packet = Packet & GetPlayerName(Player(n).Party.PlayerNums(i)) & SEP_CHAR
            End If
        Next i
        Packet = Packet & END_CHAR
    
        Call SendDataTo(n, Packet)
        For i = 1 To MAX_PARTY_MEMS
            If Player(n).Party.PlayerNums(i) > 0 Then
                Call SendDataTo(Player(n).Party.PlayerNums(i), Packet)
            End If
        Next i
    Else
        Packet = "noparty" & SEP_CHAR & END_CHAR
        Call SendDataTo(index, Packet)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendParty", Err.Number, Err.Description
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim Text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "MapMsg", Err.Number, Err.Description
End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "AlertMsg", Err.Number, Err.Description
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
On Error GoTo ErrorHandler
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "HackingAttempt", Err.Number, Err.Description
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
On Error GoTo ErrorHandler
Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot
        
        If i <> 0 Then
            ' Whoho, we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "AcceptConnection", Err.Number, Err.Description
End Sub

Sub SocketConnected(ByVal index As Long)
On Error GoTo ErrorHandler
    If index <> 0 Then
        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(index)) Then
                Call TextAdd(frmServer.txtText, "Received connection from " & GetPlayerIP(index) & ".", True)
            Else
                Call AlertMsg(index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
            End If
        'Else
           ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SocketConnected", Err.Number, Err.Description
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
On Error GoTo ErrorHandler
Dim Buffer As String
Dim Packet As String
Dim top As String * 3
Dim Start As Long

    If index > 0 Then
        frmServer.Socket(index).GetData Buffer, vbString, DataLength
        
        If Buffer = "top" Then
            top = STR(TotalOnlinePlayers)
            Call SendDataTo(index, top)
            Call CloseSocket(index)
        End If
            
        Player(index).Buffer = Player(index).Buffer & Buffer
        
        Start = InStr(Player(index).Buffer, END_CHAR)
        Do While Start > 0
            Packet = Mid(Player(index).Buffer, 1, Start - 1)
            Player(index).Buffer = Mid(Player(index).Buffer, Start + 1, Len(Player(index).Buffer))
            Player(index).DataPackets = Player(index).DataPackets + 1
            Start = InStr(Player(index).Buffer, END_CHAR)
            If Len(Packet) > 0 Then
                Call HandleData(index, Packet)
            End If
        Loop
                
        ' Check if elapsed time has passed
        Player(index).DataBytes = Player(index).DataBytes + DataLength
        If GetTickCount >= Player(index).DataTimer + 1000 Then
            Player(index).DataTimer = GetTickCount
            Player(index).DataBytes = 0
            Player(index).DataPackets = 0
            Exit Sub
        End If
        
        ' Check for data flooding
        If Player(index).DataBytes > 1500 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Data Flooding")
            Exit Sub
        End If
        
        ' Check for packet flooding
        If Player(index).DataPackets > 50 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Packet Flooding")
            Exit Sub
        End If
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "IncomingData", Err.Number, Err.Description
End Sub

Sub CloseSocket(ByVal index As Long)
On Error GoTo ErrorHandler
    ' Make sure player was/is playing the game, and if so, save'm.
    If index > 0 Then
        Call LeftGame(index)
    
        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(index) & " has been terminated.", True)
        
        frmServer.Socket(index).Close
            
        Call UpdateCaption
        Call ClearPlayer(index)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "CloseSocket", Err.Number, Err.Description
End Sub

Sub SendWhosOnline(ByVal index As Long)
On Error GoTo ErrorHandler
Dim s As String, d As String
Dim n As Long, i As Long, x As Long, c As Long

    s = ""
    d = ""
    x = 0
    c = 0
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) = True Then
            If GetPlayerAccess(i) >= 1 Then
                x = x + 1
                If d = "" Then
                    d = d & GetPlayerName(i)
                Else
                    d = d & ", " & GetPlayerName(i)
                End If
            Else
                c = c + 1
                If s = "" Then
                    s = s & GetPlayerName(i)
                Else
                    s = s & ", " & GetPlayerName(i)
                End If
            End If
            n = n + 1
        End If
    Next i
    
    Call PlayerMsg(index, "Delorians Online: " & n, Yellow)
    Call PlayerMsg(index, "GAME MODERATORS: " & x, Yellow)
    Call PlayerMsg(index, d, Yellow)
    Call PlayerMsg(index, "PLAYERS: " & c, Yellow)
    Call PlayerMsg(index, s, Yellow)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendWhosOnline", Err.Number, Err.Description
End Sub

Sub SendOnlineList()
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long
Dim n As Long
Packet = ""
n = 0
For i = 1 To MAX_PLAYERS
    If IsPlaying(i) Then
        Packet = Packet & SEP_CHAR & GetPlayerName(i) & SEP_CHAR
        n = n + 1
    End If
Next i

Packet = "ONLINELIST" & SEP_CHAR & n & Packet & END_CHAR

Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendOnlineList", Err.Number, Err.Description
End Sub

Sub SendChars(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim(Player(index).Char(i).Name) & SEP_CHAR & Trim(Class(Player(index).Char(i).Class).Name) & SEP_CHAR & Player(index).Char(i).Level & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendChars", Err.Number, Err.Description
End Sub

Sub SendJoinMap(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = ""
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index And GetPlayerMap(i) = GetPlayerMap(index) Then
            Packet = "PLAYERDATA" & SEP_CHAR
            Packet = Packet & i & SEP_CHAR
            Packet = Packet & GetPlayerName(i) & SEP_CHAR
            Packet = Packet & GetPlayerSprite(i) & SEP_CHAR
            Packet = Packet & GetPlayerMap(i) & SEP_CHAR
            Packet = Packet & GetPlayerX(i) & SEP_CHAR
            Packet = Packet & GetPlayerY(i) & SEP_CHAR
            Packet = Packet & GetPlayerDir(i) & SEP_CHAR
            Packet = Packet & GetPlayerAccess(i) & SEP_CHAR
            Packet = Packet & GetPlayerPK(i) & SEP_CHAR
            Packet = Packet & GetPlayerGuild(i) & SEP_CHAR
            Packet = Packet & GetPlayerGuildAccess(i) & SEP_CHAR
            Packet = Packet & GetPlayerClass(i) & SEP_CHAR
            Packet = Packet & Player(i).Char(Player(i).CharNum).Sex & SEP_CHAR
            Packet = Packet & END_CHAR
            Call SendDataTo(index, Packet)
        End If
    Next i
    
    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & GetPlayerName(index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(index) & SEP_CHAR
    Packet = Packet & GetPlayerX(index) & SEP_CHAR
    Packet = Packet & GetPlayerY(index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(index) & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).Sex & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendJoinMap", Err.Number, Err.Description
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & GetPlayerName(index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(index) & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & GetPlayerX(index) & SEP_CHAR
    Packet = Packet & GetPlayerY(index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(index) & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).Sex & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMapBut(index, MapNum, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendLeaveMap", Err.Number, Err.Description
End Sub

Sub SendPlayerData(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & GetPlayerName(index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(index) & SEP_CHAR
    Packet = Packet & GetPlayerX(index) & SEP_CHAR
    Packet = Packet & GetPlayerY(index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(index) & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).Sex & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendPlayerData", Err.Number, Err.Description
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
On Error GoTo ErrorHandler

Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long
Dim i As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
        With Map(MapNum).Tile(x, y)
            Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR
        End With
        Next x
    Next y
    
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).Npc(x) & SEP_CHAR
    Next x
    
    Packet = Packet & END_CHAR
    
    x = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, x)
    P2 = Mid(Packet, x + 1, Len(Packet) - x)
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendMap", Err.Number, Err.Description
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)
On Error GoTo ErrorHandler

Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendMapItemsTo", Err.Number, Err.Description
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendMapItemsToAll", Err.Number, Err.Description
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendMapNpcsTo", Err.Number, Err.Description
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendMapNpcsToMap", Err.Number, Err.Description
End Sub

Sub SendItems(ByVal index As Long)
On Error GoTo ErrorHandler

Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(index, i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendItems", Err.Number, Err.Description
End Sub

Sub SendEmoticons(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    For i = 0 To MAX_EMOTICONS
        If Trim(Emoticons(i).Command) <> "" Then
            Call SendUpdateEmoticonTo(index, i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendEmoticons", Err.Number, Err.Description
End Sub

Sub SendNpcs(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendNpcs", Err.Number, Err.Description
End Sub

Sub SendBank(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "PLAYERBANK" & SEP_CHAR
    For i = 1 To MAX_BANK
        Packet = Packet & GetPlayerBankItemNum(index, i) & SEP_CHAR & GetPlayerBankItemValue(index, i) & SEP_CHAR & GetPlayerBankItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendBank", Err.Number, Err.Description
End Sub

Sub SendBankUpdate(ByVal index As Long, ByVal BankSlot As Long)
On Error GoTo ErrorHandler
Dim Packet As String
    
    Packet = "PLAYERBANKUPDATE" & SEP_CHAR & BankSlot & SEP_CHAR & GetPlayerBankItemNum(index, BankSlot) & SEP_CHAR & GetPlayerBankItemValue(index, BankSlot) & SEP_CHAR & GetPlayerBankItemDur(index, BankSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendBankUpdate", Err.Number, Err.Description
End Sub

Sub SendInventory(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendInventory", Err.Number, Err.Description
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
On Error GoTo ErrorHandler
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendInventoryUpdate", Err.Number, Err.Description
End Sub

Sub SendWornEquipment(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
    
    Packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & GetPlayerBootsSlot(index) & SEP_CHAR & GetPlayerGlovesSlot(index) & SEP_CHAR & GetPlayerRingSlot(index) & SEP_CHAR & GetPlayerAmuletSlot(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendWornEquipment", Err.Number, Err.Description
End Sub

Sub SendHP(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendHP", Err.Number, Err.Description
End Sub

Sub SendMP(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendMP", Err.Number, Err.Description
End Sub

Sub SendSP(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendSP", Err.Number, Err.Description
End Sub

Sub SendStats(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
    
    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & GetPlayerVIT(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendStats", Err.Number, Err.Description
End Sub

Sub PlayerPoints(ByVal index As Long)
On Error GoTo ErrorHandler
    Call SendDataTo(index, "playerpoints" & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "PlayerPoints", Err.Number, Err.Description
End Sub

Sub PlayerClass(ByVal index As Long)
On Error GoTo ErrorHandler
    Call SendDataTo(index, "playerclass" & SEP_CHAR & GetPlayerClass(index) & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "PlayerClass", Err.Number, Err.Description
End Sub

Sub SendClasses(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).Locked & SEP_CHAR & Class(i).VIT & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendClasses", Err.Number, Err.Description
End Sub

Sub SendNewCharClasses(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long
    i = 0
    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).MaleSprite & SEP_CHAR & Class(i).FemaleSprite & SEP_CHAR & Class(i).Locked & SEP_CHAR & Class(i).VIT & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendNewCharClasses", Err.Number, Err.Description
End Sub

Sub SendLeftGame(ByVal index As Long)
On Error GoTo ErrorHandler

Dim Packet As String
    
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & "" & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & "" & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToAllBut(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendLeftGame", Err.Number, Err.Description
End Sub

Sub SendPlayerXY(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendPlayerXY", Err.Number, Err.Description
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
On Error GoTo ErrorHandler

Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateItemToAll", Err.Number, Err.Description
  End
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
On Error GoTo ErrorHandler

Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateItemTo", Err.Number, Err.Description
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)
On Error GoTo ErrorHandler

Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendEditItemTo", Err.Number, Err.Description
  End
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateEmoticonToAll", Err.Number, Err.Description
End Sub

Sub SendUpdateEmoticonTo(ByVal index As Long, ByVal ItemNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateEmoticonTo", Err.Number, Err.Description
End Sub

Sub SendEditEmoticonTo(ByVal index As Long, ByVal EmoNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendEditEmoticonTo", Err.Number, Err.Description
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateNpcToAll", Err.Number, Err.Description
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NpcNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateNpcTo", Err.Number, Err.Description
End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal NpcNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    'Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR
    'Packet = Packet & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & END_CHAR
    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendEditNpcTo", Err.Number, Err.Description
End Sub

Sub SendShops(ByVal index As Long)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim(Shop(i).Name) <> "" Then
            Call SendUpdateShopTo(index, i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendShops", Err.Number, Err.Description
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateShopToAll", Err.Number, Err.Description
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateShopTo", Err.Number, Err.Description
End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendEditShopTo", Err.Number, Err.Description
End Sub

Sub SendSpells(ByVal index As Long)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim(Spell(i).Name) <> "" Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendSpells", Err.Number, Err.Description
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateSpellToAll", Err.Number, Err.Description
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendUpdateSpellTo", Err.Number, Err.Description
End Sub

Sub SendEditSpellTo(ByVal index As Long, ByVal SpellNum As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendEditSpellTo", Err.Number, Err.Description
End Sub

Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
On Error GoTo ErrorHandler

Dim Packet As String
Dim i As Long, x As Long, y As Long, z As Long

    z = 0
    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
        
        ' Item #
        x = Shop(ShopNum).TradeItem(i).GetItem
        
        If Item(x).Type = ITEM_TYPE_SPELL Then
            ' Spell class requirement
            y = Spell(Item(x).Data1).ClassReq
            
            If y = 0 Then
                Call PlayerMsg(index, Trim(Item(x).Name) & " can be used by all classes.", Yellow)
            Else
                Call PlayerMsg(index, Trim(Item(x).Name) & " can only be used by a " & GetClassName(y - 1) & ".", Yellow)
            End If
        End If
        If x < 1 Then
            z = z + 1
        End If
    Next i
    Packet = Packet & END_CHAR
    
    If z = MAX_TRADES Then
        Call PlayerMsg(index, "This shop has nothing to sell!", BrightRed)
    Else
        Call SendDataTo(index, Packet)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendTrade", Err.Number, Err.Description
End Sub

Sub SendPlayerSpells(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendPlayerSpells", Err.Number, Err.Description
End Sub

Sub SendWeatherTo(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
    If RainIntensity <= 0 Then RainIntensity = 1
    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & RainIntensity & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendWeatherTo", Err.Number, Err.Description
End Sub

Sub SendWeatherToAll()
On Error GoTo ErrorHandler
Dim i As Long
Dim Weather As String
    
    Select Case GameWeather
        Case 0
            Weather = "None"
        Case 1
            Weather = "Rain"
        Case 2
            Weather = "Snow"
        Case 3
            Weather = "Thunder"
    End Select
    frmServer.Label5.Caption = "Current Weather: " & Weather
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendWeatherToAll", Err.Number, Err.Description
End Sub

Sub SendTimeTo(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendTimeTo", Err.Number, Err.Description
End Sub

Sub SendTimeToAll()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendTimeToAll", Err.Number, Err.Description
End Sub

