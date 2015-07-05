Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
'On Error GoTo errorhandler:
    frmServer.Caption = "Dual Solace"
    'frmServer.lblIP.Caption = Inet.OpenURL("http://www.whatismyip.org/")
    'frmServer.lblPort.Caption = frmServer.Socket(0).LocalPort
    If ServerState = SERVER_UNLOCKED Then
        frmServer.txtTotal.Text = TotalOnlinePlayers
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "UpdateCaption", Err.Number, Err.Description)
End Sub

Public Function PortInUse(ByVal PortNumber As Integer) As Boolean
Dim oSocket As Object
Dim bAns As Boolean

    On Error Resume Next
    Set oSocket = CreateObject("MSWinsock.Winsock.1")
        
        If Err.Number > 0 Then
            Err.Raise 30000, , "Could not create winsock object"
            Exit Function
        End If
        
    Err.Clear
    
    oSocket.LocalPort = PortNumber
    oSocket.Listen
    
    'if we get this error, it means the port is busy
    bAns = Err.Number = 10048
    oSocket.Close
    Set oSocket = Nothing
    PortInUse = bAns
End Function

Function IsConnected(ByVal index As Long) As Boolean
'On Error GoTo errorhandler:
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modServerTCP.bas", "IsConnected", Err.Number, Err.Description)
End Function

Function IsPlaying(ByVal index As Long) As Boolean
'On Error GoTo errorhandler:
    If IsConnected(index) And Player(index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modServerTCP.bas", "IsPlaying", Err.Number, Err.Description)
End Function

Function IsLoggedIn(ByVal index As Long) As Boolean
'On Error GoTo errorhandler:
    If IsConnected(index) And Trim$(Player(index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modServerTCP.bas", "IsLoggedIn", Err.Number, Err.Description)
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
'On Error GoTo errorhandler:
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase(Trim$(Player(i).Login)) = LCase(Trim$(Login)) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modServerTCP.bas", "IsMultiAccounts", Err.Number, Err.Description)
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
'On Error GoTo errorhandler:
Dim i As Long
Dim n As Long

    n = 0
    IsMultiIPOnline = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And Trim$(GetPlayerIP(i)) = Trim$(IP) Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modServerTCP.bas", "IsMultiIPOnline", Err.Number, Err.Description)
End Function

Function IsBanned(ByVal IP As String) As Boolean
'On Error GoTo errorhandler:
Dim FileName As String, fIP As String, fName As String
Dim f As Long

    IsBanned = False
    
    FileName = App.Path & "\data\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
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
        If Trim$(LCase(fIP)) = Trim$(LCase(Mid(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If
    Loop
    
    Close #f
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modServerTCP.bas", "IsBanned", Err.Number, Err.Description)
End Function

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
'On Error GoTo errorhandler:
Dim i As Long, n As Long, startc As Long

    If IsConnected(index) Then
        frmServer.Socket(index).SendData Data
        DoEvents
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendDataTo", Err.Number, Err.Description)
End Sub

Sub SendDataToAll(ByVal Data As String)
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendDataToAll", Err.Number, Err.Description)
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            Call SendDataTo(i, Data)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendDataToAllBut", Err.Number, Err.Description)
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "SendDataToMap", Err.Number, Err.Description)
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal Data As String)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "SendDataToMapBut", Err.Number, Err.Description)
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "GlobalMsg", Err.Number, Err.Description)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "AdminMsg", Err.Number, Err.Description)
End Sub

Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "PlayerMsg", Err.Number, Err.Description)
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
'On Error GoTo errorhandler:
Dim Packet As String
Dim Text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "MapMsg", Err.Number, Err.Description)
End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "AlertMsg", Err.Number, Err.Description)
End Sub

Sub TrackerMsg(ByVal index As Long, ByVal sType As String, ByVal Msg As String) 'smchronos
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "TRACKERUPDATE" & SEP_CHAR & sType & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "TrackerMsg", Err.Number, Err.Description)
End Sub

Sub ATM(ByVal sType As String, ByVal Msg As String) 'smchronos
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "TRACKERUPDATE" & SEP_CHAR & sType & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "ATM", Err.Number, Err.Description)
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
'On Error GoTo errorhandler:
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "HackingAttempt", Err.Number, Err.Description)
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "AcceptConnection", Err.Number, Err.Description)
End Sub

Sub SocketConnected(ByVal index As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "SocketConnected", Err.Number, Err.Description)
End Sub

' Converts the binary to a string
Public Function CBuf(bData() As Byte, ByVal Total As Integer) As String
'On Error GoTo errorhandler:
Dim n As Long
CBuf = ""

For n = 0 To Total - 1
CBuf = CBuf & ChrW(bData(n))
Next n
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modServerTCP.bas", "CBuf", Err.Number, Err.Description)
End Function

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
'On Error GoTo errorhandler:
On Error Resume Next

Dim Packet As String
Dim Top As String * 3
Dim Start As Integer
Dim StringData As String
Dim ByteData() As Byte

    If index > 0 Then
        frmServer.Socket(index).GetData ByteData, vbByte + vbArray, DataLength
        ' convert the byte buffer
        StringData = CBuf(ByteData(), DataLength)
        
        Dim i As Long
        For i = 0 To DataLength - 1
            Debug.Print "BYTEARRAY(" & i & "): " & ByteData(i)
        Next i
        Debug.Print "STRINGPACKET: " & StringData
        
        If StringData = "top" Then
            Top = STR(TotalOnlinePlayers)
            Call SendDataTo(index, Top)
            Call CloseSocket(index)
        End If
            
        Player(index).Buffer = Player(index).Buffer & StringData
        
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
        If Player(index).DataBytes > 1000 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Data Flooding")
            Exit Sub
        End If
        
        ' Check for packet flooding
        If Player(index).DataPackets > 25 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Packet Flooding")
            Exit Sub
        End If
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "IncomingData", Err.Number, Err.Description)
End Sub

Sub CloseSocket(ByVal index As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "CloseSocket", Err.Number, Err.Description)
End Sub

Sub SendWhosOnline(ByVal index As Long)
'On Error GoTo errorhandler:
Dim s As String
Dim n As Long, i As Long

    s = ""
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            s = s & GetPlayerName(i) & ", "
            n = n + 1
        End If
    Next i
            
    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If
        
    Call PlayerMsg(index, s, WhoColor)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendWhosOnline", Err.Number, Err.Description)
End Sub

Sub SendPlayerList(ByVal index As Long)
'On Error GoTo errorhandler:
Dim s As String
Dim n As Long, i As Long
Dim Packet As String
Dim NamePacket As String

    Packet = "PLIST"
    s = ""
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            NamePacket = NamePacket & SEP_CHAR & GetPlayerName(i)
            n = n + 1
        End If
    Next i
            
    Packet = Packet & SEP_CHAR & n & NamePacket & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendPlayerList", Err.Number, Err.Description)
End Sub

Sub SendPlayerLeave(ByVal index As Long)
'On Error GoTo errorhandler:
Dim s As String
Dim n As Long, i As Long
Dim Packet As String
Dim NamePacket As String

    Packet = "PLAYERLEAVE"
    s = ""
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            NamePacket = NamePacket & SEP_CHAR & GetPlayerName(i)
            n = n + 1
        End If
    Next i
            
    Packet = Packet & SEP_CHAR & n & NamePacket & SEP_CHAR & END_CHAR
    
    Call SendDataToAllBut(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendPlayerLeave", Err.Number, Err.Description)
End Sub

Sub SendChars(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim$(Player(index).Char(i).Name) & SEP_CHAR & Trim$(Class(Player(index).Char(i).Class).Name) & SEP_CHAR & Player(index).Char(i).Level & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendChars", Err.Number, Err.Description)
End Sub

Sub SendJoinMap(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String, FileName As String
Dim i As Long

    Packet = ""
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index And GetPlayerMap(i) = GetPlayerMap(index) Then
            Packet = Packet & "PLAYERDATA" & SEP_CHAR & i & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & GetPlayerSprite(i) & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & GetPlayerX(i) & SEP_CHAR & GetPlayerY(i) & SEP_CHAR & GetPlayerDir(i) & SEP_CHAR & GetPlayerAccess(i) & SEP_CHAR & GetPlayerPK(i) & SEP_CHAR & END_CHAR
            Call SendDataTo(index, Packet)
        End If
    Next i
    
    'Run Scripts
    For i = 0 To frmLibrary.lstLibrary.ListCount - 1
        FileName = App.Path & "\Library\" & frmLibrary.lstLibrary.List(i)
        If GetVar(FileName, "DATA", "Enabled") = "True" Then
            MyScript.ExecuteStatement frmLibrary.lstLibrary.List(i), "JoinMap " & index
        End If
    Next i
    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendJoinMap", Err.Number, Err.Description)
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String, FileName As String
Dim i As Long

    'Run Scripts
    For i = 0 To frmLibrary.lstLibrary.ListCount - 1
        FileName = App.Path & "\Library\" & frmLibrary.lstLibrary.List(i)
        If GetVar(FileName, "DATA", "Enabled") = "True" Then
            MyScript.ExecuteStatement frmLibrary.lstLibrary.List(i), "LeaveMap " & index
        End If
    Next i
    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & END_CHAR
    Call SendDataToMapBut(index, MapNum, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendLeaveMap", Err.Number, Err.Description)
End Sub

Sub SendPlayerData(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    ' Send index's player data to everyone including himself on the map
    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendPlayerData", Err.Number, Err.Description)
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Shop & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(MapNum).Tile(x, y)
                Packet = Packet & .Ground & SEP_CHAR
                Packet = Packet & .Mask & SEP_CHAR
                Packet = Packet & .Mask2 & SEP_CHAR
                Packet = Packet & .Anim & SEP_CHAR
                Packet = Packet & .Anim2 & SEP_CHAR
                Packet = Packet & .Fringe & SEP_CHAR
                Packet = Packet & .FringeAnim & SEP_CHAR
                Packet = Packet & .Fringe2 & SEP_CHAR
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
    Call SendMapAttributes(index, MapNum)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendMap", Err.Number, Err.Description)
End Sub

Sub SendMapAttributes(ByVal index As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim x As Long
Dim y As Long

    Packet = "MAPATTRIBUTES" & SEP_CHAR & MapNum & SEP_CHAR
    Call TextAdd(frmServer.txtText, "Sending Attributes for Map #" & MapNum & " to " & GetPlayerName(index) & ".", True)
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(MapNum).Tile(x, y)
                Packet = Packet & .Walkable & SEP_CHAR
                Packet = Packet & .Blocked & SEP_CHAR
                Packet = Packet & .Warp & SEP_CHAR & .WarpMap & SEP_CHAR & .WarpX & SEP_CHAR & .WarpY & SEP_CHAR
                Packet = Packet & .Item & SEP_CHAR & .ItemNum & SEP_CHAR & .ItemValue & SEP_CHAR
                Packet = Packet & .NpcAvoid & SEP_CHAR
                Packet = Packet & .Key & SEP_CHAR & .KeyNum & SEP_CHAR & .KeyTake & SEP_CHAR
                Packet = Packet & .KeyOpen & SEP_CHAR & .KeyOpenX & SEP_CHAR & .KeyOpenY & SEP_CHAR
                Packet = Packet & .North & SEP_CHAR
                Packet = Packet & .West & SEP_CHAR
                Packet = Packet & .East & SEP_CHAR
                Packet = Packet & .South & SEP_CHAR
                Packet = Packet & .Shop & SEP_CHAR & .ShopNum & SEP_CHAR
                Packet = Packet & .Bank & SEP_CHAR
                Packet = Packet & .Heal & SEP_CHAR & .HealValue & SEP_CHAR
                Packet = Packet & .Damage & SEP_CHAR & .DamageValue & SEP_CHAR
            End With
        Next x
    Next y

    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)

ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendMapAttributes", Err.Number, Err.Description)
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "SendMapItemsTo", Err.Number, Err.Description)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "SendMapItemsToAll", Err.Number, Err.Description)
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "SendMapNpcsTo", Err.Number, Err.Description)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "SendMapNpcsToMap", Err.Number, Err.Description)
End Sub

Sub SendItems(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(index, i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendItems", Err.Number, Err.Description)
End Sub

Sub SendNpcs(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim$(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendNpcs", Err.Number, Err.Description)
End Sub

Sub SendBankInv(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long

    Packet = "PLAYERBANK" & SEP_CHAR
    For i = 1 To MAX_BANK_ITEMS
        Packet = Packet & GetPlayerBankItemNum(index, i) & SEP_CHAR & GetPlayerBankItemValue(index, i) & SEP_CHAR & GetPlayerBankItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendBankInv", Err.Number, Err.Description)
End Sub

Sub SendInventory(ByVal index As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "SendInventory", Err.Number, Err.Description)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
'On Error GoTo errorhandler:
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendInventoryUpdate", Err.Number, Err.Description)
End Sub

Sub SendWornEquipment(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
    
    Packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    'Call SendEquipDataTo(Index)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendWornEquipment", Err.Number, Err.Description)
End Sub

Sub SendHP(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendHP", Err.Number, Err.Description)
End Sub

Sub SendMP(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendMP", Err.Number, Err.Description)
End Sub

Sub SendSP(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendSP", Err.Number, Err.Description)
End Sub

Sub SendStats(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
    
        Packet = "PLAYERSTATS" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & (GetPlayerNextLevel(index) - GetPlayerExp(index)) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendStats", Err.Number, Err.Description)
End Sub

Sub SendPlayers(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim n As Integer
    Packet = "PLAYERPLAYERS" & SEP_CHAR
    
    For n = 1 To MAX_PLAYERS
        If IsPlaying(n) Then
            Packet = Packet & GetPlayerName(n) & SEP_CHAR
        End If
    Next n
    
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendPlayers", Err.Number, Err.Description)
End Sub

Sub SendFriends(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim n As Integer
    Packet = "PLAYERFRIENDS" & SEP_CHAR
    
    For n = 1 To MAX_FRIENDS
        If Player(index).Char(Player(index).CharNum).Friends(n) <> vbNullString Then
            Packet = Packet & Player(index).Char(Player(index).CharNum).Friends(n) & SEP_CHAR
        End If
    Next n
    
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendFriends", Err.Number, Err.Description)
End Sub

Sub SendWelcome(ByVal index As Long)
'On Error GoTo errorhandler:
'Dim MOTD As String
'Dim f As Long
    
    ' Send them welcome
    'Call PlayerMsg(Index, "Welcome to " & GAME_NAME & "!  Programmed from scratch by yours truely Consty!  Version " & CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION, BrightBlue)
    'Call PlayerMsg(Index, "Type /help for help on commands.  Use arrow keys to move, hold down shift to run, and use ctrl to attack.", Cyan)
    ' Send them MOTD
    'MOTD = GetVar(App.Path & "\data\data.ini", "Strings", "MOTD")
    'If Trim$(MOTD) <> "" Then
    '    Call PlayerMsg(Index, "MOTD: " & MOTD, BrightCyan)
    'End If
    
    ' Send whos online
    'Call SendWhosOnline(Index)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendWelcome", Err.Number, Err.Description)
End Sub

Sub SendClasses(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendClasses", Err.Number, Err.Description)
End Sub

Sub SendClassesToAll()
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendClasses", Err.Number, Err.Description)
End Sub

Sub SendNewCharClasses(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR & Max_Visible_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).Sprite & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendNewCharClasses", Err.Number, Err.Description)
End Sub

Sub SendClassData(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long

    Packet = "CLASSDATA" & SEP_CHAR & Max_Classes & SEP_CHAR & Max_Visible_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).Sprite & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendClassData", Err.Number, Err.Description)
End Sub

Sub SendLeftGame(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & "" & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR
    Call SendDataToAllBut(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendLeftGame", Err.Number, Err.Description)
End Sub

Sub SendPlayerXY(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendPlayerXY", Err.Number, Err.Description)
End Sub

Sub SendStrings(ByVal index As Long)
Dim Packet As String

    Packet = "STRINGS" & SEP_CHAR & STRING_STRENGTH & SEP_CHAR & STRING_DEFENSE & SEP_CHAR & STRING_MAGIC & SEP_CHAR & STRING_SPEED & SEP_CHAR & STRING_HP & SEP_CHAR & STRING_MP & SEP_CHAR & STRING_SP & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendStrings", Err.Number, Err.Description)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Trim$(Item(ItemNum).Description) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).Data4 & SEP_CHAR & Item(ItemNum).Data5 & SEP_CHAR & Item(ItemNum).Sound & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateItemToAll", Err.Number, Err.Description)
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Trim$(Item(ItemNum).Description) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).Data4 & SEP_CHAR & Item(ItemNum).Data5 & SEP_CHAR & Item(ItemNum).Sound & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateItemTo", Err.Number, Err.Description)
End Sub

Sub SendUpdateBankItemTo(ByVal index As Long, ByVal ItemSlot As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "UPDATEBANKITEM" & SEP_CHAR
    Packet = Packet & ItemSlot & SEP_CHAR & GetPlayerBankItemNum(index, ItemSlot) & SEP_CHAR & GetPlayerBankItemValue(index, ItemSlot) & SEP_CHAR & GetPlayerBankItemDur(index, ItemSlot) & SEP_CHAR
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateBankItemTo", Err.Number, Err.Description)
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Trim$(Item(ItemNum).Description) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).Data4 & SEP_CHAR & Item(ItemNum).Data5 & SEP_CHAR & Item(ItemNum).Sound & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendEditItemTo", Err.Number, Err.Description)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).TintR & SEP_CHAR & Npc(NpcNum).TintG & SEP_CHAR & Npc(NpcNum).TintB & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateNpcToAll", Err.Number, Err.Description)
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NpcNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).TintR & SEP_CHAR & Npc(NpcNum).TintG & SEP_CHAR & Npc(NpcNum).TintB & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateNpcTo", Err.Number, Err.Description)
End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal NpcNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).HP & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).EXP & SEP_CHAR & Npc(NpcNum).Fear & SEP_CHAR & Npc(NpcNum).TintR & SEP_CHAR & Npc(NpcNum).TintG & SEP_CHAR & Npc(NpcNum).TintB & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendEditNpcTo", Err.Number, Err.Description)
End Sub

Sub SendShops(ByVal index As Long)
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim$(Shop(i).Name) <> "" Then
            Call SendUpdateShopTo(index, i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendShops", Err.Number, Err.Description)
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateShopToAll", Err.Number, Err.Description)
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateShopTo", Err.Number, Err.Description)
End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).Restock & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).MaxStock & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendEditShopTo", Err.Number, Err.Description)
End Sub

Sub SendSpells(ByVal index As Long)
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).Name) <> "" Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendSpells", Err.Number, Err.Description)
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateSpellToAll", Err.Number, Err.Description)
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateSpellTo", Err.Number, Err.Description)
End Sub

Sub SendEditSpellTo(ByVal index As Long, ByVal SpellNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendEditSpellTo", Err.Number, Err.Description)
End Sub

Sub SendUpdateClassToAll(ByVal ClassNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    With Class(ClassNum)
        Packet = "UPDATECLASS" & SEP_CHAR
        Packet = Packet & ClassNum & SEP_CHAR
        Packet = Packet & Trim$(.Name) & SEP_CHAR
        Packet = Packet & .Sprite & SEP_CHAR
        Packet = Packet & .HP & SEP_CHAR
        Packet = Packet & .MP & SEP_CHAR
        Packet = Packet & .SP & SEP_CHAR
        Packet = Packet & .STR & SEP_CHAR
        Packet = Packet & .DEF & SEP_CHAR
        Packet = Packet & .MAGI & SEP_CHAR
        Packet = Packet & .SPEED & SEP_CHAR
        Packet = Packet & .Map & SEP_CHAR
        Packet = Packet & .x & SEP_CHAR
        Packet = Packet & .y & SEP_CHAR
        Packet = Packet & END_CHAR
    End With
    
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateClassToAll", Err.Number, Err.Description)
End Sub

Sub SendUpdateClassTo(ByVal index As Long, ByVal ClassNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    With Class(ClassNum)
        Packet = "UPDATECLASS" & SEP_CHAR
        Packet = Packet & ClassNum & SEP_CHAR
        Packet = Packet & Trim$(.Name) & SEP_CHAR
        Packet = Packet & .Sprite & SEP_CHAR
        Packet = Packet & .HP & SEP_CHAR
        Packet = Packet & .MP & SEP_CHAR
        Packet = Packet & .SP & SEP_CHAR
        Packet = Packet & .STR & SEP_CHAR
        Packet = Packet & .DEF & SEP_CHAR
        Packet = Packet & .MAGI & SEP_CHAR
        Packet = Packet & .SPEED & SEP_CHAR
        Packet = Packet & .Map & SEP_CHAR
        Packet = Packet & .x & SEP_CHAR
        Packet = Packet & .y & SEP_CHAR
        Packet = Packet & END_CHAR
    End With
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendUpdateClassTo", Err.Number, Err.Description)
End Sub

Sub SendEditClassTo(ByVal index As Long, ByVal ClassNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    With Class(ClassNum)
        Packet = "EDITCLASS" & SEP_CHAR
        Packet = Packet & ClassNum & SEP_CHAR
        Packet = Packet & Trim$(.Name) & SEP_CHAR
        Packet = Packet & .Sprite & SEP_CHAR
        Packet = Packet & .HP & SEP_CHAR
        Packet = Packet & .MP & SEP_CHAR
        Packet = Packet & .SP & SEP_CHAR
        Packet = Packet & .STR & SEP_CHAR
        Packet = Packet & .DEF & SEP_CHAR
        Packet = Packet & .MAGI & SEP_CHAR
        Packet = Packet & .SPEED & SEP_CHAR
        Packet = Packet & .Map & SEP_CHAR
        Packet = Packet & .x & SEP_CHAR
        Packet = Packet & .y & SEP_CHAR
        Packet = Packet & END_CHAR
    End With
    Debug.Print ClassNum
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendEditClassTo", Err.Number, Err.Description)
End Sub

Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
'On Error GoTo errorhandler:
Dim Packet As String
Dim i As Long, x As Long, y As Long

    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    If Shop(ShopNum).Restock = TIME_MINUTE Then
        Packet = Packet & CStr(60 - Server_Second & " Second(s)") & SEP_CHAR
    ElseIf Shop(ShopNum).Restock = TIME_HOUR Then
        Packet = Packet & CStr((60 - Server_Minute) & " Minute(s) and " & (60 - Server_Second) & " Second(s)") & SEP_CHAR
    ElseIf Shop(ShopNum).Restock = TIME_FULL Then
        Packet = Packet & CStr((24 - Server_Hour) & " Hour(s) and " & (60 - Server_Minute) & " Minute(s)") & SEP_CHAR
    End If
    
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Stock & SEP_CHAR
        
        ' Item #
        x = Shop(ShopNum).TradeItem(i).GetItem
        
        If Item(x).Type = ITEM_TYPE_SPELL Then
            ' Spell class requirement
            y = Spell(Item(x).Data1).ClassReq
            
            If y = 0 Then
                Call PlayerMsg(index, Trim$(Item(x).Name) & " can be used by all classes.", Yellow)
            Else
                Call PlayerMsg(index, Trim$(Item(x).Name) & " can only be used by a " & GetClassName(y - 1) & ".", Yellow)
            End If
        End If
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendTrade", Err.Number, Err.Description)
End Sub

Sub SendPlayerSpells(ByVal index As Long)
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("modServerTCP.bas", "SendPlayerSpells", Err.Number, Err.Description)
End Sub

Sub SendWeatherTo(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendWeatherTo", Err.Number, Err.Description)
End Sub

Sub SendWeatherToAll()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendWeatherToAll", Err.Number, Err.Description)
End Sub

Sub SendTimeTo(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendTimeTo", Err.Number, Err.Description)
End Sub

Sub SendTimeToAll()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendTimeToAll", Err.Number, Err.Description)
End Sub

Sub SendEquipDataTo(ByVal index As Long)
'On Error GoTo errorhandler:
Dim Packet As String

Packet = "equipdata" & SEP_CHAR
'check weapon slot and others as well
If GetPlayerWeaponSlot(index) > 0 Then
    Packet = Packet & GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) & SEP_CHAR & GetPlayerInvItemDur(index, GetPlayerWeaponSlot(index)) & SEP_CHAR
Else
    Packet = Packet & 0 & SEP_CHAR & 0 & SEP_CHAR
End If

If GetPlayerArmorSlot(index) > 0 Then
    Packet = Packet & GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) & SEP_CHAR & GetPlayerInvItemDur(index, GetPlayerArmorSlot(index)) & SEP_CHAR
Else
    Packet = Packet & 0 & SEP_CHAR & 0 & SEP_CHAR
End If

If GetPlayerHelmetSlot(index) > 0 Then
    Packet = Packet & GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) & SEP_CHAR & GetPlayerInvItemDur(index, GetPlayerHelmetSlot(index)) & SEP_CHAR
Else
    Packet = Packet & 0 & SEP_CHAR & 0 & SEP_CHAR
End If

If GetPlayerShieldSlot(index) > 0 Then
    Packet = Packet & GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) & SEP_CHAR & GetPlayerInvItemDur(index, GetPlayerShieldSlot(index)) & SEP_CHAR
Else
    Packet = Packet & 0 & SEP_CHAR & 0 & SEP_CHAR
End If

Packet = Packet & END_CHAR

Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modServerTCP.bas", "SendEquipDataTo", Err.Number, Err.Description)
End Sub
