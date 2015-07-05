Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateLabel()
    frmServer.Caption = "Endieko Online"
    'frmServer.lblIpAddress = "IP Address: " & frmServer.Socket(0).LocalIP
    frmServer.lblPlayersOnline.Caption = "Current Players Online: " & TotalOnlinePlayers
    'frmServer.lblStatus.Caption = "Status: " & serverStatus
    'frmServer.lblPort.Caption = "Port: " & frmServer.Socket(0).LocalPort
End Sub

'Function IsConnected(ByVal index As Long) As Boolean
'    If frmServer.Socket(index).State = sckConnected Then
'        IsConnected = True
'    Else
'        IsConnected = False
'    End If
'End Function

Function IsConnected(ByVal Index As Long) As Boolean
    IsConnected = False
    If Index = 0 Then Exit Function
    If GameServer Is Nothing Then Exit Function
    If Not GameServer.Sockets(Index).Socket Is Nothing Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) And Player(Index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If IsConnected(Index) And Trim$(Player(Index).Login) <> vbNullString Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase$(Trim$(Player(i).Login)) = LCase$(Trim$(Login)) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next i
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
Dim i As Long
Dim n As Long

    n = 0
    IsMultiIPOnline = False
    For i = 1 To MAX_PLAYERS
'        If IsConnected(i) And Trim$(GetPlayerIP(i)) = Trim$(IP) Then --- Pre-IOCP
'            n = n + 1
'
'            If (n > 1) Then
'                IsMultiIPOnline = True
'                Exit Function
'            End If
'        End If
        If IsConnected(i) Then
             If Trim(GetPlayerIP(i)) = Trim(IP) Then
                 n = n + 1
            
                 If (n > 1) Then
                     IsMultiIPOnline = True
                     Exit Function
                 End If
             End If
        End If
    Next i
End Function

'Function IsBanned(ByVal IP As String) As Boolean
'Dim FileName As String, fIP As String, fName As String
'Dim f As Long
'
'    IsBanned = False
'
'    FileName = App.Path & "\banlist.txt"
'
'    ' Check if file exists
'    If Not FileExist("banlist.txt") Then
'        f = FreeFile
'        Open FileName For Output As #f
'        Close #f
'    End If
'
'    f = FreeFile
'    Open FileName For Input As #f
'
'    Do While Not EOF(f)
'        Input #f, fIP
'        Input #f, fName
'
'        ' Is banned?
'        If Trim$(Lcase$(fIP)) = Trim$(Lcase$(Mid$(IP, 1, Len(fIP)))) Then
'            IsBanned = True
'            Close #f
'            Exit Function
'        End If
'    Loop
'
'    Close #f
'End Function

Function IsBanned(ByVal IP As String) As Boolean

Dim FileName As String, fIP As String, fName As String
Dim f As Long
'Dim b As Integer
Dim BIp As String
Dim i As Integer

    IsBanned = False
   
    FileName = App.Path & "\banlist.ini"
   
    For i = 0 To MAX_BANS
        If Ban(i).BannedIP <> vbNullString Then
             BIp = Ban(i).BannedIP
             If IP = BIp Then
                 IsBanned = True
                 Exit Function
             Else
                 IsBanned = False
             End If
        End If
    Next i
   
End Function

Function IsBannedHD(ByVal HD As String) As Boolean
Dim FileName As String
Dim bHD As String
Dim i As Integer

    IsBannedHD = False
   
    FileName = App.Path & "\banlist.ini"
   
    For i = 0 To MAX_BANS
        If Ban(i).BannedHD <> vbNullString Then
             bHD = Ban(i).BannedHD
             If HD = bHD Then
                 IsBannedHD = True
                 Exit Function
             Else
                 IsBannedHD = False
             End If
        End If
    Next i
   
End Function

'Sub SendDataTo(ByVal Index As Long, ByVal Data As String) -- Pre-IOCP
'Dim i As Long, n As Long, startc As Long
'
'    If IsConnected(Index) Then
'        frmServer.Socket(Index).SendData Data
'        DoEvents
'    End If
'End Sub

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
Dim i As Long, n As Long, startc As Long
'Dim dbytes() As Byte
'
'    dbytes = StrConv(Data, vbFromUnicode)
'    If IsConnected(Index) Then
'        GameServer.Sockets(Index).WriteBytes dbytes
'        DoEvents
'    End If

    If IsConnected(Index) Then
        With ConQueues(Index)
          .Lock = True
          .Lines = .Lines & Data
          .Lock = False
        End With
        'DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> Index Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendQueuedData()
Dim i As Integer, n As Long
Dim TmpStr As String
Dim dbytes() As Byte
Dim Sploc As Integer
Dim ECloc As Integer
Dim lR As Long

    For i = 1 To MAX_PLAYERS
      TmpStr = ""
      With ConQueues(i)
        If Not .Lock Then
          If GameServer.Sockets(i).Socket Is Nothing Then
            .Lines = ""
          End If
          If Len(.Lines) = 0 And QueueDisconnect(i) = True Then
            Call CloseSocket(i)
            QueueDisconnect(i) = False
          Else
            If Len(.Lines) > 0 Then
               If Len(.Lines) < MAX_PACKETLEN Then
                 TmpStr = .Lines
               Else
                 TmpStr = Left(.Lines, MAX_PACKETLEN)
               End If
               .Lines = Right(.Lines, Len(.Lines) - Len(TmpStr))
            End If
          End If
          If Len(TmpStr) > 0 Then
            'Call frmServer.Socket(i).SendData(TmpStr)
            
            If IsConnected(i) Then
                TmpStr = Compress(TmpStr, lR)
                TmpStr = lR & SEP_CHAR & TmpStr
                dbytes = StrConv(TmpStr, vbFromUnicode)
                GameServer.Sockets(i).WriteBytes dbytes
                DoEvents
            End If
          End If
        End If
      End With
    Next
    DoEvents
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If
    Next i
End Sub

Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim Text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
    QueueDisconnect(Index) = True
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub AcceptConnection(Socket As JBSOCKETSERVERLib.ISocket)
Dim i As Long

    i = FindOpenPlayerSlot
   
    If i <> 0 Then
        'Whoho, we can connect them
        Socket.UserData = i
        Set GameServer.Sockets(CStr(i)).Socket = Socket
        Call SocketConnected(i)
        Socket.RequestRead
    Else
        Socket.Close
    End If
End Sub

'Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long) --- Pre-IOCP
'Dim i As Long
'
'    If (Index = 0) Then
'        i = FindOpenPlayerSlot
'
'        If i <> 0 Then
'            ' Whoho, we can connect them
'            frmServer.Socket(i).Close
'            frmServer.Socket(i).Accept SocketId
'            Call SocketConnected(i)
'        End If
'    End If
'End Sub

Sub SocketConnected(ByVal Index As Long)
    If Index <> 0 Then
        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(Index)) Then
                Call TextAdd(frmServer.txtText, "Received connection from " & GetPlayerIP(Index) & ".", True)
            Else
                Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
            End If
        'Else
           ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    End If
End Sub

'Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long) ---- Pre-IOCP
'Dim Buffer As String
'Dim Packet As String
'Dim top As String * 3
'Dim Start As Long
'
'    If Index > 0 Then
'        frmServer.Socket(Index).GetData Buffer, vbString, DataLength
'
'        If Buffer = "top" Then
'            top = STR(TotalOnlinePlayers)
'            Call SendDataTo(Index, top)
'            Call CloseSocket(Index)
'        End If
'
'        Player(Index).Buffer = Player(Index).Buffer & Buffer
'
'        Start = InStr(Player(Index).Buffer, END_CHAR)
'        Do While Start > 0
'            Packet = Mid$(Player(Index).Buffer, 1, Start - 1)
'            Player(Index).Buffer = Mid$(Player(Index).Buffer, Start + 1, Len(Player(Index).Buffer))
'            Player(Index).DataPackets = Player(Index).DataPackets + 1
'            Start = InStr(Player(Index).Buffer, END_CHAR)
'            If Len(Packet) > 0 Then
'                Call HandleData(Index, Packet)
'            End If
'        Loop
'
'        ' Check if elapsed time has passed
'        Player(Index).DataBytes = Player(Index).DataBytes + DataLength
'        If GetTickCount >= Player(Index).DataTimer + 1000 Then
'            Player(Index).DataTimer = GetTickCount
'            Player(Index).DataBytes = 0
'            Player(Index).DataPackets = 0
'            Exit Sub
'        End If
'
'        ' Check for data flooding
'        If Player(Index).DataBytes > 1000 And GetPlayerAccess(Index) <= 0 Then
'            Call HackingAttempt(Index, "Data Flooding")
'            Exit Sub
'        End If
'
'        ' Check for packet flooding
'        If Player(Index).DataPackets > 25 And GetPlayerAccess(Index) <= 0 Then
'            Call HackingAttempt(Index, "Packet Flooding")
'            Exit Sub
'        End If
'    End If
'End Sub

Sub IncomingData(Socket As JBSOCKETSERVERLib.ISocket, Data As JBSOCKETSERVERLib.IData)
On Error Resume Next

Dim Buffer As String
Dim dbytes() As Byte
Dim Packet As String
Dim top As String * 3
Dim Start As Integer
Dim Index As Long
Dim DataLength As Long
Dim lR As Long
Dim Sploc As Integer

    dbytes = Data.Read
    Socket.RequestRead
    Buffer = StrConv(dbytes(), vbUnicode)
    DataLength = Len(Buffer)
    Index = CLng(Socket.UserData)
    
    Sploc = InStr(1, Buffer, SEP_CHAR)
             lR = Mid(Buffer, 1, Sploc - 1)
             Buffer = Mid(Buffer, Sploc + 1, Len(Buffer) - Sploc)
             'Debug.Print lR & vbCrLf & "Parse(1):" & vbCrLf & Buffer & vbCrLf & "Buffer:"
             Buffer = Uncompress(Buffer, lR)
    If Buffer = "top" Then
        top = STR$(TotalOnlinePlayers)
        Call SendDataTo(Index, top)
        QueueDisconnect(Index) = True
    End If
            
    Player(Index).Buffer = Player(Index).Buffer & Buffer
       
    Start = InStr(Player(Index).Buffer, END_CHAR)
    Do While Start > 0
        Packet = Mid$(Player(Index).Buffer, 1, Start - 1)
        Player(Index).Buffer = Mid$(Player(Index).Buffer, Start + 1, Len(Player(Index).Buffer))
        Player(Index).DataPackets = Player(Index).DataPackets + 1
        Start = InStr(Player(Index).Buffer, END_CHAR)
        If Len(Packet) > 0 Then
             Call HandleData(Index, Packet)
        End If
    Loop
                
    ' Check if elapsed time has passed
    Player(Index).DataBytes = Player(Index).DataBytes + DataLength
    If GetTickCount >= Player(Index).DataTimer + 1000 Then
        Player(Index).DataTimer = GetTickCount
        Player(Index).DataBytes = 0
        Player(Index).DataPackets = 0
        Exit Sub
    End If
       
    ' Check for data flooding
    If Player(Index).DataBytes > 1000 And GetPlayerAccess(Index) <= 0 Then
        Call HackingAttempt(Index, "Data Flooding")
        Exit Sub
    End If
       
    ' Check for packet flooding
    If Player(Index).DataPackets > 25 And GetPlayerAccess(Index) <= 0 Then
        Call HackingAttempt(Index, "Packet Flooding")
        Exit Sub
    End If
End Sub

Sub HandleData(ByVal Index As Long, ByVal Data As String)
Dim Parse() As String
Dim Name As String
Dim Password As String
Dim Sex As Long
Dim Class As Long
Dim CharNum As Long
Dim Msg As String
Dim IPMask As String
Dim BanSlot As Long
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim BankNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Movement As Long
Dim i As Long, n As Long, x As Long, y As Long, f As Long
Dim MapNum As Long
Dim s As String
Dim tMapStart As Long, tMapEnd As Long
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long
Dim z As Long
Dim BIp As Integer
Dim Reason As String
        
    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
Select Case LCase$(Parse$(0))
    '::::::::::::::::::::::::::::
    ':: Warpem By Click Packet ::
    '::::::::::::::::::::::::::::
    Case "playerjump"
    'If LCase$(Parse$(0)) = "playerjump" Then
    
        ' Prevent Hacking
        If GetPlayerAccess(Index) < 1 Then
            Call HackingAttempt(Index, "Attempted Admin Jump.")
            Exit Sub
        End If
        
        ' Warp Em
        Call SetPlayerX(Index, Val(Parse$(1)))
        Call SetPlayerY(Index, Val(Parse$(2)))
        Exit Sub
    'End If
        
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Requesting classes for making a character ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    Case "getclasses"
    'If LCase$(Parse$(0)) = "getclasses" Then
        If Not IsPlaying(Index) Then
            Call SendNewCharClasses(Index)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::
    ' :: Guilds Packet ::
    ' :::::::::::::::::::
    
    ' Change Access
    Case "guildchangeaccess"
    'If LCase$(Parse$(0)) = "guildchangeaccess" Then
        ' Check the requirements.
        If FindPlayer(Parse$(1)) = 0 Then
            Call PlayerMsg(Index, "Player is offline", White)
            Exit Sub
        End If
        
        If FindPlayer(Parse$(1)) = Index Then
            Call PlayerMsg(Index, "You cant change your guild access!", BrightRed)
            Exit Sub
        End If
        
        If GetPlayerGuild(FindPlayer(Parse$(1))) <> GetPlayerGuild(Index) Then
            Call PlayerMsg(Index, "Player is not in your guild", BrightRed)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(Index) < 3 Then
            Call PlayerMsg(Index, "You need to be a higher access to do this!", BrightRed)
            Exit Sub
        End If
        
        'Set the player's new access level
        If Val(Parse$(2)) < 0 Then Parse$(2) = 0
        If Val(Parse$(2)) > 4 Then Parse$(2) = 4
        
        If Val(Parse$(2)) > GetPlayerGuildAccess(Index) Then
            Call PlayerMsg(Index, "You cant set a someones guild access higher than your own!", BrightRed)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(Index) <= GetPlayerGuildAccess(FindPlayer(Parse$(1))) Then
            Call PlayerMsg(Index, "You cant change " & GetPlayerName(FindPlayer(Parse$(1))) & "'s guild access!", BrightRed)
            Exit Sub
        End If
        
        Call SetPlayerGuildAccess(FindPlayer(Parse$(1)), Val(Parse$(2)))
        Call SendPlayerData(FindPlayer(Parse$(1)))
        Call PlayerMsg(FindPlayer(Parse$(1)), "Your guild access has been changed to " & Val(Parse$(2)) & "!", Yellow)
        Call PlayerMsg(Index, "You changed " & GetPlayerName(FindPlayer(Parse$(1))) & "'s guild access to " & Val(Parse$(2)) & "!", Yellow)
        'Can have a message here if you'd like
        Exit Sub
    'End If
    
    ' Disown--- Same as kick
    Case "guilddisown"
    'If LCase$(Parse$(0)) = "guilddisown" Then
        ' Check if all the requirements
            If FindPlayer(Parse$(1)) = 0 Then
            Call PlayerMsg(Index, "Player is offline", White)
            Exit Sub
        End If
        If GetPlayerGuild(FindPlayer(Parse$(1))) <> GetPlayerGuild(Index) Then
            Call PlayerMsg(Index, "Player is not in your guild", Red)
            Exit Sub
        End If
        If GetPlayerGuildAccess(FindPlayer(Parse$(1))) > GetPlayerGuildAccess(Index) Then
            Call PlayerMsg(Index, "Player has a higher guild level than you.", Red)
            Exit Sub
        End If
        'Player checks out, take him out of the guild
        Call SetPlayerGuild(FindPlayer(Parse$(1)), vbNullString)
        Call SetPlayerGuildAccess(FindPlayer(Parse$(1)), 0)
        Call SendPlayerData(FindPlayer(Parse$(1)))
        'Can have a message here if you'd like
        Exit Sub
    'End If

    ' Leave Guild
    Case "guildleave"
    'If LCase$(Parse$(0)) = "guildleave" Then
        ' Check if they can leave
        If GetPlayerGuild(Index) = vbNullString Then
            Call PlayerMsg(Index, "You are not in a guild.", Red)
            Exit Sub
        End If
        Call SetPlayerGuild(Index, vbNullString)
        Call SetPlayerGuildAccess(Index, 0)
        Call SendPlayerData(Index)
        Exit Sub
    'End If
    
    ' Start a new guild
    Case "buyguild"
    'If LCase$(Parse$(0)) = "buyguild" Then
        Call CreateGuild(Index, Parse$(1))
    'End If
    
    ' Make A New Guild ------ The admin Way....
    Case "makeguild"
    'If LCase$(Parse$(0)) = "makeguild" Then
        ' Check if the Owner is Online
        If FindPlayer(Parse$(1)) = 0 Then
            Call PlayerMsg(Index, "Player is offline", White)
            Exit Sub
        End If
        ' Check if they are alredy in a guild
            If GetPlayerGuild(FindPlayer(Parse$(1))) <> vbNullString Then
            Call PlayerMsg(Index, "Player is already in a guild", Red)
            Exit Sub
        End If
        ' If everything is ok then lets make the guild
        Call SetPlayerGuild(FindPlayer(Parse$(1)), (Parse$(2)))
        Call SetPlayerGuildAccess(FindPlayer(Parse$(1)), 3)
        Call SendPlayerData(FindPlayer(Parse$(1)))
        Exit Sub
    'End If
    
    ' kick players from guild
    Case "kickfromguild"
    'If LCase$(Parse$(0)) = "kickfromguild" Then
        n = FindPlayer(Parse$(1))
        
        If n = 0 Then
            Call PlayerMsg(Index, "Player is offline.", White)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(Index) <= 2 Then
            Call PlayerMsg(Index, "You need be be a higher guild access to kick someone!", Red)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(n) >= GetPlayerGuildAccess(Index) Then
            Call PlayerMsg(Index, "You cant kick people with the same or higher guild access then you!", Red)
            Exit Sub
        End If
        
        If Trim(GetPlayerGuild(n)) <> Trim(GetPlayerGuild(Index)) Then
            Call PlayerMsg(n, "The player needs to be in the same guild as you!", Red)
            Exit Sub
        End If
        
        Call PlayerMsg(n, "You have been kicked from the guild " & Trim(GetPlayerGuild(n)) & " !", Red)
        Call PlayerMsg(Index, "You kicked " & Trim(GetPlayerName(n)) & " from the guild!", Red)
        Call SetPlayerGuild(n, vbNullString)
        Call SetPlayerGuildAccess(n, 0)
        Call SendPlayerData(n)
        Exit Sub
    'End If
                
    ' Invite new member
    Case "guildinvite"
    'If LCase$(Parse$(0)) = "guildinvite" Then
        If GetPlayerGuild(Index) <> vbNullString Then
            Call PlayerMsg(Index, "You're already in a guild!", Red)
            Exit Sub
        End If
        If Trim(Player(Index).GuildTemp = vbNullString) Then
            Call PlayerMsg(Index, "No one invited you to a guild!", Red)
            Exit Sub
        End If
        If Val(Parse$(1)) = 0 Then
            Call SetPlayerGuild(Index, Player(Index).GuildTemp)
            Call SetPlayerGuildAccess(Index, 0)
            Call SendPlayerData(Index)
            Call PlayerMsg(Index, "You joined the guild " & Player(Index).GuildTemp & "!", BrightGreen)
            Call PlayerMsg(Player(Index).GuildInviter, GetPlayerName(Index) & " joined your guild!", BrightGreen)
            Player(Index).GuildInvitation = False
            Player(Index).GuildTemp = vbNullString
            Player(Index).GuildInviter = 0
        Else
            Call PlayerMsg(Index, "You you declined the invitation from the guild " & Player(Index).GuildTemp & "!", BrightRed)
            Call PlayerMsg(Player(Index).GuildInviter, GetPlayerName(Index) & " declined your guild invitation.", BrightRed)
            Player(Index).GuildInvitation = False
            Player(Index).GuildTemp = vbNullString
            Player(Index).GuildInviter = 0
        End If
        Exit Sub
    'End If
    
    ' Invite new members
    Case "invitetoguild"
    'If LCase$(Parse$(0)) = "invitetoguild" Then
        i = FindPlayer(Parse$(1))
        
        If i = 0 Then
            Call PlayerMsg(Index, "Player is offline.", White)
            Exit Sub
        End If
        
        If Player(i).GuildInvitation = True Then
            Call PlayerMsg(Index, "Player is already being invited to another guild.", White)
            Exit Sub
        End If
        
        If GetPlayerGuild(i) <> vbNullString Then
            Call PlayerMsg(Index, "Player is already in a guild.", White)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(Index) <= 0 Then
            Call PlayerMsg(Index, "You need to be a higher access to invite players into the guild!", Red)
            Exit Sub
        End If
        
        Player(i).GuildInvitation = True
        Player(i).GuildTemp = GetPlayerGuild(Index)
        Player(i).GuildInviter = Index
        
        Call PlayerMsg(i, "You have been invited to join the guild " & GetPlayerGuild(Index) & ". Do you accept (/guildaccept) or decline (/guilddecline)?", BrightGreen)
        Call PlayerMsg(Index, "You have invited " & Trim(GetPlayerName(i)) & " into your guild.", BrightGreen)
        Exit Sub
    'End If
    
    ' Who from the guild is online
    Case "guildwho"
    'If LCase$(Parse$(0)) = "guildwho" Then
        If Trim$(GetPlayerGuild(Index)) = vbNullString Then
            Call PlayerMsg(Index, "Your are not in a guild!", BrightRed)
            Exit Sub
        End If
        s = vbNullString
        n = 0
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) = True Then
                If Trim$(GetPlayerGuild(Index)) = Trim(GetPlayerGuild(i)) Then
                    If s = vbNullString Then
                        s = s & GetPlayerName(i)
                    Else
                        s = s & ", " & GetPlayerName(i)
                    End If
                    n = n + 1
                End If
            End If
        Next i
        
        Call PlayerMsg(Index, "Guild Members Online (" & n & "): " & s, BrightGreen)
        Exit Sub
    'End If
    
    ' Delete Guild
    Case "destroyguild"
    'If LCase$(Parse$(0)) = "destroyguild" Then
        Call DeleteGuild(GetPlayerGuild(Index))
    'End If
        
    ' ::::::::::::::::::::::::
    ' :: New account packet ::
    ' ::::::::::::::::::::::::
    Case "newaccount"
    'If LCase$(Parse$(0)) = "newaccount" Then
        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse$(1)
            Password = Parse$(2)
            
            ' Banned?
            If IsBannedHD(Player(Index).HDSerial) Then
                 Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", you can no longer play!")
                 Exit Sub
             End If
        
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid$(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd(frmServer.txtText, "Account " & Name & " has been created.", True)
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                Call AlertMsg(Index, "Your account has been created!")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::
    ' :: Delete account packet ::
    ' :::::::::::::::::::::::::::
    Case "delaccount"
    'If LCase$(Parse$(0)) = "delaccount" Then
        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse$(1)
            Password = Parse$(2)
            
            ' Banned?
            If IsBannedHD(Player(Index).HDSerial) Then
                 Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", you can no longer play!")
                 Exit Sub
             End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If
            
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If
                        
            ' Delete names from master name file
            Call LoadPlayer(Index, Name)
            For i = 1 To MAX_CHARS
                If Trim$(Player(Index).Char(i).Name) <> vbNullString Then
                    Call DeleteName(Player(Index).Char(i).Name)
                End If
            Next i
            Call ClearPlayer(Index)
            
            ' Everything went ok
            Call Kill(App.Path & "\accounts\" & Trim$(Name) & ".ini")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
        End If
        Exit Sub
    'End If
        
    ' ::::::::::::::::::
    ' :: Login packet ::
    ' ::::::::::::::::::
    Case "login"
    'If LCase$(Parse$(0)) = "login" Then
        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse$(1)
            Password = Parse$(2)
            
            ' Banned?
            If IsBannedHD(Player(Index).HDSerial) Then
                 Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", you can no longer play!")
                 Exit Sub
             End If
            
            ' Prevent Dupeing
            For i = 1 To Len(Name)
                n = Asc(Mid$(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(Index, "Account duping is not allowed!")
                Exit Sub
                End If
            Next i
        
            ' Check versions
            If Val(Parse$(3)) < CLIENT_MAJOR Or Val(Parse$(4)) < CLIENT_MINOR Or Val(Parse$(5)) < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit ") '& Insert Web Address Here)
                Exit Sub
            End If
            
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If
        
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If
        
            If IsMultiAccounts(Name) Then
                Call AlertMsg(Index, "Multiple account logins is not authorized.")
                Exit Sub
            End If
                
            ' Check Security Code's
            If Parse(6) = "jwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf" And Parse(7) = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas" And Parse(8) = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi" And Val(Parse(9)) = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672" Then
            ' Everything went ok
            Else
                Call AlertMsg(Index, "Haha! Stop trying to hack!")
                Exit Sub
            End If
                        
            Dim Packs As String
            Packs = "MAXINFO" & SEP_CHAR
            Packs = Packs & GAME_NAME & SEP_CHAR
            Packs = Packs & MAX_PLAYERS & SEP_CHAR
            Packs = Packs & MAX_ITEMS & SEP_CHAR
            Packs = Packs & MAX_NPCS & SEP_CHAR
            Packs = Packs & MAX_SHOPS & SEP_CHAR
            Packs = Packs & MAX_SPELLS & SEP_CHAR
            Packs = Packs & MAX_MAPS & SEP_CHAR
            Packs = Packs & MAX_MAP_ITEMS & SEP_CHAR
            Packs = Packs & MAX_MAPX & SEP_CHAR
            Packs = Packs & MAX_MAPY & SEP_CHAR
            Packs = Packs & MAX_EMOTICONS & SEP_CHAR
            Packs = Packs & END_CHAR
            Call SendDataTo(Index, Packs)
    
            ' Load the player
            Call LoadPlayer(Index, Name)
            Call SendChars(Index)
    
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", True)
        End If
        Exit Sub
    'End If

    ' ::::::::::::::::::::::::::
    ' :: Add character packet ::
    ' ::::::::::::::::::::::::::
    Case "addchar"
    'If LCase$(Parse$(0)) = "addchar" Then
        If Not IsPlaying(Index) Then
            Name = Parse$(1)
            Sex = Val(Parse$(2))
            Class = Val(Parse$(3))
            CharNum = Val(Parse$(4))
        
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Then
                Call AlertMsg(Index, "Character name must be at least three characters in length.")
                Exit Sub
            End If
            
            ' Prevent being me
            If LCase$(Trim$(Name)) = "Liam" Then
                Call AlertMsg(Index, "Lets get one thing straight, you are not me, ok? :)")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid$(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
                                    
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Prevent hacking
            If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                Call HackingAttempt(Index, "Invalid Sex (dont laugh)")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Class < 0 Or Class > Max_Classes Then
                Call HackingAttempt(Index, "Invalid Class")
                Exit Sub
            End If
        
            ' Check if char already exists in slot
            If CharExist(Index, CharNum) Then
                Call AlertMsg(Index, "Character already exists!")
                Exit Sub
            End If
            
            ' Check if name is already in use
            If FindChar(Name) Then
                Call AlertMsg(Index, "Sorry, but that name is in use!")
                Exit Sub
            End If
        
            ' Everything went ok, add the character
            Call AddChar(Index, Name, Sex, Class, CharNum)
            Call SavePlayer(Index)
            Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(Index, "Character has been created!")
        End If
        Exit Sub
    'End If
        
    ' :::::::::::::::::::::::::::::::
    ' :: Deleting character packet ::
    ' :::::::::::::::::::::::::::::::
    Case "delchar"
    'If LCase$(Parse$(0)) = "delchar" Then
        If Not IsPlaying(Index) Then
            CharNum = Val(Parse$(1))
        
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
            
            Call DelChar(Index, CharNum)
            Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(Index, "Character has been deleted!")
        End If
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Using character packet ::
    ' ::::::::::::::::::::::::::::
    Case "usechar"
    'If LCase$(Parse$(0)) = "usechar" Then
        If Not IsPlaying(Index) Then
            CharNum = Val(Parse$(1))
        
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Check to make sure the character exists and if so, set it as its current char
            If CharExist(Index, CharNum) Then
                Player(Index).CharNum = CharNum
                Call JoinGame(Index)
            
                CharNum = Player(Index).CharNum
                Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", True)
                Call UpdateLabel
                
                ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
                If Not FindChar(GetPlayerName(Index)) Then
                    f = FreeFile
                    Open App.Path & "\accounts\charlist.txt" For Append As #f
                        Print #f, GetPlayerName(Index)
                    Close #f
                End If
                
                ' Now we want to check if they are already on the master guild list (this makes it add the user if they already haven't been added to the master list for older accounts)
                If Player(Index).Char(Player(Index).CharNum).Guild <> vbNullString Then
                    f = FreeFile
                    Open App.Path & "\data\guilds\guildlist.txt" For Append As #f
                        Print #f, Player(Index).Char(Player(Index).CharNum).Guild
                    Close #f
                End If
            Else
                Call AlertMsg(Index, "Character does not exist!")
            End If
        End If
        Exit Sub
    'End If
        
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    Case "saymsg"
    'If LCase$(Parse$(0)) = "saymsg" Then
        Msg = Parse$(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Say Text Modification")
                Exit Sub
            End If
        Next i
        
        If MapChatDisabled <> 1 Then
            If Player(Index).Char(Player(Index).CharNum).MapMute = 1 Then
                Call PlayerMsg(Index, "You have been muted...", SayColor)
                Exit Sub
            Else
                Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " : " & Msg & vbNullString, PLAYER_LOG)
                Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Msg & vbNullString, SayColor)
                Call MapMsg2(GetPlayerMap(Index), Msg, Index)
                Call SendDataTo(Index, "BLTOVERHEAD" & SEP_CHAR & Blue & SEP_CHAR & Msg & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "Map chat disabled...", SayColor)
            Exit Sub
        End If
    'End If
    
    Case "emotemsg"
    'If LCase$(Parse$(0)) = "emotemsg" Then
        Msg = Parse$(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Emote Text Modification")
                Exit Sub
            End If
        Next i
        
        If EmoteChatDisabled <> 1 Then
            If Player(Index).Char(Player(Index).CharNum).EmotMute = 1 Then
                Call PlayerMsg(Index, "You have been muted...", EmoteColor)
                Exit Sub
            Else
                Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
                Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Msg, EmoteColor)
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "Emote Chat disabled...", SayColor)
            Exit Sub
        End If
    'End If
    
    Case "broadcastmsg"
    'If LCase$(Parse$(0)) = "broadcastmsg" Then
        Msg = Parse$(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Broadcast Text Modification")
                Exit Sub
            End If
        Next i
        
        If BroadcastChatDisabled <> 1 Then
            If Player(Index).Char(Player(Index).CharNum).BroadcastMute = 1 Then
                Call PlayerMsg(Index, "You have been muted...", SayColor)
                Exit Sub
            Else
                s = GetPlayerName(Index) & ": " & Msg
                Call AddLog(s, PLAYER_LOG)
                Call GlobalMsg(s, BroadcastColor)
                Call TextAdd(frmServer.txtText, s, True)
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "Broadcast chat disabled...", SayColor)
            Exit Sub
        End If
    'End If
    
    Case "globalmsg"
    'If LCase$(Parse$(0)) = "globalmsg" Then
        Msg = Parse$(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Global Text Modification")
                Exit Sub
            End If
        Next i
        
        If GlobalChatDisabled <> 1 Then
            If Player(Index).Char(Player(Index).CharNum).GlobalMute = 1 Then
                Call PlayerMsg(Index, "You have been muted...", SayColor)
                Exit Sub
            Else
                If GetPlayerAccess(Index) > 0 Then
                    s = "(global) " & GetPlayerName(Index) & ": " & Msg
                    Call AddLog(s, ADMIN_LOG)
                    Call GlobalMsg(s, GlobalColor)
                    Call TextAdd(frmServer.txtText, s, True)
                End If
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "Global chat disabled...", SayColor)
            Exit Sub
        End If
    'End If
    
    Case "adminmsg"
    'If LCase$(Parse$(0)) = "adminmsg" Then
        Msg = Parse$(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Admin Text Modification")
                Exit Sub
            End If
        Next i

        If AdminChatDisabled <> 1 Then
            If Player(Index).Char(Player(Index).CharNum).AdminMute = 1 Then
                Call PlayerMsg(Index, "You have been muted...", SayColor)
                Exit Sub
            Else
                If GetPlayerAccess(Index) > 0 Then
                    Call AddLog("[" & GetPlayerName(Index) & "]: " & Msg, ADMIN_LOG)
                    Call AdminMsg("[" & GetPlayerName(Index) & "]: " & Msg, AdminColor)
                End If
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "Admin chat disabled...", SayColor)
            Exit Sub
        End If
    'End If
    
    Case "playermsg"
    'If LCase$(Parse$(0)) = "playermsg" Then
        MsgTo = FindPlayer(Parse$(1))
        Msg = Parse$(2)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Player Msg Text Modification")
                Exit Sub
            End If
        Next i
        
        If PrivateChatDisabled <> 1 Then
            If Player(Index).Char(Player(Index).CharNum).PrivMute = 1 Then
                Call PlayerMsg(Index, "You have been muted...", SayColor)
                Exit Sub
            Else
                ' Check if they are trying to talk to themselves
                If MsgTo <> Index Then
                    If MsgTo > 0 Then
                        Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                        Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
                        Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
                    Else
                        Call PlayerMsg(Index, "Player is not online.", White)
                    End If
                Else
                    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", Green)
                End If
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "Private chat disabled...", SayColor)
            Exit Sub
        End If
    'End If
    
    Case "partychat"
    'If LCase$(Parse$(0)) = "partychat" Then
        Msg = Parse$(1)
        
        If Player(Index).Party.InParty = YES Then
            If PartyChatDisabled <> 1 Then
                If Player(Index).Char(Player(Index).CharNum).PartyMute = 1 Then
                    Call PlayerMsg(Index, "You have been muted...", Grey)
                    Exit Sub
                Else
                    If Player(Index).Party.Started = YES Then
                        n = Index
                    Else
                        n = Player(Index).Party.PlayerNums(1)
                    End If
                    Call PlayerMsg(n, GetPlayerName(Index) & " (Party): " & Msg, BrightGreen)
                    For i = 1 To MAX_PARTY_MEMBERS
                        If Player(n).Party.PlayerNums(i) > 0 Then
                            Call PlayerMsg(Player(n).Party.PlayerNums(i), GetPlayerName(Index) & " (Party): " & Msg, BrightGreen)
                        End If
                    Next i
                End If
            Else
                Call PlayerMsg(Index, "Party chat disabled...", SayColor)
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "You arent in a party!", BrightRed)
        End If
        Exit Sub
    'End If
    
    Case "guildchat"
    'If LCase$(Parse$(0)) = "guildchat" Then
        
        ' Check For Player Guild
        If Trim(GetPlayerGuild(Index)) = vbNullString Then
            Call PlayerMsg(Index, "Your are not in a guild!", BrightRed)
            Exit Sub
        End If
        
        If GuildChatDisabled <> 1 Then
            If Player(Index).Char(Player(Index).CharNum).GuildMute = 1 Then
                Call PlayerMsg(Index, "You have been muted...", Grey)
                Exit Sub
            Else
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) = True Then
                        If Trim(GetPlayerGuild(Index)) = Trim(GetPlayerGuild(i)) Then
                            Call PlayerMsg(i, GetPlayerName(Index) & " (Guild)> " & Parse$(1), BrightGreen)
                        End If
                    End If
                Next i
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "Guild chat disabled...", SayColor)
            Exit Sub
        End If
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    Case "playermove"
        If Player(Index).GettingMap = NO Then
    'If LCase$(Parse$(0)) = "playermove" And Player(Index).GettingMap = NO Then
            Dir = Val(Parse$(1))
            Movement = Val(Parse$(2))
            Call SetPlayerSP(Index, Val(Parse$(3)))
            
            ' Prevent hacking
            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Call HackingAttempt(Index, "Invalid Direction")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Movement < 1 Or Movement > 2 Then
                Call HackingAttempt(Index, "Invalid Movement")
                Exit Sub
            End If
            
            ' Prevent player from moving if they have casted a spell
            If Player(Index).CastedSpell = YES Then
                ' Check if they have already casted a spell, and if so we can't let them move
                If GetTickCount > Player(Index).AttackTimer + 1000 Then
                    Player(Index).CastedSpell = NO
                Else
                    Call SendPlayerXY(Index)
                    Exit Sub
                End If
            End If
            
            Call PlayerMove(Index, Dir, Movement)
            Call SendSP(Index)
            Exit Sub
        Else
            Exit Sub
        End If
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    Case "playerdir"
        If Player(Index).GettingMap = NO Then
    'If LCase$(Parse$(0)) = "playerdir" And Player(Index).GettingMap = NO Then
            Dir = Val(Parse$(1))
            
            ' Prevent hacking
            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Call HackingAttempt(Index, "Invalid Direction")
                Exit Sub
            End If
            
            Call SetPlayerDir(Index, Dir)
            Call SendDataToMapBut(Index, GetPlayerMap(Index), "PLAYERDIR" & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
            Exit Sub
        Else
            Exit Sub
        End If
    'End If
        
    ' :::::::::::::::::::::
    ' :: Use item packet ::
    ' :::::::::::::::::::::
    Case "useitem"
    'If LCase$(Parse$(0)) = "useitem" Then
        InvNum = Val(Parse$(1))
        CharNum = Player(Index).CharNum
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
        
        If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
            n = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
            
            Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, n5 As Long
            n1 = Item(GetPlayerInvItemNum(Index, InvNum)).StrReq
            n2 = Item(GetPlayerInvItemNum(Index, InvNum)).DefReq
            n3 = Item(GetPlayerInvItemNum(Index, InvNum)).SpeedReq
            n4 = Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq
            n5 = Item(GetPlayerInvItemNum(Index, InvNum)).AccessReq
            
            ' Find out what kind of item it is
            Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum <> GetPlayerArmorSlot(Index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(Index) <> n4 Then
                                Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(Index) < n5 Then
                            Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(Index)) < n1 Then
                            Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                            Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerArmorSlot(Index, InvNum)
                    Else
                        Call SetPlayerArmorSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)

                Case ITEM_TYPE_WEAPON
                    If InvNum <> GetPlayerWeaponSlot(Index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(Index) <> n4 Then
                                Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(Index) < n5 Then
                            Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(Index)) < n1 Then
                            Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                            Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerWeaponSlot(Index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                        
                Case ITEM_TYPE_HELMET
                    If InvNum <> GetPlayerHelmetSlot(Index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(Index) <> n4 Then
                                Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(Index) < n5 Then
                            Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(Index)) < n1 Then
                            Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                            Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerHelmetSlot(Index, InvNum)
                    Else
                        Call SetPlayerHelmetSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
            
                Case ITEM_TYPE_SHIELD
                    If InvNum <> GetPlayerShieldSlot(Index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(Index) <> n4 Then
                                Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(Index) < n5 Then
                            Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(Index)) < n1 Then
                            Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                            Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerShieldSlot(Index, InvNum)
                    Else
                        Call SetPlayerShieldSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                
                Case ITEM_TYPE_LEGS
                    If InvNum <> GetPlayerLegSlot(Index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(Index) <> n4 Then
                                Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(Index) < n5 Then
                            Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(Index)) < n1 Then
                            Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                            Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerLegSlot(Index, InvNum)
                    Else
                        Call SetPlayerLegSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                    
                Case ITEM_TYPE_BOOTS
                    If InvNum <> GetPlayerBootSlot(Index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(Index) <> n4 Then
                                Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(Index) < n5 Then
                            Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(Index)) < n1 Then
                            Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                            Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPLayerBootSlot(Index, InvNum)
                    Else
                        Call SetPLayerBootSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
            
                Case ITEM_TYPE_POTIONADDHP
                    Call SendDataTo(Index, "BLTOVERHEAD" & SEP_CHAR & BrightGreen & SEP_CHAR & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & END_CHAR)
                    Call SetPlayerHP(Index, GetPlayerHP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendHP(Index)
        
                Case ITEM_TYPE_POTIONADDMP
                    Call SendDataTo(Index, "BLTOVERHEAD" & SEP_CHAR & Magenta & SEP_CHAR & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & END_CHAR)
                    Call SetPlayerMP(Index, GetPlayerMP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(Index)
        
                Case ITEM_TYPE_POTIONADDSP
                    Call SendDataTo(Index, "BLTOVERHEAD" & SEP_CHAR & Yellow & SEP_CHAR & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & END_CHAR)
                    Call SetPlayerSP(Index, GetPlayerSP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(Index)

                Case ITEM_TYPE_POTIONSUBHP
                    Call SendDataTo(Index, "BLTOVERHEAD" & SEP_CHAR & BrightRed & SEP_CHAR & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & END_CHAR)
                    Call SetPlayerHP(Index, GetPlayerHP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendHP(Index)
                
                Case ITEM_TYPE_POTIONSUBMP
                    Call SendDataTo(Index, "BLTOVERHEAD" & SEP_CHAR & Pink & SEP_CHAR & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & END_CHAR)
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(Index)
        
                Case ITEM_TYPE_POTIONSUBSP
                    Call SendDataTo(Index, "BLTOVERHEAD" & SEP_CHAR & Black & SEP_CHAR & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & END_CHAR)
                    Call SetPlayerSP(Index, GetPlayerSP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(Index)
                    
                Case ITEM_TYPE_KEY
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            If GetPlayerY(Index) > 0 Then
                                x = GetPlayerX(Index)
                                y = GetPlayerY(Index) - 1
                            Else
                                Exit Sub
                            End If
                            
                        Case DIR_DOWN
                            If GetPlayerY(Index) < MAX_MAPY Then
                                x = GetPlayerX(Index)
                                y = GetPlayerY(Index) + 1
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_LEFT
                            If GetPlayerX(Index) > 0 Then
                                x = GetPlayerX(Index) - 1
                                y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_RIGHT
                            If GetPlayerX(Index) < MAX_MAPY Then
                                x = GetPlayerX(Index) + 1
                                y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                    End Select
                    
                    ' Check if a key exists
                    If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY Then
                        ' Check if the key they are using matches the map key
                        If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(x, y).Data1 Then
                            TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
                            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", White)
                            Else
                                Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
                            End If
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
                            
                            ' Check if we are supposed to take away the item
                            If Map(GetPlayerMap(Index)).Tile(x, y).Data2 = 1 Then
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                Call PlayerMsg(Index, "The key disolves.", Yellow)
                            End If
                        End If
                    End If
                    
                    If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_DOOR Then
                        TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        
                        Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
                    End If
                    
                Case ITEM_TYPE_SPELL
                    ' Get the spell num
                    n = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        If Spell(n).ClassReq - 1 = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                            If Spell(n).LevelReq = 0 And Player(Index).Char(Player(Index).CharNum).Access < 1 Then
                                Call PlayerMsg(Index, "This spell can only be used by admins!", BrightRed)
                                Exit Sub
                            End If
                            
                            ' Make sure they are the right level
                            i = GetSpellReqLevel(Index, n)
                            If i <= GetPlayerLevel(Index) Then
                                i = FindOpenSpellSlot(Index)
                                
                                ' Make sure they have an open spell slot
                                If i > 0 Then
                                    ' Make sure they dont already have the spell
                                    If Not HasSpell(Index, n) Then
                                        Call SetPlayerSpell(Index, i, n)
                                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                        Call PlayerMsg(Index, "You study the spell carefully...", Yellow)
                                        Call PlayerMsg(Index, "You have learned a new spell!", White)
                                    Else
                                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                        Call PlayerMsg(Index, "You have already learned this spell!  The spells crumbles into dust.", BrightRed)
                                    End If
                                Else
                                    Call PlayerMsg(Index, "You have learned all that you can learn!", BrightRed)
                                End If
                            Else
                                Call PlayerMsg(Index, "You must be level " & i & " to learn this spell.", White)
                            End If
                        Else
                            Call PlayerMsg(Index, "This spell can only be learned by a " & GetClassName(Spell(n).ClassReq - 1) & ".", White)
                        End If
                    Else
                        Call PlayerMsg(Index, "This scroll is not connected to a spell, please inform an admin!", White)
                    End If
            End Select
            
            Call SendStats(Index)
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
        End If
        Exit Sub
    'End If
        
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    Case "attack"
    'If LCase$(Parse$(0)) = "attack" Then
        
        ' Check for bow
        If GetPlayerWeaponSlot(Index) > 0 Then
             If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 > 0 Then
                 Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
                 Exit Sub
             End If
        End If
        
        ' Check for Ammo
        If Arrows(GetPlayerWeaponSlot(Index)).HasAmmo = 1 Then
            ' Check to make sure they have ammo equipped.
            If GetPlayerShieldSlot(Index) = Arrows(GetPlayerShieldSlot(Index)).Ammunition Then
                If GetPlayerInvItemValue(Index, GetPlayerShieldSlot(Index)) > 0 Then
                    Call SetPlayerInvItemValue(Index, GetPlayerShieldSlot(Index), GetPlayerInvItemValue(Index, GetPlayerShieldSlot(Index)) - 1)
                    
                    ' Try to attack a player
                    For i = 1 To MAX_PLAYERS
                        ' Make sure we dont try to attack ourselves
                        If i <> Index Then
                            ' Can we attack the player?
                            If CanAttackPlayer(Index, i) Then
                                If Not CanPlayerBlockHit(i) Then
                                    ' Get the damage we can do
                                    If Not CanPlayerCriticalHit(Index) Then
                                        Damage = GetPlayerDamage(Index) - GetPlayerProtection(i)
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                                    Else
                                        n = GetPlayerDamage(Index)
                                        Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                                        Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                                        Call PlayerMsg(i, GetPlayerName(Index) & " swings with enormous might!", BrightCyan)
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                                    End If
                                    
                                    If Damage > 0 Then
                                        Call AttackPlayer(Index, i, Damage)
                                    Else
                                        Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                                    End If
                                Else
                                    Call PlayerMsg(Index, GetPlayerName(i) & "'s " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked your hit!", BrightCyan)
                                    Call PlayerMsg(i, "Your " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                                End If
                                
                                Exit Sub
                            End If
                        End If
                    Next i
                    
                    ' Try to attack a npc
                    For i = 1 To MAX_MAP_NPCS
                        ' Can we attack the npc?
                        If CanAttackNpc(Index, i) Then
                            ' Get the damage we can do
                            If Not CanPlayerCriticalHit(Index) Then
                                Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                            Else
                                n = GetPlayerDamage(Index)
                                Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                                Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                            End If
                            
                            If Damage > 0 Then
                                Call AttackNpc(Index, i, Damage)
                                Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                            Else
                                Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                                Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                            End If
                            Exit Sub
                        End If
                    Next i
                    
                    Exit Sub
                End If
            Else
                Call PlayerMsg(Index, "Out of Ammo.", SayColor)
                Exit Sub
            End If
        Else
            ' Try to attack a player
            For i = 1 To MAX_PLAYERS
                ' Make sure we dont try to attack ourselves
                If i <> Index Then
                    ' Can we attack the player?
                    If CanAttackPlayer(Index, i) Then
                        If Not CanPlayerBlockHit(i) Then
                            ' Get the damage we can do
                            If Not CanPlayerCriticalHit(Index) Then
                                Damage = GetPlayerDamage(Index) - GetPlayerProtection(i)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                            Else
                                n = GetPlayerDamage(Index)
                                Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                                Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                                Call PlayerMsg(i, GetPlayerName(Index) & " swings with enormous might!", BrightCyan)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                            End If
                            
                            If Damage > 0 Then
                                Call AttackPlayer(Index, i, Damage)
                            Else
                                Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                            End If
                        Else
                            Call PlayerMsg(Index, GetPlayerName(i) & "'s " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked your hit!", BrightCyan)
                            Call PlayerMsg(i, "Your " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                        
                        Exit Sub
                    End If
                End If
            Next i
            
            ' Try to attack a npc
            For i = 1 To MAX_MAP_NPCS
                ' Can we attack the npc?
                If CanAttackNpc(Index, i) Then
                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                    Else
                        n = GetPlayerDamage(Index)
                        Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                        Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                    End If
                    
                    If Damage > 0 Then
                        Call AttackNpc(Index, i, Damage)
                        Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    Else
                        Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                        Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                    End If
                    Exit Sub
                End If
            Next i
            Exit Sub
        End If
    'End If
    
    ' ::::::::::::::::::::::
    ' :: Use stats packet ::
    ' ::::::::::::::::::::::
    Case "usestatpoint"
    'If LCase$(Parse$(0)) = "usestatpoint" Then
        PointType = Val(Parse$(1))
        
        ' Prevent hacking
        If (PointType < 0) Or (PointType > 3) Then
            Call HackingAttempt(Index, "Invalid Point Type")
            Exit Sub
        End If
                
        ' Make sure they have points
        If GetPlayerPOINTS(Index) > 0 Then
            ' Take away a stat point
            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & Index & "," & PointType
            Else
                Select Case PointType
                    Case 0
                        Call SetPlayerSTR(Index, GetPlayerSTR(Index) + 1)
                        Call PlayerMsg(Index, "You have gained more strength!", 15)
                    Case 1
                        Call SetPlayerDEF(Index, GetPlayerDEF(Index) + 1)
                        Call PlayerMsg(Index, "You have gained more defense!", 15)
                    Case 2
                        Call SetPlayerMAGI(Index, GetPlayerMAGI(Index) + 1)
                        Call PlayerMsg(Index, "You have gained more magic abilities!", 15)
                    Case 3
                        Call SetPlayerSPEED(Index, GetPlayerSPEED(Index) + 1)
                        Call PlayerMsg(Index, "You have gained more speed!", 15)
                End Select
            End If
        Else
            Call PlayerMsg(Index, "You have no skill points to train with!", BrightRed)
        End If
                
        ' Send the update
        Call SendStats(Index)
        Exit Sub
    'End If
        
    ' ::::::::::::::::::::::::::::::::
    ' :: Player info request packet ::
    ' ::::::::::::::::::::::::::::::::
    Case "playerinforequest"
    'If LCase$(Parse$(0)) = "playerinforequest" Then
        Name = Parse$(1)
        
        i = FindPlayer(Name)
        If i > 0 Then
            Call PlayerMsg(Index, "Account: " & Trim$(Player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)
            If GetPlayerAccess(Index) > ADMIN_MONITER Then
                Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
                Call PlayerMsg(Index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
                Call PlayerMsg(Index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
                Call PlayerMsg(Index, "STR: " & GetPlayerSTR(i) & "  DEF: " & GetPlayerDEF(i) & "  MAGI: " & GetPlayerMAGI(i) & "  SPEED: " & GetPlayerSPEED(i), BrightGreen)
                n = Int(GetPlayerSTR(i) / 2) + Int(GetPlayerLevel(i) / 2)
                i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)
                If n > 100 Then n = 100
                If i > 100 Then i = 100
                Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", BrightGreen)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Warp me to packet ::
    ' :::::::::::::::::::::::
    Case "warpmeto"
    'If LCase$(Parse$(0)) = "warpmeto" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(Parse$(1))
        
        If n <> Index Then
            If n > 0 Then
                Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
                Call PlayerMsg(n, GetPlayerName(Index) & " has warped to you.", BrightBlue)
                Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
                Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot warp to yourself!", White)
        End If
                
        Exit Sub
    'End If
        
    ' :::::::::::::::::::::::
    ' :: Warp to me packet ::
    ' :::::::::::::::::::::::
    Case "warptome"
    'If LCase$(Parse$(0)) = "warptome" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(Parse$(1))
        
        If n <> Index Then
            If n > 0 Then
                Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
                Call PlayerMsg(Index, GetPlayerName(n) & " has been summoned.", BrightBlue)
                Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
        End If
        
        Exit Sub
    'End If


    ' ::::::::::::::::::::::::
    ' :: Warp to map packet ::
    ' ::::::::::::::::::::::::
    Case "warpto"
    'If LCase$(Parse$(0)) = "warpto" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The map
        n = Val(Parse$(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_MAPS Then
            Call HackingAttempt(Index, "Invalid map")
            Exit Sub
        End If
        
        Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
        Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
        Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Set sprite packet ::
    ' :::::::::::::::::::::::
    Case "setsprite"
    'If LCase$(Parse$(0)) = "setsprite" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The sprite
        n = Val(Parse$(1))
        
        Call SetPlayerSprite(Index, n)
        Call SendPlayerData(Index)
        Exit Sub
    'End If
                
    ' ::::::::::::::::::::::::::::::
    ' :: Set player sprite packet ::
    ' ::::::::::::::::::::::::::::::
    Case "setplayersprite"
    'If LCase$(Parse$(0)) = "setplayersprite" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The sprite
        i = FindPlayer(Parse$(1))
        n = Val(Parse$(2))
                
        Call SetPlayerSprite(i, n)
        Call SendPlayerData(i)
        Exit Sub
    'End If
                
    ' ::::::::::::::::::::::::::
    ' :: Stats request packet ::
    ' ::::::::::::::::::::::::::
    Case "getstats"
    'If LCase$(Parse$(0)) = "getstats" Then
        Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", White)
        Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
        Call PlayerMsg(Index, "HP: " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index) & "  MP: " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index) & "  SP: " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index), White)
        Call PlayerMsg(Index, "STR: " & GetPlayerSTR(Index) & "  DEF: " & GetPlayerDEF(Index) & "  MAGI: " & GetPlayerMAGI(Index) & "  SPEED: " & GetPlayerSPEED(Index), White)
        n = Int(GetPlayerSTR(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
        i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
        If n > 100 Then n = 100
        If i > 100 Then i = 100
        Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", White)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player request for a new map ::
    ' ::::::::::::::::::::::::::::::::::
    Case "requestnewmap"
    'If LCase$(Parse$(0)) = "requestnewmap" Then
        Dir = Val(Parse$(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(Index, "Invalid Direction")
            Exit Sub
        End If
                
        Call PlayerMove(Index, Dir, 1)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    Case "mapdata"
    'If LCase$(Parse$(0)) = "mapdata" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        
        MapNum = GetPlayerMap(Index)
        Map(MapNum).Name = Parse$(n + 1)
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        Map(MapNum).Moral = Val(Parse$(n + 3))
        Map(MapNum).Up = Val(Parse$(n + 4))
        Map(MapNum).Down = Val(Parse$(n + 5))
        Map(MapNum).Left = Val(Parse$(n + 6))
        Map(MapNum).Right = Val(Parse$(n + 7))
        Map(MapNum).Music = Parse$(n + 8)
        Map(MapNum).BootMap = Val(Parse$(n + 9))
        Map(MapNum).BootX = Val(Parse$(n + 10))
        Map(MapNum).BootY = Val(Parse$(n + 11))
        
        n = n + 12
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
            Map(MapNum).Tile(x, y).Ground = Val(Parse$(n))
            Map(MapNum).Tile(x, y).Mask = Val(Parse$(n + 1))
            Map(MapNum).Tile(x, y).Anim = Val(Parse$(n + 2))
            Map(MapNum).Tile(x, y).Mask2 = Val(Parse$(n + 3))
            Map(MapNum).Tile(x, y).M2Anim = Val(Parse$(n + 4))
            Map(MapNum).Tile(x, y).Fringe = Val(Parse$(n + 5))
            Map(MapNum).Tile(x, y).FAnim = Val(Parse$(n + 6))
            Map(MapNum).Tile(x, y).Fringe2 = Val(Parse$(n + 7))
            Map(MapNum).Tile(x, y).F2Anim = Val(Parse$(n + 8))
            Map(MapNum).Tile(x, y).Type = Val(Parse$(n + 9))
            Map(MapNum).Tile(x, y).Data1 = Val(Parse$(n + 10))
            Map(MapNum).Tile(x, y).Data2 = Val(Parse$(n + 11))
            Map(MapNum).Tile(x, y).Data3 = Val(Parse$(n + 12))
            Map(MapNum).Tile(x, y).String1 = Parse$(n + 13)
            Map(MapNum).Tile(x, y).String2 = Parse$(n + 14)
            Map(MapNum).Tile(x, y).String3 = Parse$(n + 15)

            n = n + 16
            Next x
        Next y
       
        For x = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(x) = Val(Parse$(n))
            n = n + 1
            Call ClearMapNpc(x, MapNum)
        Next x
        
        ' Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
            Call ClearMapItem(i, GetPlayerMap(Index))
        Next i
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(Index))
        
        ' Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(Index))
        Next i
        
        ' Save the map
        Call SaveMap(MapNum)
        
        ' Refresh map for everyone online
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & SEP_CHAR & END_CHAR)
                'Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
            End If
        Next i
    
        Exit Sub
    'End If

    ' ::::::::::::::::::::::::::::
    ' :: Need map yes/no packet ::
    ' ::::::::::::::::::::::::::::
    Case "needmap"
    'If LCase$(Parse$(0)) = "needmap" Then
        ' Get yes/no value
        s = LCase$(Parse$(1))
                
        If s = "yes" Then
            Call SendMap(Index, GetPlayerMap(Index))
            Call SendMapItemsTo(Index, GetPlayerMap(Index))
            Call SendMapNpcsTo(Index, GetPlayerMap(Index))
            Call SendJoinMap(Index)
            Player(Index).GettingMap = NO
            Call SendDataTo(Index, "MAPDONE" & SEP_CHAR & END_CHAR)
        Else
            Call SendMapItemsTo(Index, GetPlayerMap(Index))
            Call SendMapNpcsTo(Index, GetPlayerMap(Index))
            Call SendJoinMap(Index)
            Player(Index).GettingMap = NO
            Call SendDataTo(Index, "MAPDONE" & SEP_CHAR & END_CHAR)
        End If
        
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to pick up something packet ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    Case "mapgetitem"
    'If LCase$(Parse$(0)) = "mapgetitem" Then
        Call PlayerMapGetItem(Index)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to drop something packet ::
    ' ::::::::::::::::::::::::::::::::::::::::::::
    Case "mapdropitem"
    'If LCase$(Parse$(0)) = "mapdropitem" Then
        InvNum = Val(Parse$(1))
        Ammount = Val(Parse$(2))
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_INV Then
            Call HackingAttempt(Index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Ammount > GetPlayerInvItemValue(Index, InvNum) Then
            Call HackingAttempt(Index, "Item ammount modification")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Ammount <= 0 Then
                Call HackingAttempt(Index, "Trying to drop 0 ammount of currency")
                Exit Sub
            End If
        End If
            
        Call PlayerMapDropItem(Index, InvNum, Ammount)
        Call SendStats(Index)
        Call SendHP(Index)
        Call SendMP(Index)
        Call SendSP(Index)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Respawn map packet ::
    ' ::::::::::::::::::::::::
    Case "maprespawn"
    'If LCase$(Parse$(0)) = "maprespawn" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
            Call ClearMapItem(i, GetPlayerMap(Index))
        Next i
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(Index))
        
        ' Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(Index))
        Next i
        
        Call PlayerMsg(Index, "Map respawned.", Blue)
        Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Map report packet ::
    ' :::::::::::::::::::::::
    Case "mapreport2"
    'If LCase$(Parse$(0)) = "mapreport2" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If

        s = "Free Maps: "
        tMapStart = 1
        tMapEnd = 1
        
        For i = 1 To MAX_MAPS
            If Trim$(Map(i).Name) = vbNullString Then
                tMapEnd = tMapEnd + 1
            Else
                If tMapEnd - tMapStart > 0 Then
                    s = s & Trim$(STR$(tMapStart)) & "-" & Trim$(STR$(tMapEnd - 1)) & ", "
                End If
                tMapStart = i + 1
                tMapEnd = i + 1
            End If
        Next i
        
        s = s & Trim$(STR$(tMapStart)) & "-" & Trim$(STR$(tMapEnd - 1)) & ", "
        s = Mid$(s, 1, Len(s) - 2)
        s = s & "."
        
        Call PlayerMsg(Index, s, Brown)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Kick player packet ::
    ' ::::::::::::::::::::::::
    Case "kickplayer"
    'If LCase$(Parse$(0)) = "kickplayer" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) <= 0 Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(Parse$(1))
        
        If n <> Index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(Index) Then
                    Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(Index) & "!", White)
                    Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                    Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
                Else
                    Call PlayerMsg(Index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot kick yourself!", White)
        End If
                
        Exit Sub
    'End If
        
    ' :::::::::::::::::::::
    ' :: Ban list packet ::
    ' :::::::::::::::::::::
    Case "banlist"
    'If LCase$(Parse$(0)) = "banlist" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
             Call HackingAttempt(Index, "Admin Cloning")
             Exit Sub
        End If
       
        n = 1
       
        'Var Password = HD
       
        For i = 0 To MAX_BANS
             BIp = Ban(i).BannedIP
             If BIp = vbNullString Then
             'skip
             Else
             Name = Ban(i).BannedBy
             s = Ban(i).BannedChar
             Password = Ban(i).BannedHD
             Call PlayerMsg(Index, n & ": " & s & " ( Banned IP " & BIp & "[" & Password & "] by " & Name & " )", White)
             n = n + 1
             End If
        Next i
       
        Exit Sub
    'End If
'    If Lcase$(Parse$(0)) = "banlist" Then
'        ' Prevent hacking
'        If GetPlayerAccess(index) < ADMIN_MAPPER Then
'            Call HackingAttempt(index, "Admin Cloning")
'            Exit Sub
'        End If
'
'        n = 1
'        f = FreeFile
'        Open App.Path & "\banlist.txt" For Input As #f
'        Do While Not EOF(f)
'            Input #f, s
'            Input #f, Name
'
'            Call PlayerMsg(index, n & ": Banned IP " & s & " by " & Name, White)
'            n = n + 1
'        Loop
'        Close #f
'        Exit Sub
'    End If

    ' :::::::::::::::::::::::::
    ' :: UnBan player packet ::
    ' :::::::::::::::::::::::::
    Case "unbanplayer"
    'If LCase$(Parse$(0)) = "unbanplayer" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
             Call HackingAttempt(Index, "Admin Cloning")
             Exit Sub
        End If
       
        ' The player index
        Name = Trim$(Parse$(1))
       
        Call UnBanIndex(Name, Index)
       
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::
    ' :: HD Serial packet ::
    ' ::::::::::::::::::::::
    Case "hdserial"
    'If LCase$(Parse$(0)) = "hdserial" Then
        Player(Index).HDSerial = Parse$(1)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Ban destroy packet ::
    ' ::::::::::::::::::::::::
    Case "bandestroy"
    'If LCase$(Parse$(0)) = "bandestroy" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
             Call HackingAttempt(Index, "Admin Cloning")
             Exit Sub
        End If
       
        For n = 0 To MAX_BANS
                 Ban(i).BannedIP = vbNullString
                 Ban(i).BannedChar = vbNullString
                 Ban(i).BannedBy = vbNullString
                 Ban(i).BannedHD = vbNullString
                 Call SaveBan(i)
        Next n
       
        Call PlayerMsg(Index, "Ban list destroyed.", White)
        Exit Sub
    'End If
        
    ' :::::::::::::::::::::::
    ' :: Ban player packet ::
    ' :::::::::::::::::::::::
    Case "banplayer"
    'If LCase$(Parse$(0)) = "banplayer" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(Parse$(1))
        
        If n <> Index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(Index) Then
                    Call BanIndex(n, Index)
                Else
                    Call PlayerMsg(Index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot ban yourself!", White)
        End If
                
        Exit Sub
    'End If
    
    '::::::::::::::::::::::::
    ':: request edit arrow ::
    '::::::::::::::::::::::::
    Case "requesteditarrow"
    'If LCase$(Parse$(0)) = "requesteditarrow" Then
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
             Call HackingAttempt(Index, "Admin Cloning")
             Exit Sub
        End If
       
        Call SendDataTo(Index, "arrowEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
       
    '::::::::::::::::
    ':: edit arrow ::
    '::::::::::::::::
    Case "editarrow"
    'If LCase$(Parse$(0)) = "editarrow" Then
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
             Call HackingAttempt(Index, "Admin Cloning")
             Exit Sub
        End If

        n = Val(Parse$(1))
       
        If n < 0 Or n > MAX_ARROWS Then
             Call HackingAttempt(Index, "Invalid arrow Index")
             Exit Sub
        End If
       
        Call AddLog(GetPlayerName(Index) & " editing arrow #" & n & ".", ADMIN_LOG)
        Call SendEditArrowTo(Index, n)
        Exit Sub
    'End If
    
    ':::::::::::::::::
    ':: edit Effect ::
    ':::::::::::::::::
    Case "editeffect"
    'If LCase$(Parse$(0)) = "editEffect" Then
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
             Call HackingAttempt(Index, "Admin Cloning")
             Exit Sub
        End If

        n = Val(Parse$(1))
       
        If n < 0 Or n > MAX_EFFECTS Then
             Call HackingAttempt(Index, "Invalid effect Index")
             Exit Sub
        End If
       
        Call AddLog(GetPlayerName(Index) & " editing effect #" & n & ".", ADMIN_LOG)
        Call SendEditEffectTo(Index, n)
        Exit Sub
    'End If
   
    '::::::::::::::::
    ':: save arrow ::
    '::::::::::::::::
    Case "savearrow"
    'If LCase$(Parse$(0)) = "savearrow" Then
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
             Call HackingAttempt(Index, "Admin Cloning")
             Exit Sub
        End If
       
        n = Val(Parse$(1))
        If n < 0 Or n > MAX_ITEMS Then
             Call HackingAttempt(Index, "Invalid arrow Index")
             Exit Sub
        End If

        Arrows(n).Name = Parse$(2)
        Arrows(n).Pic = Val(Parse$(3))
        Arrows(n).Range = Val(Parse$(4))
        Arrows(n).HasAmmo = Val(Parse$(5))
        Arrows(n).Ammunition = Val(Parse$(6))

        Call SendUpdateArrowToAll(n)
        Call SaveArrow(n)
        Call AddLog(GetPlayerName(Index) & " saved arrow #" & n & ".", ADMIN_LOG)
        Exit Sub
    'End If
   
    ':::::::::::::::::
    ':: check arrow ::
    ':::::::::::::::::
    Case "checkarrows"
    'If LCase$(Parse$(0)) = "checkarrows" Then
        n = Arrows(Val(Parse$(1))).Pic
       
        Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
   
    Case "arrowhit"
    'If LCase$(Parse$(0)) = "arrowhit" Then
        n = Val(Parse$(1))
        z = Val(Parse$(2))
        x = Val(Parse$(3))
        y = Val(Parse$(4))
        
        If n = TARGET_TYPE_PLAYER Then
            ' Make sure we dont try To attack ourselves
            If z <> Index Then
                ' Can we attack the player?
                If CanAttackPlayerWithArrow(Index, z) Then
                    If Not CanPlayerBlockHit(z) Then
                        ' Get the damage we can Do
                        If Not CanPlayerCriticalHit(Index) Then
                            Damage = GetPlayerDamage(Index) - GetPlayerProtection(z)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            n = GetPlayerDamage(Index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(z)
                            
                            Call PlayerMsg(Index, "You feel a surge of energy upon shooting!", BrightCyan)
                            Call PlayerMsg(z, GetPlayerName(Index) & " shoots With amazing accuracy!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(Index, z, Damage)
                        Else
                            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call PlayerMsg(Index, GetPlayerName(z) & "'s " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked your hit!", BrightCyan)
                        Call PlayerMsg(z, "Your " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                    End If
                    Exit Sub
                End If
            End If
        ElseIf n = TARGET_TYPE_NPC Then
            ' Can we attack the npc?
            If CanAttackNpcWithArrow(Index, z) Then
                ' Get the damage we can Do
                If Not CanPlayerCriticalHit(Index) Then
                    Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), z).Num).DEF / 2)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                Else
                    n = GetPlayerDamage(Index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), z).Num).DEF / 2)
                    
                    Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(Index, z, Damage)
                    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                Else
                    Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                End If
                Exit Sub
            End If
        End If
        Exit Sub
    'End If
        
    ' :::::::::::::::::::::::::::::
    ' :: Request edit map packet ::
    ' :::::::::::::::::::::::::::::
    Case "requesteditmap"
    'If LCase$(Parse$(0)) = "requesteditmap" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "EDITMAP" & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit item packet ::
    ' ::::::::::::::::::::::::::::::
    Case "requestedititem"
    'If LCase$(Parse$(0)) = "requestedititem" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Request edit effect packet ::
    ' ::::::::::::::::::::::::::::::::
    Case "requestediteffect"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "EFFECTEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub

    ' ::::::::::::::::::::::
    ' :: Edit item packet ::
    ' ::::::::::::::::::::::
    Case "edititem"
    'If LCase$(Parse$(0)) = "edititem" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The item #
        n = Val(Parse$(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid Item Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing item #" & n & ".", ADMIN_LOG)
        Call SendEditItemTo(Index, n)
    'End If
    
    ' ::::::::::::::::::::::
    ' :: Save item packet ::
    ' ::::::::::::::::::::::
    Case "saveitem"
    'If LCase$(Parse$(0)) = "saveitem" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse$(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid Item Index")
            Exit Sub
        End If
        
        ' Update the item
        Item(n).Name = Parse$(2)
        Item(n).Pic = Val(Parse$(3))
        Item(n).Type = Val(Parse$(4))
        Item(n).Data1 = Val(Parse$(5))
        Item(n).Data2 = Val(Parse$(6))
        Item(n).Data3 = Val(Parse$(7))
        Item(n).StrReq = Val(Parse$(8))
        Item(n).DefReq = Val(Parse$(9))
        Item(n).SpeedReq = Val(Parse$(10))
        Item(n).MagiReq = Val(Parse$(11))
        Item(n).ClassReq = Val(Parse$(12))
        Item(n).AccessReq = Val(Parse$(13))
        
        Item(n).AddHP = Val(Parse$(14))
        Item(n).AddMP = Val(Parse$(15))
        Item(n).AddSP = Val(Parse$(16))
        Item(n).AddStr = Val(Parse$(17))
        Item(n).AddDef = Val(Parse$(18))
        Item(n).AddMagi = Val(Parse$(19))
        Item(n).AddSpeed = Val(Parse$(20))
        Item(n).AddEXP = Val(Parse$(21))
        
        Item(n).Desc = Parse$(22)
        
        Item(n).CannotBeRepaired = Val(Parse$(23))
        Item(n).DropOnDeath = Val(Parse$(24))
        
        ' Save it
        Call SendUpdateItemToAll(n)
        Call SaveItem(n)
        Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
        Exit Sub
    'End If
    
    
    ' ::::::::::::::::::::::::
    ' :: Save effect packet ::
    ' ::::::::::::::::::::::::
    Case "saveeffect"
    'If LCase$(Parse$(0)) = "saveeffect" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse$(1))
        If n < 0 Or n > MAX_EFFECTS Then
            Call HackingAttempt(Index, "Invalid Effect Index")
            Exit Sub
        End If
        
        ' Update the Effect
        Effect(n).Name = Parse$(2)
        Effect(n).Effect = Val(Parse$(3))
        Effect(n).Time = Val(Parse$(4))
        Effect(n).Data1 = Val(Parse$(5))
        Effect(n).Data2 = Val(Parse$(6))
        Effect(n).Data3 = Val(Parse$(7))
        
        ' Save it
        Call SendUpdateEffectToAll(n)
        Call SaveEffect(n)
        Call AddLog(GetPlayerName(Index) & " saved effect #" & n & ".", ADMIN_LOG)
        Exit Sub
    'End If

    
    ' ::::::::::::::::::::::
    ' :: Day/Night Basics ::
    ' ::::::::::::::::::::::
    Case "daynight"
    'If LCase$(Parse$(0)) = "daynight" Then
    ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Set the Time
        If GameTime = TIME_DAY Then
            GameTime = TIME_NIGHT
        Else
            GameTime = TIME_DAY
        End If
                
        Call SendTimeToAll
        Exit Sub
    'End If
    
    ':::::::::::::::::
    '::: Game Time :::
    ':::::::::::::::::
    Case "gmtime"
    'If LCase$(Parse$(0)) = "gmtime" Then
    
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        GameTime = Val(Parse$(1))
        Call SendTimeToAll
        Exit Sub
    'End If
    
    ':::::::::::::::::
    ':: Jail Player ::
    ':::::::::::::::::
    Case "jailplayer"
    'If LCase$(Parse$(0)) = "jailplayer" Then
        
        'Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Name = Parse$(1)
        Reason = Parse$(2)
        
        Call SetPlayerJail(Name, 1)
        Call SetPlayerMap(FindPlayer(Name), JAIL_MAP)
        Call SetPlayerX(FindPlayer(Name), JAIL_X)
        Call SetPlayerY(FindPlayer(Name), JAIL_Y)
        
        If Parse$(2) <> vbNullString Then
            Call GlobalMsg(GetPlayerName(Parse$(1)) & " has been jailed for " & Reason, BrightRed)
        Else
            Call GlobalMsg(GetPlayerName(FindPlayer(Parse$(1))) & " has been jailed.", BrightRed)
        End If
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Mute Player Packet ::
    ' ::::::::::::::::::::::::
    Case "muteplayer"
    'If LCase$(Parse$(0)) = "muteplayer" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
        If n <> Index Then
            If n > 0 Then
                Player(n).Char(Player(n).CharNum).BroadcastMute = 1
                Player(n).Char(Player(n).CharNum).AdminMute = 1
                Player(n).Char(Player(n).CharNum).GlobalMute = 1
                Player(n).Char(Player(n).CharNum).PrivMute = 1
                Player(n).Char(Player(n).CharNum).EmotMute = 1
                Player(n).Char(Player(n).CharNum).MapMute = 1
                Player(n).Char(Player(n).CharNum).GuildMute = 1
                Player(n).Char(Player(n).CharNum).PartyMute = 1
                
                Call PlayerMsg(n, "You have been muted.", White)
                Call GlobalMsg(GetPlayerName(n) & " has been muted.", White)
                Call AddLog(GetPlayerName(Index) & " muted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot mute yourself!", White)
        End If
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' ::UnMute Player Packet::
    ' ::::::::::::::::::::::::
    Case "unmuteplayer"
    'If LCase$(Parse$(0)) = "unmuteplayer" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
        If n <> Index Then
            If n > 0 Then
                Player(n).Char(Player(n).CharNum).BroadcastMute = 0
                Player(n).Char(Player(n).CharNum).AdminMute = 0
                Player(n).Char(Player(n).CharNum).GlobalMute = 0
                Player(n).Char(Player(n).CharNum).PrivMute = 0
                Player(n).Char(Player(n).CharNum).EmotMute = 0
                Player(n).Char(Player(n).CharNum).MapMute = 0
                Player(n).Char(Player(n).CharNum).GuildMute = 0
                Player(n).Char(Player(n).CharNum).PartyMute = 0
                
                Call PlayerMsg(n, "You have been unmuted.", White)
                Call GlobalMsg(GetPlayerName(n) & " has been unmuted.", White)
                Call AddLog(GetPlayerName(Index) & " unmuted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot unmute yourself!", White)
        End If
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::
    ' :: Admin Monitoring ::
    ' ::::::::::::::::::::::
    Case "invisible"
    'If LCase$(Parse$(0)) = "invisible" Then
        
        ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Check to see if they're already invisible
        If Player(Index).Invisible = 1 Then
            Call SetPlayerInvisible(Index, False)
        Else
            Call SetPlayerInvisible(Index, True)
        End If
        
        ' Send their updated info
        Call SendPlayerData(Index)
    'End If
    
    
    ' :::::::::::::::::::::::::::::
    ' ::Different Types of Muting::
    ' :::::::::::::::::::::::::::::
    Case "disableplayermap"
    'If LCase$(Parse$(0)) = "disableplayermap" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
    If n <> Index Then
            If n > 0 Then
                If Player(n).Char(Player(n).CharNum).MapMute = 0 Then
                    Player(n).Char(Player(n).CharNum).MapMute = 1
                    Call PlayerMsg(n, "You have been muted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been muted.", White)
                    Call AddLog(GetPlayerName(Index) & " muted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                Else
                    Player(n).Char(Player(n).CharNum).MapMute = 0
                    Call PlayerMsg(n, "You have been unmuted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been unmuted.", White)
                    Call AddLog(GetPlayerName(Index) & " unmuted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                End If
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot unmute yourself!", White)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' ::Different Types of Muting::
    ' :::::::::::::::::::::::::::::
    Case "disableplayerbroadcast"
    'If LCase$(Parse$(0)) = "disableplayerbroadcast" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
    If n <> Index Then
            If n > 0 Then
                If Player(n).Char(Player(n).CharNum).BroadcastMute = 0 Then
                    Player(n).Char(Player(n).CharNum).BroadcastMute = 1
                    Call PlayerMsg(n, "You have been muted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been muted.", White)
                    Call AddLog(GetPlayerName(Index) & " muted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                Else
                    Player(n).Char(Player(n).CharNum).BroadcastMute = 0
                    Call PlayerMsg(n, "You have been unmuted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been unmuted.", White)
                    Call AddLog(GetPlayerName(Index) & " unmuted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                End If
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot unmute yourself!", White)
        End If
        Exit Sub
    'End If
        
    ' :::::::::::::::::::::::::::::
    ' ::Different Types of Muting::
    ' :::::::::::::::::::::::::::::
    Case "disableplayerglobal"
    'If LCase$(Parse$(0)) = "disableplayerglobal" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
    If n <> Index Then
            If n > 0 Then
                Player(n).Char(Player(n).CharNum).GlobalMute = 1
                
                Call PlayerMsg(n, "You have been unmuted.", White)
                Call GlobalMsg(GetPlayerName(n) & " has been unmuted.", White)
                Call AddLog(GetPlayerName(Index) & " unmuted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot unmute yourself!", White)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' ::Different Types of Muting::
    ' :::::::::::::::::::::::::::::
    Case "disableplayeremote"
    'If LCase$(Parse$(0)) = "disableplayeremote" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
    If n <> Index Then
            If n > 0 Then
                If Player(n).Char(Player(n).CharNum).EmotMute = 0 Then
                    Player(n).Char(Player(n).CharNum).EmotMute = 1
                    Call PlayerMsg(n, "You have been muted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been muted.", White)
                    Call AddLog(GetPlayerName(Index) & " muted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                Else
                    Player(n).Char(Player(n).CharNum).EmotMute = 0
                    Call PlayerMsg(n, "You have been unmuted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been unmuted.", White)
                    Call AddLog(GetPlayerName(Index) & " unmuted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                End If
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot unmute yourself!", White)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' ::Different Types of Muting::
    ' :::::::::::::::::::::::::::::
    Case "disableplayeradmin"
    'If LCase$(Parse$(0)) = "disableplayeradmin" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
    If n <> Index Then
            If n > 0 Then
                If Player(n).Char(Player(n).CharNum).AdminMute = 0 Then
                    Player(n).Char(Player(n).CharNum).AdminMute = 1
                    Call PlayerMsg(n, "You have been muted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been muted.", White)
                    Call AddLog(GetPlayerName(Index) & " muted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                Else
                    Player(n).Char(Player(n).CharNum).AdminMute = 0
                    Call PlayerMsg(n, "You have been unmuted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been unmuted.", White)
                    Call AddLog(GetPlayerName(Index) & " unmuted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                End If
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot unmute yourself!", White)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' ::Different Types of Muting::
    ' :::::::::::::::::::::::::::::
    Case "disableplayerprivate"
    'If LCase$(Parse$(0)) = "disableplayerprivate" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
    If n <> Index Then
            If n > 0 Then
                If Player(n).Char(Player(n).CharNum).PrivMute = 0 Then
                    Player(n).Char(Player(n).CharNum).PrivMute = 1
                    Call PlayerMsg(n, "You have been muted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been muted.", White)
                    Call AddLog(GetPlayerName(Index) & " muted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                Else
                    Player(n).Char(Player(n).CharNum).PrivMute = 0
                    Call PlayerMsg(n, "You have been unmuted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been unmuted.", White)
                    Call AddLog(GetPlayerName(Index) & " unmuted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                End If
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot unmute yourself!", White)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' ::Different Types of Muting::
    ' :::::::::::::::::::::::::::::
    Case "disableplayerguild"
    'If LCase$(Parse$(0)) = "disableplayerguild" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
    If n <> Index Then
            If n > 0 Then
                If Player(n).Char(Player(n).CharNum).GuildMute = 0 Then
                    Player(n).Char(Player(n).CharNum).GuildMute = 1
                    Call PlayerMsg(n, "You have been muted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been muted.", White)
                    Call AddLog(GetPlayerName(Index) & " muted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                Else
                    Player(n).Char(Player(n).CharNum).GuildMute = 0
                    Call PlayerMsg(n, "You have been unmuted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been unmuted.", White)
                    Call AddLog(GetPlayerName(Index) & " unmuted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                End If
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot unmute yourself!", White)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' ::Different Types of Muting::
    ' :::::::::::::::::::::::::::::
    Case "disableplayerparty"
    'If LCase$(Parse$(0)) = "disableplayerparty" Then
    ' Prevent Hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
    
    ' Time To Shut 'em up
    n = FindPlayer(Parse$(1))
    
    If n <> Index Then
            If n > 0 Then
                If Player(n).Char(Player(n).CharNum).PartyMute = 0 Then
                    Player(n).Char(Player(n).CharNum).PartyMute = 1
                    Call PlayerMsg(n, "You have been muted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been muted.", White)
                    Call AddLog(GetPlayerName(Index) & " muted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                Else
                    Player(n).Char(Player(n).CharNum).PartyMute = 0
                    Call PlayerMsg(n, "You have been unmuted.", White)
                    Call GlobalMsg(GetPlayerName(n) & " has been unmuted.", White)
                    Call AddLog(GetPlayerName(Index) & " unmuted " & GetPlayerName(n) & GetPlayerMap(Index) & ".", ADMIN_LOG)
                End If
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot unmute yourself!", White)
        End If
        Exit Sub
    'End If

    ' :::::::::::::::::::::::::::::
    ' :: Request edit npc packet ::
    ' :::::::::::::::::::::::::::::
    Case "requesteditnpc"
    'If LCase$(Parse$(0)) = "requesteditnpc" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "NPCEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    Case "editnpc"
    'If LCase$(Parse$(0)) = "editnpc" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The npc #
        n = Val(Parse$(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(Index, "Invalid NPC Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing npc #" & n & ".", ADMIN_LOG)
        Call SendEditNpcTo(Index, n)
    'End If
    
    ' :::::::::::::::::::::
    ' :: Save npc packet ::
    ' :::::::::::::::::::::
    Case "savenpc"
    'If LCase$(Parse$(0)) = "savenpc" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse$(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(Index, "Invalid NPC Index")
            Exit Sub
        End If
        
        ' Update the npc
        Npc(n).Name = Parse$(2)
        Npc(n).AttackSay = Parse$(3)
        Npc(n).Sprite = Val(Parse$(4))
        Npc(n).SpawnSecs = Val(Parse$(5))
        Npc(n).Behavior = Val(Parse$(6))
        Npc(n).Range = Val(Parse$(7))
        Npc(n).STR = Val(Parse$(8))
        Npc(n).DEF = Val(Parse$(9))
        Npc(n).SPEED = Val(Parse$(10))
        Npc(n).MAGI = Val(Parse$(11))
        Npc(n).Big = Val(Parse$(12))
        Npc(n).MaxHp = Val(Parse$(13))
        Npc(n).Exp = Val(Parse$(14))
        Npc(n).Alignment = Val(Parse$(15))
        
        z = 16
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = Val(Parse$(z))
            Npc(n).ItemNPC(i).ItemNum = Val(Parse$(z + 1))
            Npc(n).ItemNPC(i).ItemValue = Val(Parse$(z + 2))
            z = z + 3
        Next i
        
        ' Save it
        Call SendUpdateNpcToAll(n)
        Call SaveNpc(n)
        Call AddLog(GetPlayerName(Index) & " saved npc #" & n & ".", ADMIN_LOG)
        Exit Sub
    'End If
            
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit shop packet ::
    ' ::::::::::::::::::::::::::::::
    Case "requesteditshop"
    'If LCase$(Parse$(0)) = "requesteditshop" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "SHOPEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet ::
    ' ::::::::::::::::::::::
    Case "editshop"
    'If LCase$(Parse$(0)) = "editshop" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The shop #
        n = Val(Parse$(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SHOPS Then
            Call HackingAttempt(Index, "Invalid Shop Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing shop #" & n & ".", ADMIN_LOG)
        Call SendEditShopTo(Index, n)
    'End If
    
    ' ::::::::::::::::::::::
    ' :: Save shop packet ::
    ' ::::::::::::::::::::::
    Case "saveshop"
    'If (LCase$(Parse$(0)) = "saveshop") Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ShopNum = Val(Parse$(1))
        
        ' Prevent hacking
        If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
            Call HackingAttempt(Index, "Invalid Shop Index")
            Exit Sub
        End If
        
        ' Update the shop
        Shop(ShopNum).Name = Parse$(2)
        Shop(ShopNum).JoinSay = Parse$(3)
        Shop(ShopNum).LeaveSay = Parse$(4)
        Shop(ShopNum).FixesItems = Val(Parse$(5))
        
        n = 6
        For i = 1 To MAX_TRADES
            Shop(ShopNum).TradeItem(i).GiveItem = Val(Parse$(n))
            Shop(ShopNum).TradeItem(i).GiveValue = Val(Parse$(n + 1))
            Shop(ShopNum).TradeItem(i).GetItem = Val(Parse$(n + 2))
            Shop(ShopNum).TradeItem(i).GetValue = Val(Parse$(n + 3))
            n = n + 4
        Next i
        
        ' Save it
        Call SendUpdateShopToAll(ShopNum)
        Call SaveShop(ShopNum)
        Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Request edit Effect packet ::
    ' ::::::::::::::::::::::::::::::::
    Case "requestediteffect"
    'If LCase$(Parse$(0)) = "requesteditEffect" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "EFFECTEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit spell packet ::
    ' :::::::::::::::::::::::::::::::
    Case "requesteditspell"
    'If LCase$(Parse$(0)) = "requesteditspell" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet ::
    ' :::::::::::::::::::::::
    Case "editspell"
    'If LCase$(Parse$(0)) = "editspell" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The spell #
        n = Val(Parse$(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(Index, "Invalid Spell Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing spell #" & n & ".", ADMIN_LOG)
        Call SendEditSpellTo(Index, n)
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Save spell packet ::
    ' :::::::::::::::::::::::
    Case "savespell"
    'If (LCase$(Parse$(0)) = "savespell") Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Spell #
        n = Val(Parse$(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(Index, "Invalid Spell Index")
            Exit Sub
        End If
        
        ' Update the spell
        Spell(n).Name = Parse$(2)
        Spell(n).Pic = Val(Parse$(3))
        Spell(n).ClassReq = Val(Parse$(4))
        Spell(n).LevelReq = Val(Parse$(5))
        Spell(n).Type = Val(Parse$(6))
        Spell(n).Data1 = Val(Parse$(7))
        Spell(n).Data2 = Val(Parse$(8))
        Spell(n).Data3 = Val(Parse$(9))
        Spell(n).MPCost = Val(Parse$(10))
        Spell(n).Sound = Val(Parse$(11))
                
        ' Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(Index) & " saving spell #" & n & ".", ADMIN_LOG)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Set access packet ::
    ' :::::::::::::::::::::::
    Case "setaccess"
    'If LCase$(Parse$(0)) = "setaccess" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Call HackingAttempt(Index, "Trying to use powers not available")
            Exit Sub
        End If
        
        ' The index
        n = FindPlayer(Parse$(1))
        ' The access
        i = Val(Parse$(2))
        
        
        ' Check for invalid access level
        If i >= 0 Or i <= 3 Then
            If GetPlayerName(Index) <> GetPlayerName(n) Then
                If GetPlayerAccess(Index) > GetPlayerAccess(n) Then
                    ' Check if player is on
                    If n > 0 Then
                        If GetPlayerAccess(n) <= 0 Then
                            Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
                        End If
                    
                        Call SetPlayerAccess(n, i)
                        Call SendPlayerData(n)
                        Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Player is not online.", White)
                    End If
                Else
                    Call PlayerMsg(Index, "Your access level is lower than " & GetPlayerName(n) & "´s.", Red)
                End If
            Else
                Call PlayerMsg(Index, "You cant change your access.", Red)
            End If
        Else
            Call PlayerMsg(Index, "Invalid access level.", Red)
        End If
                
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Whos online packet ::
    ' ::::::::::::::::::::::::
    Case "whosonline"
    'If LCase$(Parse$(0)) = "whosonline" Then
        Call SendWhosOnline(Index)
        Exit Sub
    'End If
    
    ' :::::::::::::::::
    ' :: Online list ::
    ' :::::::::::::::::
    Case "onlinelist"
    'If LCase$(Parse$(0)) = "onlinelist" Then
        'Stop
        Call SendOnlineList
        Exit Sub
    'End If

    ' :::::::::::::::::::::
    ' :: Set MOTD packet ::
    ' :::::::::::::::::::::
    Case "setmotd"
    'If LCase$(Parse$(0)) = "setmotd" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse$(1))
        Call GlobalMsg("MOTD changed to: " & Parse$(1), BrightCyan)
        Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Parse$(1), ADMIN_LOG)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::
    ' :: Trade request packet ::
    ' ::::::::::::::::::::::::::
    Case "traderequest"
    'If LCase$(Parse$(0)) = "traderequest" Then
        ' Trade num
        n = Val(Parse$(1))
        
        ' Prevent hacking
        If (n <= 0) Or (n > MAX_TRADES) Then
            Call HackingAttempt(Index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Index for shop
        i = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        
        ' Check if inv full
        x = FindOpenInvSlot(Index, Shop(i).TradeItem(n).GetItem)
        If x = 0 Then
            Call PlayerMsg(Index, "Trade unsuccessful, inventory full.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(Index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
            Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
            Call GiveItem(Index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
            Call PlayerMsg(Index, "The trade was successful!", Yellow)
        Else
            Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
        End If
        Exit Sub
    'End If

    ' :::::::::::::::::::::
    ' :: Fix item packet ::
    ' :::::::::::::::::::::
    Case "fixitem"
    'If LCase$(Parse$(0)) = "fixitem" Then
        ' Inv num
        n = Val(Parse$(1))
        
        ' Make sure its a equipable item
        If Item(GetPlayerInvItemNum(Index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, n)).Type > ITEM_TYPE_BOOTS Then
            Call PlayerMsg(Index, "You can only fix weapons, armors, helmets, and shields.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have a full inventory
        If FindOpenInvSlot(Index, GetPlayerInvItemNum(Index, n)) <= 0 Then
            Call PlayerMsg(Index, "You have no inventory space left!", BrightRed)
            Exit Sub
        End If
        
        ' Now check the rate of pay
        ItemNum = GetPlayerInvItemNum(Index, n)
        i = Int(Item(GetPlayerInvItemNum(Index, n)).Data2 / 5)
        If i <= 0 Then i = 1
        
        DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, n)
        GoldNeeded = Int(DurNeeded * i / 2)
        If GoldNeeded <= 0 Then GoldNeeded = 1
        
        ' Check if they even need it repaired
        If DurNeeded <= 0 Then
            Call PlayerMsg(Index, "This item is in perfect condition!", White)
            Exit Sub
        End If
        
        ' Check if it's unrepairable
        If Item(GetPlayerInvItemNum(Index, n)).CannotBeRepaired = 1 Then
            Call PlayerMsg(Index, "This item cannot be repaired.", White)
            Exit Sub
        End If
        
        ' Check if they have enough for at least one point
        If HasItem(Index, 1) >= i Then
            ' Check if they have enough for a total restoration
            If HasItem(Index, 1) >= GoldNeeded Then
                Call TakeItem(Index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(Index, n, Item(ItemNum).Data1)
                Call PlayerMsg(Index, "Item has been totally restored for " & GoldNeeded & " gold!", BrightBlue)
            Else
                ' They dont so restore as much as we can
                DurNeeded = (HasItem(Index, 1) / i)
                GoldNeeded = Int(DurNeeded * i / 2)
                If GoldNeeded <= 0 Then GoldNeeded = 1
                
                Call TakeItem(Index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(Index, n, GetPlayerInvItemDur(Index, n) + DurNeeded)
                Call PlayerMsg(Index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
            End If
        Else
            Call PlayerMsg(Index, "Insufficient gold to fix this item!", BrightRed)
        End If
        Exit Sub
    'End If

    ' :::::::::::::::::::
    ' :: Search packet ::
    ' :::::::::::::::::::
    Case "search"
    'If LCase$(Parse$(0)) = "search" Then
        x = Val(Parse$(1))
        y = Val(Parse$(2))
        
        ' Prevent subscript out of range
        If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
            Exit Sub
        End If
        
        ' Check for a player
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) And GetPlayerX(i) = x And GetPlayerY(i) = y Then
            
                ' Change target
                Player(Index).Target = i
                Player(Index).TargetType = TARGET_TYPE_PLAYER
                If Player(i).Invisible = 0 Then
                    Call PlayerMsg(Index, "You see " & GetPlayerName(i) & ".", Yellow)
                End If
                Exit Sub
            End If
        Next i
        
        ' Check for an item
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(Index), i).Num > 0 Then
                If MapItem(GetPlayerMap(Index), i).x = x And MapItem(GetPlayerMap(Index), i).y = y Then
                    Call PlayerMsg(Index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i
        
        ' Check for an npc
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(Index), i).Num > 0 Then
                If MapNpc(GetPlayerMap(Index), i).x = x And MapNpc(GetPlayerMap(Index), i).y = y Then
                    ' Change target
                    Player(Index).Target = i
                    Player(Index).TargetType = TARGET_TYPE_NPC
                    Call PlayerMsg(Index, "You see " & Trim$(Npc(MapNpc(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i
        
        Exit Sub
    'End If
    
' ::::::::::::::::::::::::::::::::
' :: Player Chat System Packets ::
' ::::::::::::::::::::::::::::::::
    Case "playerchat"
    'If LCase$(Parse$(0)) = "playerchat" Then
        n = FindPlayer(Parse$(1))
        If n < 1 Then
            Call PlayerMsg(Index, "Player is not online.", White)
            Exit Sub
        End If
        If n = Index Then
            Exit Sub
        End If
        If Player(Index).InChat = 1 Then
            Call PlayerMsg(Index, "Your already in a chat with another player!", Pink)
            Exit Sub
        End If

        If Player(n).InChat = 1 Then
            Call PlayerMsg(Index, "Player is already in a chat with another player!", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Chat request has been sent to " & GetPlayerName(n) & ".", Pink)
        Call PlayerMsg(n, GetPlayerName(Index) & " wants you to chat with them.  Type /chat to accept, or /chatdecline to decline.", Pink)
    
        Player(n).ChatPlayer = Index
        Player(Index).ChatPlayer = n
        Exit Sub
    'End If
    
    Case "achat"
    'If LCase$(Parse$(0)) = "achat" Then
        n = Player(Index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If
        If Player(n).ChatPlayer <> Index Then
            Call PlayerMsg(Index, "Chat failed.", Pink)
            Exit Sub
        End If
                        
        Call SendDataTo(Index, "PPCHATTING" & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "PPCHATTING" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    Case "dchat"
    'If LCase$(Parse$(0)) = "dchat" Then
        n = Player(Index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Declined chat request.", Pink)
        Call PlayerMsg(n, GetPlayerName(Index) & " declined your request.", Pink)
        
        Player(Index).ChatPlayer = 0
        Player(Index).InChat = 0
        Player(n).ChatPlayer = 0
        Player(n).InChat = 0
        Exit Sub
    'End If
    
    Case "qchat"
    'If LCase$(Parse$(0)) = "qchat" Then
        n = Player(Index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If

        Call SendDataTo(Index, "qchat" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qchat" & SEP_CHAR & END_CHAR)
        
        Player(Index).ChatPlayer = 0
        Player(Index).InChat = 0
        Player(n).ChatPlayer = 0
        Player(n).InChat = 0
        Exit Sub
    'End If
    
    Case "sendchat"
    'If LCase$(Parse$(0)) = "sendchat" Then
        n = Player(Index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If

        Call SendDataTo(n, "sendchat" & SEP_CHAR & Parse$(1) & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
' ::::::::::::::::::::::::::::::::::::
' :: END Player Chat System Packets ::
' ::::::::::::::::::::::::::::::::::::
    
    ' ::::::::::::::::::::::
    ' :: P2P Trade Packet ::
    ' ::::::::::::::::::::::
    Case "pptrade"
    'If LCase$(Parse$(0)) = "pptrade" Then
        n = FindPlayer(Parse$(1))
        
        ' Check if player is online
        If n < 1 Then
            Call PlayerMsg(Index, "Player is not online.", White)
            Exit Sub
        End If
        
        ' Prevent trading with self
        If n = Index Then
            Exit Sub
        End If
                
        ' Check if the player is in another trade
        If Player(Index).InTrade = 1 Then
            Call PlayerMsg(Index, "Your already in a trade with someone else!", Pink)
            Exit Sub
        End If
        
        ' Check where both players are
        If GetPlayerX(n) + 1 <> GetPlayerX(Index) Xor GetPlayerY(n) = GetPlayerY(Index) Then
            If GetPlayerX(n) - 1 <> GetPlayerX(Index) Xor GetPlayerY(n) = GetPlayerY(Index) Then
                If GetPlayerX(n) = GetPlayerX(Index) Xor GetPlayerY(n) + 1 <> GetPlayerY(Index) Then
                    If GetPlayerX(n) = GetPlayerX(Index) Xor GetPlayerY(n) - 1 <> GetPlayerY(Index) Then
                        Call PlayerMsg(Index, "The player needs to be beside you to trade!", Pink)
                        Call PlayerMsg(n, "You need to be beside the player to trade!", Pink)
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        ' Check to see if player is already in a trade
        If Player(n).InTrade = 1 Then
            Call PlayerMsg(Index, "Player is already in a trade!", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Trade request has been sent to " & GetPlayerName(n) & ".", Pink)
        Call PlayerMsg(n, GetPlayerName(Index) & " wants you to trade with them.  Type /accept to accept, or /decline to decline.", Pink)
    
        Player(n).TradePlayer = Index
        Player(Index).TradePlayer = n
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::
    ' :: Accept trade packet ::
    ' :::::::::::::::::::::::::
    Case "atrade"
    'If LCase$(Parse$(0)) = "atrade" Then
        n = Player(Index).TradePlayer
        
        ' Check if anyone requested a trade
        If n < 1 Then
            Call PlayerMsg(Index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        ' Check if its the right player
        If Player(n).TradePlayer <> Index Then
            Call PlayerMsg(Index, "Trade failed.", Pink)
            Exit Sub
        End If
        
        ' Check where both players are
        If GetPlayerX(n) + 1 <> GetPlayerX(Index) Xor GetPlayerY(n) = GetPlayerY(Index) Then
            If GetPlayerX(n) - 1 <> GetPlayerX(Index) Xor GetPlayerY(n) = GetPlayerY(Index) Then
                If GetPlayerX(n) = GetPlayerX(Index) Xor GetPlayerY(n) + 1 <> GetPlayerY(Index) Then
                    If GetPlayerX(n) = GetPlayerX(Index) Xor GetPlayerY(n) - 1 <> GetPlayerY(Index) Then
                        Call PlayerMsg(Index, "The player needs to be beside you to trade!", Pink)
                        Call PlayerMsg(n, "You need to be beside the player to trade!", Pink)
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        Call PlayerMsg(Index, "You are trading with " & GetPlayerName(n) & "!", Pink)
        Call PlayerMsg(n, GetPlayerName(Index) & " accepted your trade request!", Pink)
        
        Call SendDataTo(Index, "PPTRADING" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "PPTRADING" & SEP_CHAR & END_CHAR)
        
        For i = 1 To MAX_PLAYER_TRADES
            Player(Index).Trading(i).InvNum = 0
            Player(Index).Trading(i).InvName = vbNullString
            Player(n).Trading(i).InvNum = 0
            Player(n).Trading(i).InvName = vbNullString
        Next i
        
        Player(Index).InTrade = 1
        Player(Index).TradeItemMax = 0
        Player(Index).TradeItemMax2 = 0
        Player(n).InTrade = 1
        Player(n).TradeItemMax = 0
        Player(n).TradeItemMax2 = 0
        Exit Sub
    'End If

    ' :::::::::::::::::::::::
    ' :: Stop trade packet ::
    ' :::::::::::::::::::::::
    Case "qtrade"
    'If LCase$(Parse$(0)) = "qtrade" Then
        n = Player(Index).TradePlayer
        
        ' Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(Index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Stopped trading.", Pink)
        Call PlayerMsg(n, GetPlayerName(Index) & " stopped trading with you!", Pink)
        Call SendDataTo(Index, "qtrade" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        
        Player(Index).TradePlayer = 0
        Player(Index).InTrade = 0
        Player(n).TradePlayer = 0
        Player(n).InTrade = 0
        Call SendDataTo(Index, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::
    ' :: Decline trade packet ::
    ' ::::::::::::::::::::::::::
    Case "dtrade"
    'If LCase$(Parse$(0)) = "dtrade" Then
        n = Player(Index).TradePlayer
        
        ' Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(Index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Declined trade request.", Pink)
        Call PlayerMsg(n, GetPlayerName(Index) & " declined your request.", Pink)
        
        Player(Index).TradePlayer = 0
        Player(Index).InTrade = 0
        Player(n).TradePlayer = 0
        Player(n).InTrade = 0
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::
    ' :: Update trade packet ::
    ' :::::::::::::::::::::::::
    Case "updatetradeinv"
    'If LCase$(Parse$(0)) = "updatetradeinv" Then
        n = Val(Parse$(1))
    
        Player(Index).Trading(n).InvNum = Val(Parse$(2))
        Player(Index).Trading(n).InvName = Trim$(Parse$(3))
        If Player(Index).Trading(n).InvNum = 0 Then
            Player(Index).TradeItemMax = Player(Index).TradeItemMax - 1
            Player(Index).TradeOk = 0
            Player(n).TradeOk = 0
            Call SendDataTo(Index, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Else
            Player(Index).TradeItemMax = Player(Index).TradeItemMax + 1
        End If
                
        Call SendDataTo(Player(Index).TradePlayer, "updatetradeitem" & SEP_CHAR & n & SEP_CHAR & Player(Index).Trading(n).InvNum & SEP_CHAR & Player(Index).Trading(n).InvName & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Swap trade items packet ::
    ' :::::::::::::::::::::::::::::
    Case "swapitems"
    'If LCase$(Parse$(0)) = "swapitems" Then

        n = Player(Index).TradePlayer
        
        If Player(Index).TradeOk = 0 Then
            Player(Index).TradeOk = 1
            Call SendDataTo(n, "trading" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
        ElseIf Player(Index).TradeOk = 1 Then
            Player(Index).TradeOk = 0
            Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        End If
                
        If Player(Index).TradeOk = 1 And Player(n).TradeOk = 1 Then
            Player(Index).TradeItemMax2 = 0
            Player(n).TradeItemMax2 = 0

            For i = 1 To MAX_INV
                If Player(Index).TradeItemMax = Player(Index).TradeItemMax2 Then
                    Exit For
                End If
                If GetPlayerInvItemNum(n, i) < 1 Then
                    Player(Index).TradeItemMax2 = Player(Index).TradeItemMax2 + 1
                End If
            Next i

            For i = 1 To MAX_INV
                If Player(n).TradeItemMax = Player(n).TradeItemMax2 Then
                    Exit For
                End If
                If GetPlayerInvItemNum(Index, i) < 1 Then
                    Player(n).TradeItemMax2 = Player(n).TradeItemMax2 + 1
                End If
            Next i
            
            If Player(Index).TradeItemMax2 = Player(Index).TradeItemMax And Player(n).TradeItemMax2 = Player(n).TradeItemMax Then
                For i = 1 To Player(Index).TradeItemMax
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(n, x) < 1 Then
                            Call GiveItem(n, GetPlayerInvItemNum(Index, Player(Index).Trading(i).InvNum), 1)
                            Call TakeItem(Index, GetPlayerInvItemNum(Index, Player(Index).Trading(i).InvNum), 1)
                            Exit For
                        End If
                    Next x
                Next i

                For i = 1 To Player(n).TradeItemMax
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(Index, x) < 1 Then
                            Call GiveItem(Index, GetPlayerInvItemNum(n, Player(Index).Trading(i).InvNum), 1)
                            Call TakeItem(n, GetPlayerInvItemNum(n, Player(Index).Trading(i).InvNum), 1)
                            Exit For
                        End If
                    Next x
                Next i

                Call PlayerMsg(n, "Trade Successfull!", BrightGreen)
                Call PlayerMsg(Index, "Trade Successfull!", BrightGreen)
                Call SendInventory(n)
                Call SendInventory(Index)
            Else
                If Player(Index).TradeItemMax2 < Player(Index).TradeItemMax Then
                    Call PlayerMsg(Index, "Your inventory is full!", BrightRed)
                    Call PlayerMsg(n, GetPlayerName(Index) & "'s inventory is full!", BrightRed)
                End If
                If Player(n).TradeItemMax2 < Player(n).TradeItemMax Then
                    Call PlayerMsg(n, "Your inventory is full!", BrightRed)
                    Call PlayerMsg(Index, GetPlayerName(n) & "'s inventory is full!", BrightRed)
                End If
            End If
            
            Player(Index).TradePlayer = 0
            Player(Index).InTrade = 0
            Player(n).TradePlayer = 0
            Player(n).InTrade = 0
            
            Call SendDataTo(Index, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Call SendDataTo(Index, "qtrade" & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
    'End If

    ' ::::::::::::::::::
    ' :: Party packet ::
    ' ::::::::::::::::::
    Case "party"
    'If LCase$(Parse$(0)) = "party" Then
        n = FindPlayer(Parse$(1))
        
        ' Prevent partying with self
        If n = Index Then
            Exit Sub
        End If
                
        ' Check for a previous party and if so drop it
        If Player(Index).InParty = YES Then
            Call PlayerMsg(Index, "You are already in a party!", Pink)
            Exit Sub
        End If
        
        If n > 0 Then
            ' Check if its an admin
            If GetPlayerAccess(Index) > ADMIN_MONITER Then
                Call PlayerMsg(Index, "You can't join a party, you are an admin!", BrightBlue)
                Exit Sub
            End If
        
            If GetPlayerAccess(n) > ADMIN_MONITER Then
                Call PlayerMsg(Index, "Admins cannot join parties!", BrightBlue)
                Exit Sub
            End If
            
            ' Make sure they are in right level range
            If GetPlayerLevel(Index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(Index) - 5 > GetPlayerLevel(n) Then
                Call PlayerMsg(Index, "There is more then a 5 level gap between you two, party failed.", Pink)
                Exit Sub
            End If
            
            ' Check to see if player is already in a party
            If Player(n).InParty = NO Then
                Call PlayerMsg(Index, "Party request has been sent to " & GetPlayerName(n) & ".", Pink)
                Call PlayerMsg(n, GetPlayerName(Index) & " wants you to join their party.  Type /join to join, or /leave to decline.", Pink)
            
                Player(Index).PartyStarter = YES
                Player(Index).PartyPlayer = n
                Player(n).PartyPlayer = Index
            Else
                Call PlayerMsg(Index, "Player is already in a party!", Pink)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If
        Exit Sub
    'End If

    ' :::::::::::::::::::::::
    ' :: Join party packet ::
    ' :::::::::::::::::::::::
    Case "joinparty"
    'If LCase$(Parse$(0)) = "joinparty" Then
        n = Player(Index).PartyPlayer
        
        If n > 0 Then
            ' Check to make sure they aren't the starter
            If Player(Index).PartyStarter = NO Then
                ' Check to make sure that each of there party players match
                If Player(n).PartyPlayer = Index Then
                    Call PlayerMsg(Index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)
                    Call PlayerMsg(n, GetPlayerName(Index) & " has joined your party!", Pink)
                    
                    Player(Index).InParty = YES
                    Player(n).InParty = YES
                Else
                    Call PlayerMsg(Index, "Party failed.", Pink)
                End If
            Else
                Call PlayerMsg(Index, "You have not been invited to join a party!", Pink)
            End If
        Else
            Call PlayerMsg(Index, "You have not been invited into a party!", Pink)
        End If
        Exit Sub
    'End If

    ' ::::::::::::::::::::::::
    ' :: Leave party packet ::
    ' ::::::::::::::::::::::::
    Case "leaveparty"
    'If LCase$(Parse$(0)) = "leaveparty" Then
        n = Player(Index).PartyPlayer
        
        If n > 0 Then
            If Player(Index).InParty = YES Then
                Call PlayerMsg(Index, "You have left the party.", Pink)
                Call PlayerMsg(n, GetPlayerName(Index) & " has left the party.", Pink)
                
                Player(Index).PartyPlayer = 0
                Player(Index).PartyStarter = NO
                Player(Index).InParty = NO
                Player(n).PartyPlayer = 0
                Player(n).PartyStarter = NO
                Player(n).InParty = NO
            Else
                Call PlayerMsg(Index, "Declined party request.", Pink)
                Call PlayerMsg(n, GetPlayerName(Index) & " declined your request.", Pink)
                
                Player(Index).PartyPlayer = 0
                Player(Index).PartyStarter = NO
                Player(Index).InParty = NO
                Player(n).PartyPlayer = 0
                Player(n).PartyStarter = NO
                Player(n).InParty = NO
            End If
        Else
            Call PlayerMsg(Index, "You are not in a party!", Pink)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    Case "spells"
    'If LCase$(Parse$(0)) = "spells" Then
        Call SendPlayerSpells(Index)
        Exit Sub
    'End If
    
    ' :::::::::::::::::
    ' :: Cast packet ::
    ' :::::::::::::::::
    Case "cast"
    'If LCase$(Parse$(0)) = "cast" Then
        ' Spell slot
        n = Val(Parse$(1))
        
        Call CastSpell(Index, n)
        
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::
    ' :: Forget spell packet ::
    ' :::::::::::::::::::::::::
    Case "forgetspell"
    'If LCase$(Parse$(0)) = "forgetspell" Then
        ' Spell slot
        n = CLng(Parse$(1))
    
        ' Prevent subscript out of range
        If n <= 0 Or n > MAX_PLAYER_SPELLS Then
            HackingAttempt Index, "Invalid Spell Slot"
            Exit Sub
        End If
    
        With Player(Index).Char(Player(Index).CharNum)
            If .Spell(n) = 0 Then
                 PlayerMsg Index, "No spell here.", Red
            Else
                 PlayerMsg Index, "You have forgotten the spell """ & Trim$(Spell(.Spell(n)).Name) & """", Green
                 .Spell(n) = 0
                 SendSpells Index
            End If
        End With
        Exit Sub
    'End If

    ' :::::::::::::::::::::
    ' :: Location packet ::
    ' :::::::::::::::::::::
    Case "requestlocation"
    'If LCase$(Parse$(0)) = "requestlocation" Then
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Map: " & GetPlayerMap(Index) & ", X: " & GetPlayerX(Index) & ", Y: " & GetPlayerY(Index), Pink)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Time to Deposit Bank Stuff!::
    ' ::::::::::::::::::::::::::::::::
    Case "deposit"
    'If LCase$(Parse$(0)) = "deposit" Then
        BankNum = Parse$(1)
        InvNum = Parse$(2)

        ItemNum = GetPlayerInvItemNum(Index, InvNum)
        i = FindOpenBankSlot(Index, ItemNum)
        If i > 0 Then
            Call SetPlayerBankItemNum(Index, i, ItemNum)
            Call SetPlayerBankItemValue(Index, i, GetPlayerBankItemValue(Index, i) + GetPlayerInvItemValue(Index, InvNum))
            
            If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_BOOTS) Then
                Call SetPlayerBankItemDur(Index, i, GetPlayerInvItemDur(Index, InvNum))
            End If

            If GetPlayerWeaponSlot(Index) = InvNum Then
                Call SetPlayerWeaponSlot(Index, 0)
            End If
            If GetPlayerArmorSlot(Index) = InvNum Then
                Call SetPlayerArmorSlot(Index, 0)
            End If
            If GetPlayerShieldSlot(Index) = InvNum Then
                Call SetPlayerShieldSlot(Index, 0)
            End If
            If GetPlayerHelmetSlot(Index) = InvNum Then
                Call SetPlayerHelmetSlot(Index, 0)
            End If
            
            Call SetPlayerInvItemNum(Index, InvNum, 0)
            Call SetPlayerInvItemValue(Index, InvNum, 0)
            Call SetPlayerInvItemDur(Index, InvNum, 0)

            Call SendWornEquipment(Index)
            Call SendInventoryUpdate(Index, InvNum)
            Call SendBankUpdate(Index, i)
        Else
            Call PlayerMsg(Index, "Your bank is full!", Grey)
        End If
        Call PlayerMsg(Index, "Transfer Complete!", White)
        Exit Sub
        
        Call SendWornEquipment(Index)
    'End If

    ' ::::::::::::::::::::::::::::::::
    ' :: Time to Withdraw Bank Stuff::
    ' ::::::::::::::::::::::::::::::::
    Case "withdraw"
    'If LCase$(Parse$(0)) = "withdraw" Then
        BankNum = Parse$(1)
        InvNum = Parse$(2)
        'Dim ItemNum As Integer
        
        ItemNum = GetPlayerBankItemNum(Index, BankNum)
        i = FindOpenInvSlot(Index, ItemNum)
        If i > 0 Then
            Call SetPlayerInvItemNum(Index, i, ItemNum)
            Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + GetPlayerBankItemValue(Index, BankNum))
            
            If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_BOOTS) Then
                Call SetPlayerInvItemDur(Index, i, GetPlayerBankItemDur(Index, BankNum))
            End If
            
            Call SendInventoryUpdate(Index, i)
            
            Call SetPlayerBankItemNum(Index, BankNum, 0)
            Call SetPlayerBankItemValue(Index, BankNum, 0)
            Call SetPlayerBankItemDur(Index, BankNum, 0)
            
            Call SendBankUpdate(Index, BankNum)
        Else
            Call PlayerMsg(Index, "Your inventory is full.", Grey)
        End If
        Call PlayerMsg(Index, "Transfer Complete!", White)
        Exit Sub
    'End If

    
    ' :::::::::::::::::::::::::::
    ' :: Refresh Player Packet ::
    ' :::::::::::::::::::::::::::
    Case "refresh"
    'If LCase$(Parse$(0)) = "refresh" Then
        Call SendJoinMap(Index)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::
    ' :: Refresh Player Packet ::
    ' :::::::::::::::::::::::::::
    Case "buysprite"
    'If LCase$(Parse$(0)) = "buysprite" Then
        ' Check if player stepped on sprite changing tile
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
            Call PlayerMsg(Index, "You need to be on a sprite tile to buy it!", BrightRed)
            Exit Sub
        End If
        
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 = 0 Then
            Call PlayerMsg(Index, "You have bought a new sprite!", BrightGreen)
            Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
                If Item(GetPlayerInvItemNum(Index, i)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(Index, i) >= Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 Then
                        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3)
                        If GetPlayerInvItemValue(Index, i) <= 0 Then
                            Call SetPlayerInvItemNum(Index, i, 0)
                        End If
                        Call PlayerMsg(Index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                        Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(Index)
                    End If
                Else
                    If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i Then
                        Call SetPlayerInvItemNum(Index, i, 0)
                        Call PlayerMsg(Index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                        Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(Index)
                    End If
                End If
                If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i Then
                    Exit Sub
                End If
            End If
        Next i
        
        Call PlayerMsg(Index, "You dont have enough to buy this sprite!", BrightRed)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Call the admins packet ::
    ' ::::::::::::::::::::::::::::
    Case "calladmins"
    'If LCase$(Parse$(0)) = "calladmins" Then
        If GetPlayerAccess(Index) = 0 Then
            Call GlobalMsg(GetPlayerName(Index) & " needs an admin!", BrightGreen)
        Else
            Call PlayerMsg(Index, "You are an admin!", BrightGreen)
        End If
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Call check commands packet ::
    ' ::::::::::::::::::::::::::::::::
    Case "checkcommands"
    'If LCase$(Parse$(0)) = "checkcommands" Then
        s = Parse$(1)
        If Scripting = 1 Then
            PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & Index, Trim$(s)
            MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & Index
        Else
            Call PlayerMsg(Index, "Thats not a valid command!", 12)
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    Case "prompt"
    'If LCase$(Parse$(0)) = "prompt" Then
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & Index & "," & Val(Parse$(1)) & "," & Val(Parse$(2))
        End If
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Request edit emoticon packet ::
    ' ::::::::::::::::::::::::::::::::::
    Case "requesteditemoticon"
    'If LCase$(Parse$(0)) = "requesteditemoticon" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "EMOTICONEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    Case "editemoticon"
    'If LCase$(Parse$(0)) = "editemoticon" Then
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse$(1))
        
        If n < 0 Or n > MAX_EMOTICONS Then
            Call HackingAttempt(Index, "Invalid Emoticon Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing emoticon #" & n & ".", ADMIN_LOG)
        Call SendEditEmoticonTo(Index, n)
    'End If
    
    Case "saveemoticon"
    'If LCase$(Parse$(0)) = "saveemoticon" Then
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse$(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid Emoticon Index")
            Exit Sub
        End If

        Emoticons(n).Command = Parse$(2)
        Emoticons(n).Pic = Val(Parse$(3))

        Call SendUpdateEmoticonToAll(n)
        Call SaveEmoticon(n)
        Call AddLog(GetPlayerName(Index) & " saved emoticon #" & n & ".", ADMIN_LOG)
        Exit Sub
    'End If
    
    Case "checkemoticons"
    'If LCase$(Parse$(0)) = "checkemoticons" Then
        n = Emoticons(Val(Parse$(1))).Pic
        
        Call SendDataToMap(GetPlayerMap(Index), "checkemoticons" & SEP_CHAR & Index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If
    
    Case "mapreport"
    'If (LCase$(Parse$(0)) = "mapreport") Then
        
        Packs = "mapreport" & SEP_CHAR
        For i = 1 To MAX_MAPS
            Packs = Packs & Map(i).Name & SEP_CHAR
        Next i
        Packs = Packs & END_CHAR
        
        Call SendDataTo(Index, Packs)
        Exit Sub
    'End If
End Select
End Sub

'Sub CloseSocket(ByVal Index As Long) ---- Pre-IOCP
'    ' Make sure player was/is playing the game, and if so, save'm.
'    If Index > 0 Then
'        Call LeftGame(Index)
'
'        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(Index) & " has been terminated.", True)
'
'        frmServer.Socket(Index).Close
'
'        Call UpdateLabel
'        Call ClearPlayer(Index)
'    End If
'End Sub

Sub CloseSocket(ByVal Index As Long)
    ' Make sure player was/is playing the game, and if so, save'm.
    If Index > 0 And IsConnected(Index) Then
        Call LeftGame(Index)
   
        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(Index) & " has been terminated.", True)
       
        Call GameServer.Sockets(Index).Shutdown(ShutdownBoth)
        Call GameServer.Sockets(Index).CloseSocket

        'Call UpdateCaption
        Call ClearPlayer(Index)
    End If
End Sub

Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim n As Long, i As Long

    s = vbNullString
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> Index Then
            s = s & GetPlayerName(i) & ", "
            n = n + 1
        End If
    Next i
            
    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If
        
    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub SendOnlineList()
Dim Packet As String
Dim i As Long
Dim n As Long
Packet = vbNullString
n = 0
For i = 1 To MAX_PLAYERS
    If IsPlaying(i) Then
        If Player(i).Invisible = 0 Then
            Packet = Packet & SEP_CHAR & GetPlayerName(i) & SEP_CHAR
            n = n + 1
        End If
    End If
Next i

Packet = "ONLINELIST" & SEP_CHAR & n & Packet & END_CHAR

Call SendDataToAll(Packet)
End Sub

Sub SendChars(ByVal Index As Long)
Dim Packet As String
Dim i As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim$(Player(Index).Char(i).Name) & SEP_CHAR & Trim$(Class(Player(Index).Char(i).Class).Name) & SEP_CHAR & Player(Index).Char(i).Level & SEP_CHAR & Player(Index).Char(i).Sprite & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendJoinMap(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = vbNullString
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> Index And GetPlayerMap(i) = GetPlayerMap(Index) Then
            Packet = Packet & "PLAYERDATA" & SEP_CHAR & i & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & GetPlayerSprite(i) & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & GetPlayerX(i) & SEP_CHAR & GetPlayerY(i) & SEP_CHAR & GetPlayerDir(i) & SEP_CHAR & GetPlayerAccess(i) & SEP_CHAR & GetPlayerPK(i) & SEP_CHAR & GetPlayerGuild(i) & SEP_CHAR & GetPlayerGuildAccess(i) & SEP_CHAR & END_CHAR
            Call SendDataTo(Index, Packet)
        End If
    Next i
    
    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & GetPlayerGuild(Index) & SEP_CHAR & GetPlayerGuildAccess(Index) & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & GetPlayerGuild(Index) & SEP_CHAR & GetPlayerGuildAccess(Index) & SEP_CHAR & END_CHAR
    Call SendDataToMapBut(Index, MapNum, Packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & GetPlayerGuild(Index) & SEP_CHAR & GetPlayerGuildAccess(Index) & SEP_CHAR & GetPlayerClass(Index) & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long
Dim i As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR
    
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
    P1 = Mid$(Packet, 1, x)
    P2 = Mid$(Packet, x + 1, Len(Packet) - x)
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendItems(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) <> vbNullString Then
            Call SendUpdateItemTo(Index, i)
        End If
    Next i
End Sub

Sub SendEmoticons(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(i).Command) <> vbNullString Then
            Call SendUpdateEmoticonTo(Index, i)
        End If
    Next i
End Sub

Sub SendNpcs(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim$(Npc(i).Name) <> vbNullString Then
            Call SendUpdateNpcTo(Index, i)
        End If
    Next i
End Sub

Sub SendInventory(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemDur(Index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendBankUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERBANKUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerBankItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerBankItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerBankItemDur(Index, InvSlot) & SEP_CHAR & Index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendWornEquipment(ByVal Index As Long)
Dim Packet As String
    
    Packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & GetPlayerLegSlot(Index) & SEP_CHAR & GetPlayerBootSlot(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendHP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendStats(ByVal Index As Long)
Dim Packet As String
    
    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).MaleSprite & SEP_CHAR & Class(i).FemaleSprite & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(Index, Packet)
End Sub

Sub SendLeftGame(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & vbNullString & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR
    Call SendDataToAllBut(Index, Packet)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagiReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).CannotBeRepaired & SEP_CHAR & Item(ItemNum).DropOnDeath
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateEffectToAll(ByVal EffectNum As Long)
Dim Packet As String

    Packet = "UPDATEEFFECT" & SEP_CHAR & EffectNum & SEP_CHAR & Trim$(Effect(EffectNum).Name) & SEP_CHAR & Effect(EffectNum).Effect & SEP_CHAR & Effect(EffectNum).Time & SEP_CHAR & Effect(EffectNum).Data1 & SEP_CHAR & Effect(EffectNum).Data2 & SEP_CHAR & Effect(EffectNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateEffectTo(ByVal Index As Long, ByVal EffectNum As Long)
Dim Packet As String

    Packet = "UPDATEEFFECT" & SEP_CHAR & EffectNum & SEP_CHAR & Trim$(Effect(EffectNum).Name) & SEP_CHAR & Effect(EffectNum).Effect & SEP_CHAR & Effect(EffectNum).Time & SEP_CHAR & Effect(EffectNum).Data1 & SEP_CHAR & Effect(EffectNum).Data2 & SEP_CHAR & Effect(EffectNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditEmoticonTo(ByVal Index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).Alignment & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendShops(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim$(Shop(i).Name) <> vbNullString Then
            Call SendUpdateShopTo(Index, i)
        End If
    Next i
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpells(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).Name) <> vbNullString Then
            Call SendUpdateSpellTo(Index, i)
        End If
    Next i
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).Pic & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, x As Long, y As Long, z As Long, XX As Long

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
                Call PlayerMsg(Index, Trim(Item(x).Name) & " can be used by all classes.", Yellow)
            Else
                Call PlayerMsg(Index, Trim(Item(x).Name) & " can only be used by a " & GetClassName(y - 1) & ".", Yellow)
            End If
        End If
        If x < 1 Then
            z = z + 1
        End If
    Next i
    Packet = Packet & END_CHAR
    
    If z = (MAX_TRADES * 6) Then
        Call PlayerMsg(Index, "This shop has nothing to sell!", BrightRed)
    Else
        Call SendDataTo(Index, Packet)
    End If
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(Index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherTo(ByVal Index As Long)
Dim Packet As String

    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If
    Next i
End Sub

Sub SendTimeTo(ByVal Index As Long)
Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTimeToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If
    Next i
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal Index As Long)
Dim Packet As String

    Packet = "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & Index & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendArrows(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(Index, i)
    Next i
End Sub

Sub SendEffects(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_EFFECTS
        Call SendUpdateEffectTo(Index, i)
    Next i
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEARROW" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).HasAmmo & SEP_CHAR & Arrows(ItemNum).Ammunition & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateArrowTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEARROW" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).HasAmmo & SEP_CHAR & Arrows(ItemNum).Ammunition & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditArrowTo(ByVal Index As Long, ByVal ArrowNum As Long)
Dim Packet As String

    Packet = "EDITARROW" & SEP_CHAR & ArrowNum & SEP_CHAR & Trim$(Arrows(ArrowNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditEffectTo(ByVal Index As Long, ByVal EffectNum As Long)
Dim Packet As String

    Packet = "EDITEFFECT" & SEP_CHAR & EffectNum & SEP_CHAR & Trim$(Effect(EffectNum).Name) & SEP_CHAR & Effect(EffectNum).Effect & SEP_CHAR & Effect(EffectNum).Time & SEP_CHAR & Effect(EffectNum).Data1 & SEP_CHAR & Effect(EffectNum).Data2 & SEP_CHAR & Effect(EffectNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub LoadBank(ByVal Index As Integer)
    Dim Packet As String
    Dim i As Integer
    
    For i = 1 To MAX_BANK
        Call SendBankUpdate(Index, i)
    Next i
    Packet = "loadbank" & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub CreateGuild(ByVal Index As Long, ByVal Name As String)
    ' Check if they are alredy in a guild
    If GetPlayerGuild(Index) <> vbNullString Then
        Call PlayerMsg(Index, "You are already in a guild!", BrightRed)
        Exit Sub
    End If
    
    ' Check if name is already in use
    If FindGuild(Name) Then
        Call PlayerMsg(Index, "Guild already exists!", SayColor)
        Exit Sub
    End If
    
    ' Check to see if they have the $ to start it.
    For i = 1 To MAX_INV
        If Trim$(LCase$(Item(GetPlayerInvItemNum(Index, i)).Name)) = "gold" Then
            If GetPlayerInvItemValue(Index, i) >= 5000 Then
                Call TakeItem(Index, GetPlayerInvItemNum(Index, i), 5000)
                Call SetPlayerGuild(Index, (Parse$(1)))
                Call SetPlayerGuildAccess(Index, 4)
                Call SendPlayerData(Index)
                Call PlayerMsg(Index, "You have successfully created a guild!", SayColor)
                Exit Sub
            Else
                Call PlayerMsg(Index, "You need " & 5000 - GetPlayerInvItemValue(Index, i) & " more gold to start a guild!", BrightRed)
                Exit Sub
            End If
        End If
    Next i
    
    Call PlayerMsg(Index, "You need 5000 Gold to start a guild!", SayColor)
    
    Call MakeGuildFile(Name)
End Sub

Sub MakeGuildFile(ByVal Name As String, ByVal GuildNum As Byte)
Dim FileName As String
Dim nFileNum As Integer
Dim StartByte As Long
   
    FileName = App.Path & "\data\guilds\" & Name & ".bin"
    
    nFileNum = FreeFile
    Open FileName For Binary As #nFileNum
    
            StartByte = 26 * (GuildNum - 1) + 1
            
        Put #nFileNum, StartByte, Guild(GuildNum).Name
        For i = 1 To MAX_GUILD_MEMBERS
            Put #nFileNum, , Guild(GuildNum).PlayerName(i)
            Put #nFileNum, , Guild(GuildNum).Rank(i)
        Next i
    Close #nFileNum
End Sub
