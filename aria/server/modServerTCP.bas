Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = GAME_NAME & " - Server"
    frmServer.lblIP.Caption = "Ip Address: " & frmServer.Socket(0).LocalIP
    frmServer.lblPort.Caption = "Port: " & STR(frmServer.Socket(0).LocalPort)
    frmServer.TPO.Caption = "Total Players Online: " & TotalOnlinePlayers
    If frmServer.Visible = False Then
        nid.szTip = frmServer.Caption & vbNewLine & TotalOnlinePlayers & " Player(s) Online" & vbNullChar
        Call Shell_NotifyIcon(NIM_MODIFY, nid)
    End If
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
If Index <= 0 Or Index > MAX_PLAYERS Then
    IsPlaying = False
    Exit Function
End If
    If IsConnected(Index) And Player(Index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If IsConnected(Index) And Trim(Player(Index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim I As Long

    IsMultiAccounts = False
    For I = 1 To MAX_PLAYERS
        If IsConnected(I) And LCase(Trim(Player(I).Login)) = LCase(Trim(Login)) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next I
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
Dim I As Long
Dim n As Long

    n = 0
    IsMultiIPOnline = False
    For I = 1 To MAX_PLAYERS
        If IsConnected(I) And Trim(GetPlayerIP(I)) = Trim(IP) Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next I
End Function

Function IsBanned(ByVal IP As String) As Boolean
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
End Function

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
Dim I As Long, n As Long, startc As Long

    If IsConnected(Index) Then
        frmServer.Socket(Index).SendData Data
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
        If IsPlaying(I) And I <> Index Then
            Call SendDataTo(I, Data)
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
            If GetPlayerMap(I) = MapNum And I <> Index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim I As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerAccess(I) > 0 Then
            Call SendDataTo(I, Packet)
        End If
    Next I
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
    Call CloseSocket(Index)
End Sub

Sub PlainMsg(ByVal Index As Long, ByVal Msg As String, ByVal num As Long)
Dim Packet As String

    Packet = "PLAINMSG" & SEP_CHAR & Msg & SEP_CHAR & num & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
            Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", True)
        End If
    
        Call AlertMsg(Index, "You have been kicked from " & GAME_NAME & " for (" & Reason & ")!")
    End If
End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
Dim I As Long

    If (Index = 0) Then
        I = FindOpenPlayerSlot
        
        If I <> 0 Then
            ' Whoho, we can connect them
            frmServer.Socket(I).Close
            frmServer.Socket(I).Accept SocketId
            Call SocketConnected(I)
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    If Index <> 0 Then
        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(Index)) Then
                Call TextAdd(frmServer.txtText(0), "Received connection from " & GetPlayerIP(Index) & ".", True)
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
Dim Buffer As String
Dim Packet As String
Dim top As String * 3
Dim Start As Long

    If Index > 0 Then
        frmServer.Socket(Index).GetData Buffer, vbString, DataLength
        
        If Buffer = "top" Then
            top = STR(TotalOnlinePlayers)
            Call SendDataTo(Index, top)
            Call CloseSocket(Index)
        End If
            
        Player(Index).Buffer = Player(Index).Buffer & Buffer
        
        Start = InStr(Player(Index).Buffer, END_CHAR)
        Do While Start > 0
            Packet = Mid(Player(Index).Buffer, 1, Start - 1)
            Player(Index).Buffer = Mid(Player(Index).Buffer, Start + 1, Len(Player(Index).Buffer))
            Player(Index).DataPackets = Player(Index).DataPackets + 1
            Start = InStr(Player(Index).Buffer, END_CHAR)
            If Len(Packet) > 0 Then
                Call HandleData(Index, Packet)
            End If
        Loop
                
        ' Not useful
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
    End If
End Sub

Sub HandleData(ByVal Index As Long, ByVal Data As String)
Dim Parse() As String
Dim Name As String
Dim Password As String
Dim Email As String
Dim Sex As Long
Dim Class As Long
Dim CharNum As Long
Dim Msg As String
Dim IPMask As String
Dim BanSlot As Long
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Amount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Movement As Long
Dim I As Long, n As Long, X As Long, Y As Long, f As Long
Dim MapNum As Long
Dim s As String
Dim tMapStart As Long, tMapEnd As Long
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long
Dim z As Long
Dim Packet As String
Dim BX As Long, BY As Long

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)

' Parse's Without Being Online
If Not IsPlaying(Index) Then
    Select Case LCase(Parse(0))
        Case "gatglasses"
            Call SendNewCharClasses(Index)
            Exit Sub
            
        Case "newfaccountied"
            If Not IsLoggedIn(Index) Then
                Name = Parse(1)
                Password = Parse(2)
                Email = Parse(3)
                        
                For I = 1 To Len(Name)
                    n = Asc(Mid(Name, I, 1))
                    
                    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                    Else
                        Call PlainMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", 1)
                        Exit Sub
                    End If
                Next I
                
                If Not AccountExist(Name) Then
                    Call AddAccount(Index, Name, Password, Email)
                    Call TextAdd(frmServer.txtText(0), "Account " & Name & " has been created.", True)
                    Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                    Call PlainMsg(Index, "Your account has been created!", 1)
                Else
                    Call PlainMsg(Index, "Sorry, that account name is already taken!", 1)
                End If
            End If
            Exit Sub
        
        Case "delimaccounted"
            If Not IsLoggedIn(Index) Then
                Name = Parse(1)
                Password = Parse(2)
                            
                If Not AccountExist(Name) Then
                    Call PlainMsg(Index, "That account name does not exist.", 2)
                    Exit Sub
                End If
                
                If Not PasswordOK(Name, Password) Then
                    Call PlainMsg(Index, "Incorrect password.", 2)
                    Exit Sub
                End If
                            
                Call LoadPlayer(Index, Name)
                For I = 1 To MAX_CHARS
                    If Trim(Player(Index).Char(I).Name) <> "" Then
                        Call DeleteName(Player(Index).Char(I).Name)
                    End If
                Next I
                Call ClearPlayer(Index)
                
                Call Kill(App.Path & "\accounts\" & Trim(Name) & ".ini")
                Call AddLog("Account " & Trim(Name) & " has been deleted.", PLAYER_LOG)
                Call PlainMsg(Index, "Your account has been deleted.", 2)
            End If
            Exit Sub
            
        Case "logination"
            If Not IsLoggedIn(Index) Then
                Name = Parse(1)
                Password = Parse(2)
                
                For I = 1 To Len(Name)
                    n = Asc(Mid(Name, I, 1))
                    
                    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                    Else
                        Call PlainMsg(Index, "Account duping is not allowed!", 3)
                    Exit Sub
                    End If
                Next I
            
                If Val(Parse(3)) < CLIENT_MAJOR Or Val(Parse(4)) < CLIENT_MINOR Or Val(Parse(5)) < CLIENT_REVISION Then
                    Call PlainMsg(Index, "Version outdated, please visit " & Trim(GetVar(App.Path & "\Data.ini", "CONFIG", "WebSite")), 3)
                    Exit Sub
                End If
                            
                If Not AccountExist(Name) Then
                    Call PlainMsg(Index, "That account name does not exist.", 3)
                    Exit Sub
                End If
            
                If Not PasswordOK(Name, Password) Then
                    Call PlainMsg(Index, "Incorrect password.", 3)
                    Exit Sub
                End If
            
                If IsMultiAccounts(Name) Then
                    Call PlainMsg(Index, "Multiple account logins is not authorized.", 3)
                    Exit Sub
                End If
                
                If frmServer.Closed.Value = Checked Then
                    Call PlainMsg(Index, "The server is closed at the moment!", 3)
                    Exit Sub
                End If
                    
                If Parse(6) = SEC_CODE1 And Parse(7) = SEC_CODE2 And Parse(8) = SEC_CODE3 And Val(Parse(9)) = Val(SEC_CODE4) Then
                Else
                    Call AlertMsg(Index, "Script Kiddy Alert!")
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
                Packs = Packs & MAX_FORMS & SEP_CHAR
                Packs = Packs & MAX_OBJECTS & SEP_CHAR
                Packs = Packs & RsText & SEP_CHAR
                Packs = Packs & MiniMap & SEP_CHAR
                Packs = Packs & END_CHAR
                Call SendDataTo(Index, Packs)
        
                Call LoadPlayer(Index, Name)
                Call SendChars(Index)
        
                Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", True)
            End If
            Exit Sub
    
        Case "addachara"
                Name = Parse(1)
                Sex = Val(Parse(2))
                Class = Val(Parse(3))
                CharNum = Val(Parse(4))
                        
                If LCase(Trim(Name)) = "Liam" Then
                    Call PlainMsg(Index, "Lets get one thing straight, you are not me, ok? :)", 4)
                    Exit Sub
                End If
                                
                For I = 1 To Len(Name)
                    n = Asc(Mid(Name, I, 1))
                    
                    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                    Else
                        Call PlainMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", 4)
                        Exit Sub
                    End If
                Next I
                                        
                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(Index, "Invalid CharNum")
                    Exit Sub
                End If
            
                If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                    Call HackingAttempt(Index, "Invalid Sex")
                    Exit Sub
                End If
                
                If Class < 0 Or Class > MAX_CLASSES Then
                    Call HackingAttempt(Index, "Invalid Class")
                    Exit Sub
                End If
    
                If CharExist(Index, CharNum) Then
                    Call PlainMsg(Index, "Character already exists!", 4)
                    Exit Sub
                End If
    
                If FindChar(Name) Then
                    Call PlainMsg(Index, "Sorry, but that name is in use!", 4)
                    Exit Sub
                End If
    
                Call AddChar(Index, Name, Sex, Class, CharNum)
                Call SavePlayer(Index)
                Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
                Call SendChars(Index)
                Call PlainMsg(Index, "Character has been created!", 5)
            Exit Sub
    
        Case "delimbocharu"
                CharNum = Val(Parse(1))
    
                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(Index, "Invalid CharNum")
                    Exit Sub
                End If
                
                Call DelChar(Index, CharNum)
                Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
                Call SendChars(Index)
                Call PlainMsg(Index, "Character has been deleted!", 5)
            Exit Sub
        
        Case "usagakarim"
                CharNum = Val(Parse(1))
    
                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(Index, "Invalid CharNum")
                    Exit Sub
                End If
            
                If CharExist(Index, CharNum) Then
                    Player(Index).CharNum = CharNum
                    
                    If frmServer.GMOnly.Value = Checked Then
                        If GetPlayerAccess(Index) <= 0 Then
                            Call PlainMsg(Index, "The server is only available to GMs at the moment!", 5)
                            'Call HackingAttempt(Index, "The server is only available to GMs at the moment!")
                            Exit Sub
                        End If
                    End If
                                        
                    Call JoinGame(Index)
                
                    CharNum = Player(Index).CharNum
                    Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
                    Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", True)
                    Call UpdateCaption
                    
                    If Not FindChar(GetPlayerName(Index)) Then
                        f = FreeFile
                        Open App.Path & "\accounts\charlist.txt" For Append As #f
                            Print #f, GetPlayerName(Index)
                        Close #f
                    End If
                Else
                    Call PlainMsg(Index, "Character does not exist!", 5)
                End If
            Exit Sub
    End Select
End If
        
' Parse's With Being Online And Playing
If IsPlaying(Index) = False Then Exit Sub
If IsConnected(Index) = False Then Exit Sub
Select Case LCase(Parse(0))
    ' :::::::::::::::::::
    ' :: Guilds Packet ::
    ' :::::::::::::::::::
    ' Access
    Case "guildchangeaccess"
        ' Check the requirements.
        If Parse(1) = "" Then
            Call PlayerMsg(Index, "You must enter a player Name To proceed.", White)
            Exit Sub
        End If
       
        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(Index, "Player Is offline", White)
            Exit Sub
        End If
   
        If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then
            Call PlayerMsg(Index, "Player Is Not In your guild", Red)
            Exit Sub
        End If
   
        'Set the player's New access level
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), Parse(2))
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
    
    ' Disown
    Case "guilddisown"
        ' Check if all the requirements
        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(Index, "Player is offline", White)
            Exit Sub
        End If
        
        If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then
            Call PlayerMsg(Index, "Player is not in your guild", Red)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(FindPlayer(Parse(1))) > GetPlayerGuildAccess(Index) Then
            Call PlayerMsg(Index, "Player has a higher guild level than you.", Red)
            Exit Sub
        End If
        
        'Player checks out, take him out of the guild
        Call SetPlayerGuild(FindPlayer(Parse(1)), "")
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub

    ' Leave Guild
    Case "guildleave"
        ' Check if they can leave
        If GetPlayerGuild(Index) = "" Then
            Call PlayerMsg(Index, "You are not in a guild.", Red)
            Exit Sub
        End If
        
        Call SetPlayerGuild(Index, "")
        Call SetPlayerGuildAccess(Index, 0)
        Call SendPlayerData(Index)
        Exit Sub
    
    ' Make A New Guild
    Case "makeguild"
        ' Check if the Owner is Online
        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(Index, "Player is offline", White)
            Exit Sub
        End If
        
        ' Check if they are alredy in a guild
            If GetPlayerGuild(FindPlayer(Parse(1))) <> "" Then
            Call PlayerMsg(Index, "Player is already in a guild", Red)
            Exit Sub
        End If
        
        ' If everything is ok then lets make the guild
        Call SetPlayerGuild(FindPlayer(Parse(1)), (Parse(2)))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 3)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
    
    ' Make A Member
    Case "guildmember"
        ' Check if its possible to admit the member
        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(Index, "Player is offline", White)
            Exit Sub
        End If
        
        If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then
            Call PlayerMsg(Index, "That player is not in your guild", Red)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(FindPlayer(Parse(1))) > 1 Then
            Call PlayerMsg(Index, "That player has already been admitted", Red)
            Exit Sub
        End If
        
        'All has gone well, set the guild access to 1
        Call SetPlayerGuild(FindPlayer(Parse(1)), GetPlayerGuild(Index))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 1)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
    
    ' Make A Trainie
    Case "guildtrainee"
        ' Check if its possible to induct member
        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(Index, "Player is offline", White)
            Exit Sub
        End If
        
        If GetPlayerGuild(FindPlayer(Parse(1))) <> "" Then
            Call PlayerMsg(Index, "Player is already in a guild", Red)
            Exit Sub
        End If
        
        'It is possible, so set the guild to index's guild, and the access level to 0
        Call SetPlayerGuild(FindPlayer(Parse(1)), GetPlayerGuild(Index))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
        
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    Case "saymsg"
        Msg = Parse(1)
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Say Text Modification")
                Exit Sub
            End If
        Next I
        
        If frmServer.chkM.Value = Unchecked Then
            If GetPlayerAccess(Index) <= 0 Then
                Call PlayerMsg(Index, "Map messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If
        End If
        
        Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " : " & Msg & "", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " : " & Msg & "", SayColor)
        Call MapMsg2(GetPlayerMap(Index), Msg, Index)
        TextAdd frmServer.txtText(3), GetPlayerName(Index) & " On Map " & GetPlayerMap(Index) & ": " & Msg, True
        Exit Sub

    Case "emotemsg"
        Msg = Parse(1)
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Emote Text Modification")
                Exit Sub
            End If
        Next I
        
        If frmServer.chkE.Value = Unchecked Then
            If GetPlayerAccess(Index) <= 0 Then
                Call PlayerMsg(Index, "Emote messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If
        End If
        
        Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Msg, EmoteColor)
        TextAdd frmServer.txtText(6), GetPlayerName(Index) & " " & Msg, True
        Exit Sub
 
    Case "broadcastmsg"
        Msg = Parse(1)
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Broadcast Text Modification")
                Exit Sub
            End If
        Next I
        
        If frmServer.chkBC.Value = Unchecked Then
            If GetPlayerAccess(Index) <= 0 Then
                Call PlayerMsg(Index, "Broadcast messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If
        End If
        
        If Player(Index).Mute = True Then Exit Sub
        
        s = GetPlayerName(Index) & ": " & Msg
        Call AddLog(s, PLAYER_LOG)
        Call GlobalMsg(s, BroadcastColor)
        Call TextAdd(frmServer.txtText(0), s, True)
        TextAdd frmServer.txtText(1), GetPlayerName(Index) & ": " & Msg, True
        Exit Sub
    
    Case "globalmsg"
        Msg = Parse(1)
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Global Text Modification")
                Exit Sub
            End If
        Next I
        
        If frmServer.chkG.Value = Unchecked Then
            If GetPlayerAccess(Index) <= 0 Then
                Call PlayerMsg(Index, "Global messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If
        End If
        
        If Player(Index).Mute = True Then Exit Sub
        
        If GetPlayerAccess(Index) > 0 Then
            s = "(Global) " & GetPlayerName(Index) & ": " & Msg
            Call AddLog(s, ADMIN_LOG)
            Call GlobalMsg(s, GlobalColor)
            Call TextAdd(frmServer.txtText(0), s, True)
        End If
        TextAdd frmServer.txtText(2), GetPlayerName(Index) & ": " & Msg, True
        Exit Sub
    
    Case "adminmsg"
        Msg = Parse(1)
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Admin Text Modification")
                Exit Sub
            End If
        Next I
        
        If frmServer.chkA.Value = Unchecked Then
            Call PlayerMsg(Index, "Admin messages have been disabled by the server!", BrightRed)
            Exit Sub
        End If
        
        If GetPlayerAccess(Index) > 0 Then
            Call AddLog("(Admin) " & GetPlayerName(Index) & ": " & Msg, ADMIN_LOG)
            Call AdminMsg("(Admin) " & GetPlayerName(Index) & ": " & Msg, AdminColor)
        End If
        TextAdd frmServer.txtText(5), GetPlayerName(Index) & ": " & Msg, True
        Exit Sub
  
    Case "playermsg"
        MsgTo = FindPlayer(Parse(1))
        Msg = Parse(2)
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Player Msg Text Modification")
                Exit Sub
            End If
        Next I
        
        If frmServer.chkP.Value = Unchecked Then
            If GetPlayerAccess(Index) <= 0 Then
                Call PlayerMsg(Index, "PM messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If
        End If
        
        ' Check if they are trying to talk to themselves
        If MsgTo <> Index Then
            If MsgTo > 0 Then
                Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                Call PlayerMsg(MsgTo, "(Private) " & GetPlayerName(Index) & ": " & Msg, TellColor)
                Call PlayerMsg(Index, "(Private) To " & GetPlayerName(MsgTo) & ": " & Msg, TellColor)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
                Exit Sub
            End If
        End If
        TextAdd frmServer.txtText(4), "To " & GetPlayerName(MsgTo) & " From " & GetPlayerName(Index) & ": " & Msg, True
        Exit Sub
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    Case "playermove"
        If Player(Index).GettingMap = YES Then
            Exit Sub
        End If
        
        Dir = Val(Parse(1))
        Movement = Val(Parse(2))
        
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
        Exit Sub
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    Case "playerdir"
        If Player(Index).GettingMap = YES Then
            Exit Sub
        End If
    
        Dir = Val(Parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(Index, "Invalid Direction")
            Exit Sub
        End If
        
        Call SetPlayerDir(Index, Dir)
        Call SendDataToMapBut(Index, GetPlayerMap(Index), "PLAYERDIR" & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
        Exit Sub
        
    ' ::::::::::::::::::
    ' :: Custom Forms ::
    ' ::::::::::::::::::
    Case "customformbuttonclick"
        MyScript.ExecuteStatement "Scripts\Main.txt", "FormButtonClick " & Index & "," & Val(Parse(1)) & "," & Val(Parse(2))
        Exit Sub
    
    ' ::::::::::::::::::
    ' :: Custom Forms ::
    ' ::::::::::::::::::
    Case "customformtextclick"
        MyScript.ExecuteStatement "Scripts\Main.txt", "FormTextClick " & Index & "," & Val(Parse(1)) & "," & Val(Parse(2))
        Exit Sub
        
    ' :::::::::::::::::::::
    ' :: Use item packet ::
    ' :::::::::::::::::::::
    Case "useitem"
        InvNum = Val(Parse(1))
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
            n3 = Item(GetPlayerInvItemNum(Index, InvNum)).LuckReq
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
                            Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required Strength (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                            Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Defence (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLUCK(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your Luck is too low to equip this item!  Required Luck (" & n3 & ")", BrightRed)
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
                            Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required Strength (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                            Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Defence (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLUCK(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your Luck is too low to equip this item!  Required Luck (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerWeaponSlot(Index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                    
                Case ITEM_TYPE_LAMP
                    If InvNum <> GetPlayerWeaponSlot(Index) Then
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
                        ElseIf Int(GetPlayerLUCK(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your luck is too low to equip this item!  Required Luck (" & n3 & ")", BrightRed)
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
                            Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required Strength (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                            Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Defence (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLUCK(Index)) < n3 Then
                            Call PlayerMsg(Index, "Your luck is too low to equip this item!  Required Luck (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerShieldSlot(Index, InvNum)
                    Else
                        Call SetPlayerShieldSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
            
                Case ITEM_TYPE_POTIONADDHP
                    Call SetPlayerHP(Index, GetPlayerHP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendHP(Index)
                
                Case ITEM_TYPE_POTIONADDMP
                    Call SetPlayerMP(Index, GetPlayerMP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendMP(Index)
        
                Case ITEM_TYPE_POTIONADDSP
                    Call SetPlayerSP(Index, GetPlayerSP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendSP(Index)

                Case ITEM_TYPE_POTIONSUBHP
                    Call SetPlayerHP(Index, GetPlayerHP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendHP(Index)
                
                Case ITEM_TYPE_POTIONSUBMP
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendMP(Index)
        
                Case ITEM_TYPE_POTIONSUBSP
                    Call SetPlayerSP(Index, GetPlayerSP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendSP(Index)
                    
                Case ITEM_TYPE_KEY
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            If GetPlayerY(Index) > 0 Then
                                X = GetPlayerX(Index)
                                Y = GetPlayerY(Index) - 1
                            Else
                                Exit Sub
                            End If
                            
                        Case DIR_DOWN
                            If GetPlayerY(Index) < MAX_MAPY Then
                                X = GetPlayerX(Index)
                                Y = GetPlayerY(Index) + 1
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_LEFT
                            If GetPlayerX(Index) > 0 Then
                                X = GetPlayerX(Index) - 1
                                Y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_RIGHT
                            If GetPlayerX(Index) < MAX_MAPY Then
                                X = GetPlayerX(Index) + 1
                                Y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                    End Select
                    
                    ' Check if a key exists
                    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY Then
                        ' Check if the key they are using matches the map key
                        If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(X, Y).Data1 Then
                            TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
                            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            If Trim(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = "" Then
                                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", White)
                            Else
                                Call MapMsg(GetPlayerMap(Index), Trim(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
                            End If
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
                            
                            ' Check if we are supposed to take away the item
                            If Map(GetPlayerMap(Index)).Tile(X, Y).Data2 = 1 Then
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                Call PlayerMsg(Index, "The key disolves.", Yellow)
                            End If
                        End If
                    End If
                    
                    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_DOOR Then
                        TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        
                        Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
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
                            I = GetSpellReqLevel(Index, n)
                            If I <= GetPlayerLevel(Index) Then
                                I = FindOpenSpellSlot(Index)
                                
                                ' Make sure they have an open spell slot
                                If I > 0 Then
                                    ' Make sure they dont already have the spell
                                    If Not HasSpell(Index, n) Then
                                        Call SetPlayerSpell(Index, I, n)
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
                                Call PlayerMsg(Index, "You must be level " & I & " to learn this spell.", White)
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
        
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    Case "attack"
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & Index
        If GetPlayerWeaponSlot(Index) > 0 Then
            If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 > 0 Then
                If Arrows(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3).Ammo > 0 And GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index)) <> Arrows(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3).Ammo Then
                    If HasItem(Index, Arrows(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3).Ammo) Then
                        Call TakeItem(Index, Arrows(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3).Ammo, 1)
                        Call SendInventory(Index)
                        Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
                    Else
                        Call BattleMsg(Index, "You have run out of " & Trim(Item(Arrows(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3).Ammo).Name) & "!", 12, 0)
                    End If
                Else
                    Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
                End If
            End If
        End If
        
        ' Try to attack a player
        For I = 1 To MAX_PLAYERS
            ' Make sure we dont try to attack ourselves
            If I <> Index Then
                ' Can we attack the player?
                If CanAttackPlayer(Index, I) Then
                    If Not CanPlayerBlockHit(I) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(Index) Then
                            Damage = GetPlayerDamage(Index) - GetPlayerProtection(I)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            n = GetPlayerDamage(Index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(I)
                            Call BattleMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan, 0)
                            Call BattleMsg(I, GetPlayerName(Index) & " swings with enormous might!", BrightCyan, 1)
                            
                            'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                            'Call PlayerMsg(I, GetPlayerName(index) & " swings with enormous might!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(Index, I, Damage)
                        Else
                            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call BattleMsg(Index, GetPlayerName(I) & " blocked your hit!", BrightCyan, 0)
                        Call BattleMsg(I, "You blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan, 1)
                        
                        'Call PlayerMsg(index, GetPlayerName(I) & "'s " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked your hit!", BrightCyan)
                        'Call PlayerMsg(I, "Your " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                    End If
                    
                    Exit Sub
                End If
            End If
        Next I
        
        ' Try to attack a npc
        For I = 1 To MAX_MAP_NPCS
            ' Can we attack the npc?
            If CanAttackNpc(Index, I) Then
                ' Get the damage we can do
                If Not CanPlayerCriticalHit(Index) Then
                    Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), I).num).DEF / 2)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                Else
                    n = GetPlayerDamage(Index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), I).num).DEF / 2)
                    Call BattleMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan, 0)
                    
                    'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(Index, I, Damage)
                    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & I & SEP_CHAR & END_CHAR)
                Else
                    Call BattleMsg(Index, "Your attack does nothing.", BrightRed, 0)
                    
                    'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & I & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                End If
                Exit Sub
            End If
        Next I
        Exit Sub
    
    ' ::::::::::::::::::::::
    ' :: Use stats packet ::
    ' ::::::::::::::::::::::
    Case "usestatpoint"
        PointType = Val(Parse(1))
        
        ' Prevent hacking
        If (PointType < 0) Or (PointType > 3) Then
            Call HackingAttempt(Index, "Invalid Point Type")
            Exit Sub
        End If
                
        ' Make sure they have points
        If GetPlayerPOINTS(Index) > 0 Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & Index & "," & PointType
            Else
                Select Case PointType
                    Case 0
                        Call SetPlayerSTR(Index, GetPlayerSTR(Index) + 1)
                        Call BattleMsg(Index, "You have gained more strength!", 15, 0)
                    Case 1
                        Call SetPlayerDEF(Index, GetPlayerDEF(Index) + 1)
                        Call BattleMsg(Index, "You have gained more defense!", 15, 0)
                    Case 2
                        Call SetPlayerMAGI(Index, GetPlayerMAGI(Index) + 1)
                        Call BattleMsg(Index, "You have gained more magic abilities!", 15, 0)
                    Case 3
                        Call SetPlayerLUCK(Index, GetPlayerLUCK(Index) + 1)
                        Call BattleMsg(Index, "You have gained more luck!", 15, 0)
                End Select
                Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
            End If
        Else
            Call BattleMsg(Index, "You have no skill points to train with!", BrightRed, 0)
        End If
        
        Call SendHP(Index)
        Call SendMP(Index)
        Call SendSP(Index)
        Call SendStats(Index)
        
        Call SendDataTo(Index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & SEP_CHAR & END_CHAR)
        Exit Sub
        
    ' ::::::::::::::::::::::::::::::::
    ' :: Player info request packet ::
    ' ::::::::::::::::::::::::::::::::
    Case "playerinforequest"
        Name = Parse(1)
        
        I = FindPlayer(Name)
        If I > 0 Then
            Call PlayerMsg(Index, "Account: " & Trim(Player(I).Login) & ", Name: " & GetPlayerName(I), BrightGreen)
            If GetPlayerAccess(Index) > ADMIN_MONITER Then
                Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(I) & " -=-", BrightGreen)
                Call PlayerMsg(Index, "Level: " & GetPlayerLevel(I) & "  Exp: " & GetPlayerExp(I) & "/" & GetPlayerNextLevel(I), BrightGreen)
                Call PlayerMsg(Index, "HP: " & GetPlayerHP(I) & "/" & GetPlayerMaxHP(I) & "  MP: " & GetPlayerMP(I) & "/" & GetPlayerMaxMP(I) & "  SP: " & GetPlayerSP(I) & "/" & GetPlayerMaxSP(I), BrightGreen)
                Call PlayerMsg(Index, "STR: " & GetPlayerSTR(I) & "  DEF: " & GetPlayerDEF(I) & "  MAGI: " & GetPlayerMAGI(I) & "  LUCK: " & GetPlayerLUCK(I), BrightGreen)
                n = Int(GetPlayerSTR(I) / 2) + Int(GetPlayerLevel(I) / 2)
                I = Int(GetPlayerDEF(I) / 2) + Int(GetPlayerLevel(I) / 2)
                If n > 100 Then n = 100
                If I > 100 Then I = 100
                Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & I & "%", BrightGreen)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If
        Exit Sub
    
    ' :::::::::::::::::::::::
    ' :: Set sprite packet ::
    ' :::::::::::::::::::::::
    Case "setsprite"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The sprite
        n = Val(Parse(1))
        
        Call SetPlayerSprite(Index, n)
        Call SendPlayerData(Index)
        Exit Sub
  
    ' ::::::::::::::::::::::::::::::
    ' :: Set player sprite packet ::
    ' ::::::::::::::::::::::::::::::
    Case "setplayersprite"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The sprite
        I = FindPlayer(Parse(1))
        n = Val(Parse(2))
                
        Call SetPlayerSprite(I, n)
        Call SendPlayerData(I)
        Exit Sub
                
    ' ::::::::::::::::::::::::::
    ' :: Stats request packet ::
    ' ::::::::::::::::::::::::::
    Case "getstats"
        Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", White)
        Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
        Call PlayerMsg(Index, "HP: " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index) & "  MP: " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index) & "  SP: " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index), White)
        Call PlayerMsg(Index, "STR: " & GetPlayerSTR(Index) & "  DEF: " & GetPlayerDEF(Index) & "  MAGI: " & GetPlayerMAGI(Index) & "  LUCK: " & GetPlayerLUCK(Index), White)
        n = Int(GetPlayerSTR(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
        I = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
        If n > 100 Then n = 100
        If I > 100 Then I = 100
        Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & I & "%", White)
        Exit Sub
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player request for a new map ::
    ' ::::::::::::::::::::::::::::::::::
    Case "requestnewmap"
        Dir = Val(Parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(Index, "Invalid Direction")
            Exit Sub
        End If
                
        Call PlayerMove(Index, Dir, 1)
        Exit Sub
   
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    Case "mapdata"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        
        MapNum = GetPlayerMap(Index)
        Map(MapNum).Name = Parse(n + 1)
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        Map(MapNum).Moral = Val(Parse(n + 3))
        Map(MapNum).Up = Val(Parse(n + 4))
        Map(MapNum).Down = Val(Parse(n + 5))
        Map(MapNum).Left = Val(Parse(n + 6))
        Map(MapNum).Right = Val(Parse(n + 7))
        Map(MapNum).Music = Parse(n + 8)
        Map(MapNum).BootMap = Val(Parse(n + 9))
        Map(MapNum).BootX = Val(Parse(n + 10))
        Map(MapNum).BootY = Val(Parse(n + 11))
        Map(MapNum).Indoors = Val(Parse(n + 12))
        
        n = n + 13
        
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
            Map(MapNum).Tile(X, Y).Ground = Val(Parse(n))
            Map(MapNum).Tile(X, Y).Mask = Val(Parse(n + 1))
            Map(MapNum).Tile(X, Y).Anim = Val(Parse(n + 2))
            Map(MapNum).Tile(X, Y).Mask2 = Val(Parse(n + 3))
            Map(MapNum).Tile(X, Y).M2Anim = Val(Parse(n + 4))
            Map(MapNum).Tile(X, Y).Fringe = Val(Parse(n + 5))
            Map(MapNum).Tile(X, Y).FAnim = Val(Parse(n + 6))
            Map(MapNum).Tile(X, Y).Fringe2 = Val(Parse(n + 7))
            Map(MapNum).Tile(X, Y).F2Anim = Val(Parse(n + 8))
            Map(MapNum).Tile(X, Y).Type = Val(Parse(n + 9))
            Map(MapNum).Tile(X, Y).Data1 = Val(Parse(n + 10))
            Map(MapNum).Tile(X, Y).Data2 = Val(Parse(n + 11))
            Map(MapNum).Tile(X, Y).Data3 = Val(Parse(n + 12))
            Map(MapNum).Tile(X, Y).String1 = Parse(n + 13)
            Map(MapNum).Tile(X, Y).String2 = Parse(n + 14)
            Map(MapNum).Tile(X, Y).String3 = Parse(n + 15)
            Map(MapNum).Tile(X, Y).Light = Val(Parse(n + 16))
            Map(MapNum).Tile(X, Y).GroundSet = Val(Parse(n + 17))
            Map(MapNum).Tile(X, Y).MaskSet = Val(Parse(n + 18))
            Map(MapNum).Tile(X, Y).AnimSet = Val(Parse(n + 19))
            Map(MapNum).Tile(X, Y).Mask2Set = Val(Parse(n + 20))
            Map(MapNum).Tile(X, Y).M2AnimSet = Val(Parse(n + 21))
            Map(MapNum).Tile(X, Y).FringeSet = Val(Parse(n + 22))
            Map(MapNum).Tile(X, Y).FAnimSet = Val(Parse(n + 23))
            Map(MapNum).Tile(X, Y).Fringe2Set = Val(Parse(n + 24))
            Map(MapNum).Tile(X, Y).F2AnimSet = Val(Parse(n + 25))

            n = n + 26
            Next X
        Next Y
       
        For X = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(X) = Val(Parse(n))
            n = n + 1
            Call ClearMapNpc(X, MapNum)
        Next X
        
        ' Clear out it all
        For I = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), I).X, MapItem(GetPlayerMap(Index), I).Y)
            Call ClearMapItem(I, GetPlayerMap(Index))
        Next I
        
        ' Save the map
        Call SaveMap(MapNum)
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(Index))
        
        ' Respawn NPCS
        For I = 1 To MAX_MAP_NPCS
            Call SpawnNpc(I, GetPlayerMap(Index))
        Next I
        
        ' Refresh map for everyone online
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(I) & SEP_CHAR & Map(GetPlayerMap(I)).Revision & SEP_CHAR & END_CHAR)
                'Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
            End If
        Next I
    
        Exit Sub

    

    ' ::::::::::::::::::::::::::::
    ' :: Need map yes/no packet ::
    ' ::::::::::::::::::::::::::::
    Case "needmap"
        ' Get yes/no value
        s = LCase(Parse(1))
                
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
        
    Case "needmapnum2"
        Call SendMap(Index, GetPlayerMap(Index))
        Exit Sub
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to pick up something packet ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    Case "mapgetitem"
        Call PlayerMapGetItem(Index)
        Exit Sub
        
    ' ::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to drop something packet ::
    ' ::::::::::::::::::::::::::::::::::::::::::::
    Case "mapdropitem"
        InvNum = Val(Parse(1))
        Amount = Val(Parse(2))
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_INV Then
            Call HackingAttempt(Index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Amount <= 0 Then
                Call PlayerMsg(Index, "You must drop more than 0!", BrightRed)
                Exit Sub
            End If
            
            If Amount > GetPlayerInvItemValue(Index, InvNum) Then
                Call PlayerMsg(Index, "You dont have that much to drop!", BrightRed)
                Exit Sub
            End If
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(Index, InvNum)).Type <> ITEM_TYPE_CURRENCY Then
            If Amount > GetPlayerInvItemValue(Index, InvNum) Then
                Call HackingAttempt(Index, "Item amount modification")
                Exit Sub
            End If
        End If
        
        Call PlayerMapDropItem(Index, InvNum, Amount)
        Call SendStats(Index)
        Call SendHP(Index)
        Call SendMP(Index)
        Call SendSP(Index)
        Exit Sub
    
    ' ::::::::::::::::::::::::
    ' :: Respawn map packet ::
    ' ::::::::::::::::::::::::
    Case "maprespawn"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Clear out it all
        For I = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), I).X, MapItem(GetPlayerMap(Index), I).Y)
            Call ClearMapItem(I, GetPlayerMap(Index))
        Next I
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(Index))
        
        ' Respawn NPCS
        For I = 1 To MAX_MAP_NPCS
            Call SpawnNpc(I, GetPlayerMap(Index))
        Next I
        
        Call PlayerMsg(Index, "Map respawned.", Blue)
        Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
        Exit Sub
        
        
    ' ::::::::::::::::::::::::
    ' :: Kick player packet ::
    ' ::::::::::::::::::::::::
    Case "kickplayer"
        ' Prevent hacking
        If GetPlayerAccess(Index) <= 0 Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(Parse(1))
        
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
      
    ' :::::::::::::::::::::
    ' :: Ban list packet ::
    ' :::::::::::::::::::::
    Case "banlist"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        f = FreeFile
        Open App.Path & "\banlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            Input #f, Name
            
            Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, White)
            n = n + 1
        Loop
        Close #f
        Exit Sub
    
    ' ::::::::::::::::::::::::
    ' :: Ban destroy packet ::
    ' ::::::::::::::::::::::::
    Case "bandestroy"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call Kill(App.Path & "\banlist.txt")
        Call PlayerMsg(Index, "Ban list destroyed.", White)
        Exit Sub
        
    ' :::::::::::::::::::::::
    ' :: Ban player packet ::
    ' :::::::::::::::::::::::
    Case "banplayer"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(Parse(1))
        
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
    
    ' :::::::::::::::::::::::::::::
    ' :: Request edit map packet ::
    ' :::::::::::::::::::::::::::::
    Case "requesteditmap"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "EDITMAP" & SEP_CHAR & END_CHAR)
        Exit Sub
    
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit item packet ::
    ' ::::::::::::::::::::::::::::::
    Case "requestedititem"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub

    ' ::::::::::::::::::::::
    ' :: Edit item packet ::
    ' ::::::::::::::::::::::
    Case "edititem"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The item #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid Item Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing item #" & n & ".", ADMIN_LOG)
        Call SendEditItemTo(Index, n)
        Exit Sub
    
    ' ::::::::::::::::::::::
    ' :: Save item packet ::
    ' ::::::::::::::::::::::
    Case "saveitem"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid Item Index")
            Exit Sub
        End If
        
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).StrReq = Val(Parse(8))
        Item(n).DefReq = Val(Parse(9))
        Item(n).LuckReq = Val(Parse(10))
        Item(n).ClassReq = Val(Parse(11))
        Item(n).AccessReq = Val(Parse(12))
        
        Item(n).AddHP = Val(Parse(13))
        Item(n).AddMP = Val(Parse(14))
        Item(n).AddSP = Val(Parse(15))
        Item(n).AddStr = Val(Parse(16))
        Item(n).AddDef = Val(Parse(17))
        Item(n).AddMagi = Val(Parse(18))
        Item(n).AddLuck = Val(Parse(19))
        Item(n).AddEXP = Val(Parse(20))
        Item(n).Desc = Parse(21)
        Item(n).AttackSpeed = Val(Parse(22))
        
        ' Save it
        Call SendUpdateItemToAll(n)
        Call SaveItem(n)
        Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
        Exit Sub

    ' :::::::::::::::::::::
    ' :: Day/Night Stuff ::
    ' :::::::::::::::::::::

    Case "enabledaynight"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
    If TimeDisable = False Then
        Gamespeed = 0
        frmServer.GameTimeSpeed.Text = 0
        TimeDisable = True
        frmServer.Timer1.Enabled = False
        frmServer.Command69.Caption = "Enable Time"
    Else
        Gamespeed = 1
        frmServer.GameTimeSpeed.Text = 1
        TimeDisable = False
        frmServer.Timer1.Enabled = True
        frmServer.Command69.Caption = "Disable Time"
    End If
            
        Exit Sub
    
    Case "daynight"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        If Hours > 12 Then
            Hours = Hours - 12
        Else
            Hours = Hours + 12
        End If
            
        Exit Sub

    ' :::::::::::::::::::::::::::::
    ' :: Request edit npc packet ::
    ' :::::::::::::::::::::::::::::
    Case "requesteditnpc"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "NPCEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    
    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    Case "editnpc"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The npc #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(Index, "Invalid NPC Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing npc #" & n & ".", ADMIN_LOG)
        Call SendEditNpcTo(Index, n)
        Exit Sub
    
    ' :::::::::::::::::::::
    ' :: Save npc packet ::
    ' :::::::::::::::::::::
    Case "savenpc"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(Index, "Invalid NPC Index")
            Exit Sub
        End If
        
        ' Update the npc
        Npc(n).Name = Parse(2)
        Npc(n).AttackSay = Parse(3)
        Npc(n).Sprite = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).Range = Val(Parse(7))
        Npc(n).STR = Val(Parse(8))
        Npc(n).DEF = Val(Parse(9))
        Npc(n).Luck = Val(Parse(10))
        Npc(n).Magi = Val(Parse(11))
        Npc(n).Big = Val(Parse(12))
        Npc(n).MaxHp = Val(Parse(13))
        Npc(n).Exp = Val(Parse(14))
        Npc(n).SpawnTime = Val(Parse(15))
        
        z = 16
        For I = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(I).Chance = Val(Parse(z))
            Npc(n).ItemNPC(I).ItemNum = Val(Parse(z + 1))
            Npc(n).ItemNPC(I).ItemValue = Val(Parse(z + 2))
            z = z + 3
        Next I
        
        ' Save it
        Call SendUpdateNpcToAll(n)
        Call SaveNpc(n)
        Call AddLog(GetPlayerName(Index) & " saved npc #" & n & ".", ADMIN_LOG)
        Exit Sub
            
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit shop packet ::
    ' ::::::::::::::::::::::::::::::
    Case "requesteditshop"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "SHOPEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet ::
    ' ::::::::::::::::::::::
    Case "editshop"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The shop #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SHOPS Then
            Call HackingAttempt(Index, "Invalid Shop Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing shop #" & n & ".", ADMIN_LOG)
        Call SendEditShopTo(Index, n)
        Exit Sub
    
    ' ::::::::::::::::::::::
    ' :: Save shop packet ::
    ' ::::::::::::::::::::::
    Case "saveshop"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ShopNum = Val(Parse(1))
        
        ' Prevent hacking
        If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
            Call HackingAttempt(Index, "Invalid Shop Index")
            Exit Sub
        End If
        
        ' Update the shop
        Shop(ShopNum).Name = Parse(2)
        Shop(ShopNum).JoinSay = Parse(3)
        Shop(ShopNum).LeaveSay = Parse(4)
        Shop(ShopNum).FixesItems = Val(Parse(5))
        
        n = 6
        For z = 1 To 6
            For I = 1 To MAX_TRADES
                Shop(ShopNum).TradeItem(z).Value(I).GiveItem = Val(Parse(n))
                Shop(ShopNum).TradeItem(z).Value(I).GiveValue = Val(Parse(n + 1))
                Shop(ShopNum).TradeItem(z).Value(I).GetItem = Val(Parse(n + 2))
                Shop(ShopNum).TradeItem(z).Value(I).GetValue = Val(Parse(n + 3))
                n = n + 4
            Next I
        Next z
        
        ' Save it
        Call SendUpdateShopToAll(ShopNum)
        Call SaveShop(ShopNum)
        Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
        Exit Sub
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit spell packet ::
    ' :::::::::::::::::::::::::::::::
    Case "requesteditspell"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet ::
    ' :::::::::::::::::::::::
    Case "editspell"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(Index, "Invalid Spell Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing spell #" & n & ".", ADMIN_LOG)
        Call SendEditSpellTo(Index, n)
        Exit Sub
    
    ' :::::::::::::::::::::::
    ' :: Save spell packet ::
    ' :::::::::::::::::::::::
    Case "savespell"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(Index, "Invalid Spell Index")
            Exit Sub
        End If
        
        ' Update the spell
        Spell(n).Name = Parse(2)
        Spell(n).ClassReq = Val(Parse(3))
        Spell(n).LevelReq = Val(Parse(4))
        Spell(n).Type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
        Spell(n).MPCost = Val(Parse(9))
        Spell(n).Sound = Val(Parse(10))
        Spell(n).Range = Val(Parse(11))
        Spell(n).SpellAnim = Val(Parse(12))
        Spell(n).SpellTime = Val(Parse(13))
        Spell(n).SpellDone = Val(Parse(14))
        Spell(n).AE = Val(Parse(15))
                
        ' Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(Index) & " saving spell #" & n & ".", ADMIN_LOG)
        Exit Sub
    
    ' :::::::::::::::::::::::
    ' :: Set access packet ::
    ' :::::::::::::::::::::::
    Case "setaccess"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Call HackingAttempt(Index, "Trying to use powers not available")
            Exit Sub
        End If
        
        ' The index
        n = FindPlayer(Parse(1))
        ' The access
        I = Val(Parse(2))
        
        
        ' Check for invalid access level
        If I >= 0 Or I <= 3 Then
            If GetPlayerName(Index) <> GetPlayerName(n) Then
                If GetPlayerAccess(Index) > GetPlayerAccess(n) Then
                    ' Check if player is on
                    If n > 0 Then
                        If GetPlayerAccess(n) <= 0 Then
                            Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
                        End If
                    
                        Call SetPlayerAccess(n, I)
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
    
    Case "whosonline"
        Call SendWhosOnline(Index)
        Exit Sub

    Case "onlinelist"
        Call SendOnlineList
        Exit Sub

    Case "setmotd"
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse(1))
        Call GlobalMsg("MOTD changed to: " & Parse(1), BrightCyan)
        Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Parse(1), ADMIN_LOG)
        Exit Sub
    
    Case "traderequest"
        ' Trade num
        n = Val(Parse(1))
        z = Val(Parse(2))
        
        ' Prevent hacking
        If (n < 1) Or (n > 6) Then
            Call HackingAttempt(Index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Prevent hacking
        If (z <= 0) Or (z > (MAX_TRADES * 6)) Then
            Call HackingAttempt(Index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Index for shop
        I = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        
        ' Check if inv full
        If I <= 0 Then Exit Sub
        X = FindOpenInvSlot(Index, Shop(I).TradeItem(n).Value(z).GetItem)
        If X = 0 Then
            Call PlayerMsg(Index, "Trade unsuccessful, inventory full.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(Index, Shop(I).TradeItem(n).Value(z).GiveItem) >= Shop(I).TradeItem(n).Value(z).GiveValue Then
            Call TakeItem(Index, Shop(I).TradeItem(n).Value(z).GiveItem, Shop(I).TradeItem(n).Value(z).GiveValue)
            Call GiveItem(Index, Shop(I).TradeItem(n).Value(z).GetItem, Shop(I).TradeItem(n).Value(z).GetValue)
            Call PlayerMsg(Index, "The trade was successful!", Yellow)
        Else
            Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
        End If
        Exit Sub

    Case "fixitem"
        ' Inv num
        n = Val(Parse(1))
        
        ' Make sure its a equipable item
        If Item(GetPlayerInvItemNum(Index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, n)).Type > ITEM_TYPE_SHIELD Then
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
        I = Int(Item(GetPlayerInvItemNum(Index, n)).Data2 / 5)
        If I <= 0 Then I = 1
        
        DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, n)
        GoldNeeded = Int(DurNeeded * I / 2)
        If GoldNeeded <= 0 Then GoldNeeded = 1
        
        ' Check if they even need it repaired
        If DurNeeded <= 0 Then
            Call PlayerMsg(Index, "This item is in perfect condition!", White)
            Exit Sub
        End If
        
        ' Check if they have enough for at least one point
        If HasItem(Index, 1) >= I Then
            ' Check if they have enough for a total restoration
            If HasItem(Index, 1) >= GoldNeeded Then
                Call TakeItem(Index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(Index, n, Item(ItemNum).Data1)
                Call PlayerMsg(Index, "Item has been totally restored for " & GoldNeeded & " gold!", BrightBlue)
            Else
                ' They dont so restore as much as we can
                DurNeeded = (HasItem(Index, 1) / I)
                GoldNeeded = Int(DurNeeded * I / 2)
                If GoldNeeded <= 0 Then GoldNeeded = 1
                
                Call TakeItem(Index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(Index, n, GetPlayerInvItemDur(Index, n) + DurNeeded)
                Call PlayerMsg(Index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
            End If
        Else
            Call PlayerMsg(Index, "Insufficient gold to fix this item!", BrightRed)
        End If
        Exit Sub

    Case "search"
        X = Val(Parse(1))
        Y = Val(Parse(2))
        
        ' Prevent subscript out of range
        If X < 0 Or X > MAX_MAPX Or Y < 0 Or Y > MAX_MAPY Then
            Exit Sub
        End If
        
        ' Check for a player
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) And GetPlayerMap(Index) = GetPlayerMap(I) And GetPlayerX(I) = X And GetPlayerY(I) = Y Then
                
                ' Consider the player
                If GetPlayerLevel(I) >= GetPlayerLevel(Index) + 5 Then
                    Call PlayerMsg(Index, "You wouldn't stand a chance.", BrightRed)
                Else
                    If GetPlayerLevel(I) > GetPlayerLevel(Index) Then
                        Call PlayerMsg(Index, "This one seems to have an advantage over you.", Yellow)
                    Else
                        If GetPlayerLevel(I) = GetPlayerLevel(Index) Then
                            Call PlayerMsg(Index, "This would be an even fight.", White)
                        Else
                            If GetPlayerLevel(Index) >= GetPlayerLevel(I) + 5 Then
                                Call PlayerMsg(Index, "You could slaughter that player.", BrightBlue)
                            Else
                                If GetPlayerLevel(Index) > GetPlayerLevel(I) Then
                                    Call PlayerMsg(Index, "You would have an advantage over that player.", Yellow)
                                End If
                            End If
                        End If
                    End If
                End If
            
                ' Change target
                Player(Index).Target = I
                Player(Index).TargetType = TARGET_TYPE_PLAYER
                Call PlayerMsg(Index, "Your target is now " & GetPlayerName(I) & ".", Yellow)
                Exit Sub
            End If
        Next I
        
        ' Check for an npc
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(Index), I).num > 0 Then
                If MapNpc(GetPlayerMap(Index), I).X = X And MapNpc(GetPlayerMap(Index), I).Y = Y Then
                    ' Change target
                    Player(Index).Target = I
                    Player(Index).TargetType = TARGET_TYPE_NPC
                    Call PlayerMsg(Index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(Index), I).num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next I

        BX = X
        BY = Y
        
        ' Check for an item
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(Index), I).num > 0 Then
                If MapItem(GetPlayerMap(Index), I).X = X And MapItem(GetPlayerMap(Index), I).Y = Y Then
                    Call PlayerMsg(Index, "You see a " & Trim(Item(MapItem(GetPlayerMap(Index), I).num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next I
        Exit Sub
    
    Case "playerchat"
        n = FindPlayer(Parse(1))
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
        
        If Parse(1) = "" Then
            Call PlayerMsg(Index, "Click on the player you wish to chat to first.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Chat request has been sent to " & GetPlayerName(n) & ".", Pink)
        Call PlayerMsg(n, GetPlayerName(Index) & " wants you to chat with them.  Type /chat to accept, or /chatdecline to decline.", Pink)
    
        Player(n).ChatPlayer = Index
        Player(Index).ChatPlayer = n
        Exit Sub
    
    Case "achat"
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
    
    Case "dchat"
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

    Case "qchat"
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
    
    Case "sendchat"
        n = Player(Index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If

        Call SendDataTo(n, "sendchat" & SEP_CHAR & Parse(1) & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "pptrade"
        n = FindPlayer(Parse(1))
        
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
        Dim CanTrade As Boolean
        CanTrade = False
        
        If GetPlayerX(Index) = GetPlayerX(n) And GetPlayerY(Index) + 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(Index) = GetPlayerX(n) And GetPlayerY(Index) - 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(Index) + 1 = GetPlayerX(n) And GetPlayerY(Index) = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(Index) - 1 = GetPlayerX(n) And GetPlayerY(Index) = GetPlayerY(n) Then CanTrade = True
            
        If CanTrade = True Then
            ' Check to see if player is already in a trade
            If Player(n).InTrade = 1 Then
                Call PlayerMsg(Index, "Player is already in a trade!", Pink)
                Exit Sub
            End If
            
            Call PlayerMsg(Index, "Trade request has been sent to " & GetPlayerName(n) & ".", Pink)
            Call PlayerMsg(n, GetPlayerName(Index) & " wants you to trade with them.  Type /accept to accept, or /decline to decline.", Pink)
        
            Player(n).TradePlayer = Index
            Player(Index).TradePlayer = n
        Else
            Call PlayerMsg(Index, "You need to be beside the player to trade!", Pink)
            Call PlayerMsg(n, "The player needs to be beside you to trade!", Pink)
        End If
        Exit Sub

    Case "atrade"
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
        CanTrade = False
        
        If GetPlayerX(Index) = GetPlayerX(n) And GetPlayerY(Index) + 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(Index) = GetPlayerX(n) And GetPlayerY(Index) - 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(Index) + 1 = GetPlayerX(n) And GetPlayerY(Index) = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(Index) - 1 = GetPlayerX(n) And GetPlayerY(Index) = GetPlayerY(n) Then CanTrade = True
            
        If CanTrade = True Then
            Call PlayerMsg(Index, "You are trading with " & GetPlayerName(n) & "!", Pink)
            Call PlayerMsg(n, GetPlayerName(Index) & " accepted your trade request!", Pink)
            
            Call SendDataTo(Index, "PPTRADING" & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "PPTRADING" & SEP_CHAR & END_CHAR)
            
            For I = 1 To MAX_PLAYER_TRADES
                Player(Index).Trading(I).InvNum = 0
                Player(Index).Trading(I).InvName = ""
                Player(n).Trading(I).InvNum = 0
                Player(n).Trading(I).InvName = ""
            Next I
            
            Player(Index).InTrade = 1
            Player(Index).TradeItemMax = 0
            Player(Index).TradeItemMax2 = 0
            Player(n).InTrade = 1
            Player(n).TradeItemMax = 0
            Player(n).TradeItemMax2 = 0
        Else
            Call PlayerMsg(Index, "The player needs to be beside you to trade!", Pink)
            Call PlayerMsg(n, "You need to be beside the player to trade!", Pink)
        End If
        Exit Sub

    Case "qtrade"
        n = Player(Index).TradePlayer
        
        ' Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(Index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Stopped trading.", Pink)
        Call PlayerMsg(n, GetPlayerName(Index) & " stopped trading with you!", Pink)

        Player(Index).TradeOk = 0
        Player(n).TradeOk = 0
        Player(Index).TradePlayer = 0
        Player(Index).InTrade = 0
        Player(n).TradePlayer = 0
        Player(n).InTrade = 0
        Call SendDataTo(Index, "qtrade" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        Exit Sub

    Case "dtrade"
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

    Case "updatetradeinv"
        n = Val(Parse(1))
    
        Player(Index).Trading(n).InvNum = Val(Parse(2))
        Player(Index).Trading(n).InvName = Trim(Parse(3))
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
    
    Case "swapitems"
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

            For I = 1 To MAX_INV
                If Player(Index).TradeItemMax = Player(Index).TradeItemMax2 Then
                    Exit For
                End If
                If GetPlayerInvItemNum(n, I) < 1 Then
                    Player(Index).TradeItemMax2 = Player(Index).TradeItemMax2 + 1
                End If
            Next I

            For I = 1 To MAX_INV
                If Player(n).TradeItemMax = Player(n).TradeItemMax2 Then
                    Exit For
                End If
                If GetPlayerInvItemNum(Index, I) < 1 Then
                    Player(n).TradeItemMax2 = Player(n).TradeItemMax2 + 1
                End If
            Next I
            
            If Player(Index).TradeItemMax2 = Player(Index).TradeItemMax And Player(n).TradeItemMax2 = Player(n).TradeItemMax Then
                For I = 1 To MAX_PLAYER_TRADES
                    For X = 1 To MAX_INV
                        If GetPlayerInvItemNum(n, X) < 1 Then
                            If Player(Index).Trading(I).InvNum > 0 Then
                                Call GiveItem(n, GetPlayerInvItemNum(Index, Player(Index).Trading(I).InvNum), 1)
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, Player(Index).Trading(I).InvNum), 1)
                                Exit For
                            End If
                        End If
                    Next X
                Next I

                For I = 1 To MAX_PLAYER_TRADES
                    For X = 1 To MAX_INV
                        If GetPlayerInvItemNum(Index, X) < 1 Then
                            If Player(n).Trading(I).InvNum > 0 Then
                                Call GiveItem(Index, GetPlayerInvItemNum(n, Player(n).Trading(I).InvNum), 1)
                                Call TakeItem(n, GetPlayerInvItemNum(n, Player(n).Trading(I).InvNum), 1)
                                Exit For
                            End If
                        End If
                    Next X
                Next I

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
            Player(Index).TradeOk = 0
            Player(n).TradePlayer = 0
            Player(n).InTrade = 0
            Player(n).TradeOk = 0
            Call SendDataTo(Index, "qtrade" & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
        
    Case "party"
        n = FindPlayer(Parse(1))
       
        ' Prevent partying with self
        If n = Index Then
            Exit Sub
        End If
               
        ' Check for a full party and if so drop it
        Dim g As Integer
        g = 0
        If Player(Index).InParty = True Then
            For I = 1 To MAX_PARTY_MEMBERS
            If Player(Index).Party.Member(I) > 0 Then g = g + 1
            Next I
            If g > (MAX_PARTY_MEMBERS - 1) Then
            Call PlayerMsg(Index, "Party is full!", Pink)
            Exit Sub
            End If
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
           
            ' Check to see if player is already in a party
            If Player(n).InParty = False Then
                Call PlayerMsg(Index, GetPlayerName(n) & " has been invited to your party.", Pink)
                Call PlayerMsg(n, GetPlayerName(Index) & " has invited you to join their party.  Type /join to join, or /leave to decline.", Pink)
               
                Player(n).InvitedBy = Index
            Else
                Call PlayerMsg(Index, "Player is already in a party!", Pink)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If
        Exit Sub

    Case "joinparty"
        n = Player(Index).InvitedBy
       
        If n > 0 Then
            ' Check to make sure they aren't the starter
                ' Check to make sure that each of there party players match
                    Call PlayerMsg(Index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)
                 
                    If Player(n).InParty = False Then ' Set the party leader up
                    Call SetPMember(n, n) 'Make them the first member and make them the leader
                    Player(n).InParty = True 'Set them to be 'InParty' status
                    Call SetPShare(n, True)
                    End If
                   
                    Player(Index).InParty = True 'Player joined
                    Player(Index).Party.Leader = n 'Set party leader
                    Call SetPMember(n, Index) 'Add the member and update the party
                   
                    ' Make sure they are in right level range
                    If GetPlayerLevel(Index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(Index) - 5 > GetPlayerLevel(n) Then
                        Call PlayerMsg(Index, "There is more then a 5 level gap between you two, you will not share experience.", Pink)
                        Call PlayerMsg(n, "There is more then a 5 level gap between you two, you will not share experience.", Pink)
                        Call SetPShare(Index, False) 'Do not share experience with party
                    Else
                        Call SetPShare(Index, True) 'Share experience with party
                    End If
                   
                    For I = 1 To MAX_PARTY_MEMBERS
                        If Player(Index).Party.Member(I) > 0 And Player(Index).Party.Member(I) <> Index Then Call PlayerMsg(Player(Index).Party.Member(I), GetPlayerName(Index) & " has joined your party!", Pink)
                    Next I
                                       
        Else
            Call PlayerMsg(Index, "You have not been invited into a party!", Pink)
        End If
        Exit Sub

    Case "leaveparty"
        n = Player(Index).InvitedBy
       
        If n > 0 Or Player(Index).Party.Leader = Index Then
            If Player(Index).InParty = True Then
                Call PlayerMsg(Index, "You have left the party.", Pink)
                For I = 1 To MAX_PARTY_MEMBERS
                    If Player(Index).Party.Member(I) > 0 Then Call PlayerMsg(Player(Index).Party.Member(I), GetPlayerName(Index) & " has left the party.", Pink)
                Next I
               
                Call RemovePMember(Index) 'this handles removing them and updating the entire party
               
            Else
                Call PlayerMsg(Index, "Declined party request.", Pink)
                Call PlayerMsg(n, GetPlayerName(Index) & " declined your request.", Pink)
               
                Player(Index).InParty = False
                Player(Index).InvitedBy = 0
               
            End If
        Else
            Call PlayerMsg(Index, "You are not in a party!", Pink)
        End If
        Exit Sub
        
    Case "partychat"
        For I = 1 To MAX_PARTY_MEMBERS
            If Player(Index).Party.Member(I) > 0 Then Call PlayerMsg(Player(Index).Party.Member(I), Parse(1), Blue)
        Next I
        Exit Sub
    
    Case "spells"
        Call SendPlayerSpells(Index)
        Exit Sub
    
    Case "cast"
        n = Val(Parse(1))
        Call CastSpell(Index, n)
        Exit Sub

    Case "requestlocation"
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Map: " & GetPlayerMap(Index) & ", X: " & GetPlayerX(Index) & ", Y: " & GetPlayerY(Index), Pink)
        Exit Sub
    
    Case "refresh"
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & SEP_CHAR & END_CHAR)
        'Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Exit Sub
    
    Case "buysprite"
        ' Check if player stepped on sprite changing tile
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
            Call PlayerMsg(Index, "You need to be on a sprite tile to buy it!", BrightRed)
            Exit Sub
        End If
        
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 = 0 Then
            Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, I) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
                If Item(GetPlayerInvItemNum(Index, I)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(Index, I) >= Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 Then
                        Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3)
                        If GetPlayerInvItemValue(Index, I) <= 0 Then
                            Call SetPlayerInvItemNum(Index, I, 0)
                        End If
                        Call PlayerMsg(Index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                        Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(Index)
                    End If
                Else
                    If GetPlayerWeaponSlot(Index) <> I And GetPlayerArmorSlot(Index) <> I And GetPlayerShieldSlot(Index) <> I And GetPlayerHelmetSlot(Index) <> I Then
                        Call SetPlayerInvItemNum(Index, I, 0)
                        Call PlayerMsg(Index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                        Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(Index)
                    End If
                End If
                If GetPlayerWeaponSlot(Index) <> I And GetPlayerArmorSlot(Index) <> I And GetPlayerShieldSlot(Index) <> I And GetPlayerHelmetSlot(Index) <> I Then
                    Exit Sub
                End If
            End If
        Next I
        
        Call PlayerMsg(Index, "You dont have enough to buy this sprite!", BrightRed)
        Exit Sub
        
    Case "checkcommands"
        s = Parse(1)
        If Scripting = 1 Then
            PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & Index, Trim(s)
            MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & Index
        Else
            Call PlayerMsg(Index, "Thats not a valid command!", 12)
        End If
        Exit Sub
    
    Case "prompt"
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & Index & "," & Val(Parse(1)) & "," & Val(Parse(2))
        End If
        Exit Sub
                
    Case "requesteditarrow"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "arrowEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub

    Case "editarrow"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))
        
        If n < 0 Or n > MAX_ARROWS Then
            Call HackingAttempt(Index, "Invalid arrow Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing arrow #" & n & ".", ADMIN_LOG)
        Call SendEditArrowTo(Index, n)
        Exit Sub

    Case "savearrow"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid arrow Index")
            Exit Sub
        End If

        Arrows(n).Name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Arrows(n).Amount = Val(Parse(5))
        Arrows(n).Ammo = Val(Parse(6))

        Call SendUpdateArrowToAll(n)
        Call SaveArrow(n)
        Call AddLog(GetPlayerName(Index) & " saved arrow #" & n & ".", ADMIN_LOG)
        Exit Sub
    Case "checkarrows"
        n = Arrows(Val(Parse(1))).Pic
        
        Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "requesteditemoticon"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "EMOTICONEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub

    Case "editemoticon"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))
        
        If n < 0 Or n > MAX_EMOTICONS Then
            Call HackingAttempt(Index, "Invalid Emoticon Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing emoticon #" & n & ".", ADMIN_LOG)
        Call SendEditEmoticonTo(Index, n)
        Exit Sub

    Case "saveemoticon"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid Emoticon Index")
            Exit Sub
        End If

        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))

        Call SendUpdateEmoticonToAll(n)
        Call SaveEmoticon(n)
        Call AddLog(GetPlayerName(Index) & " saved emoticon #" & n & ".", ADMIN_LOG)
        Exit Sub
    
    Case "checkemoticons"
        n = Emoticons(Val(Parse(1))).Pic
        
        Call SendDataToMap(GetPlayerMap(Index), "checkemoticons" & SEP_CHAR & Index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub
    
    Case "mapreport"
        Packs = "mapreport" & SEP_CHAR
        For I = 1 To MAX_MAPS
            Packs = Packs & Map(I).Name & SEP_CHAR
        Next I
        Packs = Packs & END_CHAR
        
        Call SendDataTo(Index, Packs)
        Exit Sub
        
    Case "gmtime"
        GameTime = Val(Parse(1))
        Call SendTimeToAll
        Exit Sub
        
    Case "weather"
        GameWeather = Val(Parse(1))
        Call SendWeatherToAll
        Exit Sub
        
    Case "warpto"
        Call PlayerWarp(Index, Val(Parse(1)), GetPlayerX(Index), GetPlayerY(Index))
        Exit Sub
        
    Case "getnotes"
        Call SendDataTo(Index, "NOTES" & SEP_CHAR & CStr(GetVar("accounts\" & Trim(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "Notes")) & SEP_CHAR & END_CHAR)
        Exit Sub
    
    Case "notes"
        Call PutVar("accounts\" & Trim(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "Notes", CStr(Parse$(1)))
        Exit Sub
        
    Case "arrowhit"
        n = Val(Parse(1))
        z = Val(Parse(2))
        X = Val(Parse(3))
        Y = Val(Parse(4))
       
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
                            Call BattleMsg(Index, "You feel a surge of energy upon shooting!", BrightCyan, 0)
                            Call BattleMsg(z, GetPlayerName(Index) & " shoots With amazing accuracy!", BrightCyan, 1)
                           
                            'Call PlayerMsg(index, "You feel a surge of energy upon shooting!", BrightCyan)
                            'Call PlayerMsg(z, GetPlayerName(index) & " shoots With amazing accuracy!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If
                       
                        If Damage > 0 Then
                            Call AttackPlayer(Index, z, Damage)
                        Else
                            Call BattleMsg(Index, "Your attack does nothing.", BrightRed, 0)
                            Call BattleMsg(z, GetPlayerName(Index) & "'s attack did nothing.", BrightRed, 1)
                           
                            'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call BattleMsg(Index, GetPlayerName(z) & " blocked your hit!", BrightCyan, 0)
                        Call BattleMsg(z, "You blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan, 1)
                       
                        'Call PlayerMsg(index, GetPlayerName(z) & "'s " & Trim(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked your hit!", BrightCyan)
                        'Call PlayerMsg(z, "Your " & Trim(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
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
                    Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), z).num).DEF / 2)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                Else
                    n = GetPlayerDamage(Index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), z).num).DEF / 2)
                    Call BattleMsg(Index, "You feel a surge of energy upon shooting!", BrightCyan, 0)
                   
                    'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If
               
                If Damage > 0 Then
                    Call AttackNpc(Index, z, Damage)
                    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                Else
                    Call BattleMsg(Index, "Your attack does nothing.", BrightRed, 0)
                   
                    'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                End If
                Exit Sub
            End If
        End If
        Exit Sub
End Select

Call HackingAttempt(Index, "Packet modification")
Exit Sub
End Sub

Sub CloseSocket(ByVal Index As Long)
    ' Make sure player was/is playing the game, and if so, save'm.
    If Index > 0 Then
        Call LeftGame(Index)
    
        Call TextAdd(frmServer.txtText(0), "Connection from " & GetPlayerIP(Index) & " has been terminated.", True)
        
        frmServer.Socket(Index).Close
            
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If
End Sub

Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim n As Long, I As Long

    s = ""
    n = 0
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And I <> Index Then
            s = s & GetPlayerName(I) & ", "
            n = n + 1
        End If
    Next I
            
    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If
        
    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub SendOnlineList()
Dim Packet As String
Dim I As Long
Dim n As Long
Packet = ""
n = 0
For I = 1 To MAX_PLAYERS
    If IsPlaying(I) Then
        Packet = Packet & SEP_CHAR & GetPlayerName(I) & SEP_CHAR
        n = n + 1
    End If
Next I

Packet = "ONLINELIST" & SEP_CHAR & n & Packet & END_CHAR

Call SendDataToAll(Packet)
End Sub

Sub SendChars(ByVal Index As Long)
Dim Packet As String
Dim I As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For I = 1 To MAX_CHARS
        Packet = Packet & Trim(Player(Index).Char(I).Name) & SEP_CHAR & Trim(Class(Player(Index).Char(I).Class).Name) & SEP_CHAR & Player(Index).Char(I).Level & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendJoinMap(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = ""
    
    ' Send all players on current map to index
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And I <> Index And GetPlayerMap(I) = GetPlayerMap(Index) Then
            Packet = "PLAYERDATA" & SEP_CHAR
            Packet = Packet & I & SEP_CHAR
            Packet = Packet & GetPlayerName(I) & SEP_CHAR
            Packet = Packet & GetPlayerSprite(I) & SEP_CHAR
            Packet = Packet & GetPlayerMap(I) & SEP_CHAR
            Packet = Packet & GetPlayerX(I) & SEP_CHAR
            Packet = Packet & GetPlayerY(I) & SEP_CHAR
            Packet = Packet & GetPlayerDir(I) & SEP_CHAR
            Packet = Packet & GetPlayerAccess(I) & SEP_CHAR
            Packet = Packet & GetPlayerPK(I) & SEP_CHAR
            Packet = Packet & GetPlayerGuild(I) & SEP_CHAR
            Packet = Packet & GetPlayerGuildAccess(I) & SEP_CHAR
            Packet = Packet & GetPlayerClass(I) & SEP_CHAR
            Packet = Packet & GetPlayerColor(I) & SEP_CHAR
            Packet = Packet & GetPlayerGhost(I) & SEP_CHAR
            Packet = Packet & GetPlayerHP(I) & SEP_CHAR
            Packet = Packet & GetPlayerMaxHP(I) & SEP_CHAR
            Packet = Packet & GetPlayerLevel(I) & SEP_CHAR
            Packet = Packet & END_CHAR
            Call SendDataTo(Index, Packet)
        End If
    Next I
    
    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(Index) & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & GetPlayerColor(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGhost(Index) & SEP_CHAR
    Packet = Packet & GetPlayerHP(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMaxHP(Index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & GetPlayerColor(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGhost(Index) & SEP_CHAR
    Packet = Packet & GetPlayerHP(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMaxHP(Index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMapBut(Index, MapNum, Packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(Index) & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & GetPlayerColor(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGhost(Index) & SEP_CHAR
    Packet = Packet & GetPlayerHP(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMaxHP(Index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String, P1 As String, P2 As String
Dim X As Long
Dim Y As Long
Dim I As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
        With Map(MapNum).Tile(X, Y)
            Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .Light & SEP_CHAR
            Packet = Packet & .GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR
        End With
        Next X
    Next Y
    
    For X = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).Npc(X) & SEP_CHAR
    Next X
        
    Packet = Packet & END_CHAR
    
    X = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, X)
    P2 = Mid(Packet, X + 1, Len(Packet) - X)
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_ITEMS
        If MapNum > 0 Then
            Packet = Packet & MapItem(MapNum, I).num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).X & SEP_CHAR & MapItem(MapNum, I).Y & SEP_CHAR
        End If
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, I).num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).X & SEP_CHAR & MapItem(MapNum, I).Y & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_NPCS
        If MapNum > 0 Then
            Packet = Packet & MapNpc(MapNum, I).num & SEP_CHAR & MapNpc(MapNum, I).X & SEP_CHAR & MapNpc(MapNum, I).Y & SEP_CHAR & MapNpc(MapNum, I).Dir & SEP_CHAR
        End If
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, I).num & SEP_CHAR & MapNpc(MapNum, I).X & SEP_CHAR & MapNpc(MapNum, I).Y & SEP_CHAR & MapNpc(MapNum, I).Dir & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendItems(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    For I = 1 To MAX_ITEMS
        If Trim(Item(I).Name) <> "" Then
            Call SendUpdateItemTo(Index, I)
        End If
    Next I
End Sub

Sub SendEmoticons(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    For I = 0 To MAX_EMOTICONS
        If Trim(Emoticons(I).Command) <> "" Then
            Call SendUpdateEmoticonTo(Index, I)
        End If
    Next I
End Sub

Sub SendArrows(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    For I = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(Index, I)
    Next I
End Sub

Sub SendNpcs(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    For I = 1 To MAX_NPCS
        If Trim(Npc(I).Name) <> "" Then
            Call SendUpdateNpcTo(Index, I)
        End If
    Next I
End Sub

Sub SendInventory(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "PLAYERINV" & SEP_CHAR & Index & SEP_CHAR
    For I = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, I) & SEP_CHAR & GetPlayerInvItemValue(Index, I) & SEP_CHAR & GetPlayerInvItemDur(Index, I) & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendWornEquipment(ByVal Index As Long)
Dim Packet As String
    
    If IsPlaying(Index) Then
        Packet = "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If
End Sub

Sub SendHP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
    
    Packet = "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & SEP_CHAR & END_CHAR
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
    
    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerLUCK(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendClasses(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "CLASSESDATA" & SEP_CHAR & MAX_CLASSES & SEP_CHAR
    For I = 0 To MAX_CLASSES
        Packet = Packet & GetClassName(I) & SEP_CHAR & GetClassMaxHP(I) & SEP_CHAR & GetClassMaxMP(I) & SEP_CHAR & GetClassMaxSP(I) & SEP_CHAR & Class(I).STR & SEP_CHAR & Class(I).DEF & SEP_CHAR & Class(I).Luck & SEP_CHAR & Class(I).Magi & SEP_CHAR & Class(I).Locked & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & MAX_CLASSES & SEP_CHAR
    For I = 0 To MAX_CLASSES
        Packet = Packet & GetClassName(I) & SEP_CHAR & GetClassMaxHP(I) & SEP_CHAR & GetClassMaxMP(I) & SEP_CHAR & GetClassMaxSP(I) & SEP_CHAR & Class(I).STR & SEP_CHAR & Class(I).DEF & SEP_CHAR & Class(I).Luck & SEP_CHAR & Class(I).Magi & SEP_CHAR & Class(I).MaleSprite & SEP_CHAR & Class(I).FemaleSprite & SEP_CHAR & Class(I).Locked & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR

    Call SendDataTo(Index, Packet)
End Sub

Sub SendLeftGame(ByVal Index As Long)
Dim Packet As String
    
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
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
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToAllBut(Index, Packet)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).LuckReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddLuck & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).LuckReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddLuck & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).LuckReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddLuck & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditEmoticonTo(ByVal Index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & SEP_CHAR & Arrows(ItemNum).Ammo & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateArrowTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & SEP_CHAR & Arrows(ItemNum).Ammo & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditArrowTo(ByVal Index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITArrow" & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Arrows(EmoNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String
Dim I As Long

    'Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR
    'Packet = Packet & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Luck & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & END_CHAR
    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Luck & SEP_CHAR & Npc(NpcNum).Magi & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR
    For I = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(I).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(I).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(I).ItemValue & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendShops(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_SHOPS
        If Trim(Shop(I).Name) <> "" Then
            Call SendUpdateShopTo(Index, I)
        End If
    Next I
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim I As Long, z As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For I = 1 To 6
        For z = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(I).Value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(z).GetValue & SEP_CHAR
        Next z
    Next I
    Packet = Packet & END_CHAR

    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpells(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_SPELLS
        If Trim(Spell(I).Name) <> "" Then
            Call SendUpdateSpellTo(Index, I)
        End If
    Next I
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim I As Long, X As Long, Y As Long, z As Long, XX As Long

    z = 0
    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For I = 1 To 6
        For XX = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(I).Value(XX).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(XX).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(XX).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(XX).GetValue & SEP_CHAR
            
            ' Item #
            X = Shop(ShopNum).TradeItem(I).Value(XX).GetItem
            
            If Item(X).Type = ITEM_TYPE_SPELL Then
                ' Spell class requirement
                Y = Spell(Item(X).Data1).ClassReq
                
                If Y = 0 Then
                    Call PlayerMsg(Index, Trim(Item(X).Name) & " can be used by all classes.", Yellow)
                Else
                    Call PlayerMsg(Index, Trim(Item(X).Name) & " can only be used by a " & GetClassName(Y - 1) & ".", Yellow)
                End If
            End If
            If X < 1 Then
                z = z + 1
            End If
        Next XX
    Next I
    Packet = Packet & END_CHAR
    
    If z = (MAX_TRADES * 6) Then
        Call PlayerMsg(Index, "This shop has nothing to sell!", BrightRed)
    Else
        Call SendDataTo(Index, Packet)
    End If
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "SPELLS" & SEP_CHAR
    For I = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(Index, I) & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherTo(ByVal Index As Long)
Dim Packet As String
    If RainIntensity <= 0 Then RainIntensity = 1
    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & RainIntensity & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherToAll()
Dim I As Long
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
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendWeatherTo(I)
        End If
    Next I
End Sub

Sub SendGameClockTo(ByVal Index As Long)
Dim Packet As String

    Packet = "GAMECLOCK" & SEP_CHAR & GameClock & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
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
Dim Packet As String

    Packet = "NEWS" & SEP_CHAR & ReadINI("DATA", "ServerNews", App.Path & "\News.ini") & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTimeTo(ByVal Index As Long)
Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
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
Dim Packet As String

    Packet = "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & Index & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
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
Dim Packet As String

    Packet = "DTIME" & SEP_CHAR & TimeDisable & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub
